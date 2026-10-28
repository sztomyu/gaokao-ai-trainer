# -*- coding: utf-8 -*-
"""
app.py — 高考学习训练系统 Web 版 Flask 后端
单进程托管：REST API + 静态前端(web/) + 题目图片媒体(server/data/media/)
启动：python server/app.py --host 127.0.0.1 --port 7100
"""
import os
import sys
import json
import argparse
import functools
import hashlib
import random
import re
import secrets
import threading
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
WEB_DIR = os.path.join(PROJECT_ROOT, "web")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from flask import Flask, request, jsonify, send_from_directory, abort
from werkzeug.exceptions import HTTPException

import excel_store as store
import engine
import study_loop
import analytics
import ai_recognize

store.init_all()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024     # 上传上限 60MB
app.json.ensure_ascii = False                            # 中文不转义，便于调试

WRITE_LOCK = threading.Lock()        # openpyxl 整表重写，所有写操作串行化
UPLOADS = {}                         # file_id -> {"path":..., "filename":...}
JOBS = {}                            # job_id  -> 识别任务状态（内存态，重启即失）


# ---------------- 基础工具 ----------------
def ok(data=None):
    return jsonify({"ok": True, "data": data})


def err(msg, code=400):
    return jsonify({"ok": False, "error": msg}), code


def body():
    return request.get_json(force=True, silent=True) or {}


def _clean(v):
    """把 Excel 读出的 datetime / 整数值 float 转成 JSON 友好形式。"""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def q_json(q):
    """题目 dict → 前端友好结构：options/media_paths 解析成数组，附 media_count。"""
    d = {k: _clean(v) for k, v in dict(q).items()}
    d["options"] = store.parse_options(q)
    mp = store.media_paths_of(q)
    d["media_paths"] = mp
    d["media_count"] = len(mp)
    try:
        d["difficulty"] = int(float(d.get("difficulty") or 3))
    except (TypeError, ValueError):
        d["difficulty"] = 3
    return d


# ---------------- 鉴权 ----------------
def _token_for(username, pwd_hash):
    return hashlib.sha256((username + ":" + pwd_hash).encode("utf-8")).hexdigest()


def current_user():
    token = request.headers.get("X-Auth-Token", "")
    if not token:
        return None
    for u in store.list_users():
        if str(u.get("status")) != "启用":
            continue
        if _token_for(str(u["username"]), str(u["password_hash"])) == token:
            return {"username": str(u["username"]), "role": u["role"], "name": u["name"]}
    return None


def require_auth(fn):
    @functools.wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            return err("未登录或登录已失效", 401)
        return fn(u, *a, **kw)
    return wrapper


def require_teacher(fn):
    @functools.wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            return err("未登录或登录已失效", 401)
        if u["role"] != "teacher":
            return err("需要教师权限", 403)
        return fn(u, *a, **kw)
    return wrapper


@app.errorhandler(Exception)
def handle_exc(e):
    if isinstance(e, HTTPException):
        if request.path.startswith("/api/"):
            return err(e.description or "请求错误", e.code)
        return e
    import traceback
    traceback.print_exc()
    if request.path.startswith("/api/"):
        return err("服务器内部错误：%s" % e, 500)
    return "Internal Server Error", 500


# ---------------- 静态资源 ----------------
@app.route("/")
def index():
    return send_from_directory(WEB_DIR, "index.html")


@app.route("/js/<path:fname>")
def js_files(fname):
    return send_from_directory(os.path.join(WEB_DIR, "js"), fname)


@app.route("/css/<path:fname>")
def css_files(fname):
    return send_from_directory(os.path.join(WEB_DIR, "css"), fname)


@app.route("/vendor/<path:fname>")
def vendor_files(fname):
    return send_from_directory(os.path.join(WEB_DIR, "vendor"), fname)


@app.route("/media/<path:fname>")
def media_files(fname):
    # 防目录穿越：拒绝 .. 与路径分隔符；允许中文文件名（docx 提取的图片可能含中文）
    if not fname or ".." in fname or "/" in fname or "\\" in fname:
        abort(404)
    return send_from_directory(store.media_dir(), fname)


# ---------------- 健康检查与登录 ----------------
@app.route("/api/health")
def health():
    return ok({"status": "up", "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")})


@app.route("/api/login", methods=["POST"])
def login():
    b = body()
    u = store.verify_login(str(b.get("username", "")).strip(), str(b.get("password", "")))
    if not u:
        return err("用户名或密码错误", 401)
    row = next((r for r in store.list_users() if str(r["username"]) == u["username"]), None)
    token = _token_for(u["username"], str(row["password_hash"]))
    return ok({"token": token, "user": u})


@app.route("/api/meta")
def meta():
    return ok({"subjects": store.SUBJECTS, "qtypes": store.QTYPES,
               "kp_tree": store.get_kp_tree()})


# ---------------- 题库（教师） ----------------
Q_EDIT_FIELDS = ("subject", "chapter", "knowledge_point", "qtype", "difficulty",
                 "stem", "options", "answer", "analysis", "source", "tag",
                 "media_type", "media_paths")


def _sanitize_q(b):
    it = {}
    for k in Q_EDIT_FIELDS:
        if k in b:
            v = b[k]
            if k == "difficulty":
                try:
                    v = max(1, min(5, int(float(v))))
                except (TypeError, ValueError):
                    v = 3
            it[k] = v
    return it


@app.route("/api/questions")
@require_teacher
def questions_list(user):
    a = request.args
    rows = store.load_questions(subject=a.get("subject") or None,
                                qtype=a.get("qtype") or None,
                                keyword=a.get("keyword") or None,
                                tag=a.get("tag") or None,
                                kp=a.get("kp") or None)
    total = len(rows)
    try:
        page = max(1, int(a.get("page", 1) or 1))
        size = max(1, min(200, int(a.get("size", 20) or 20)))
    except ValueError:
        page, size = 1, 20
    seg = rows[(page - 1) * size:(page - 1) * size + size]
    return ok({"total": total, "items": [q_json(r) for r in seg],
               "page": page, "size": size})


@app.route("/api/questions", methods=["POST"])
@require_teacher
def question_add(user):
    item = _sanitize_q(body())
    if not str(item.get("stem", "")).strip():
        return err("题干不能为空")
    with WRITE_LOCK:
        new_ids, skipped = store.add_questions([item])
    if not new_ids:
        return err("入库失败：与现有题目重复", 409)
    return ok({"qid": new_ids[0], "skipped": skipped})


@app.route("/api/questions/<qid>", methods=["PUT"])
@require_teacher
def question_update(user, qid):
    fields = _sanitize_q(body())
    if not fields:
        return err("没有需要更新的字段")
    # update_question 只自动序列化 options；media_paths 在此统一序列化
    if isinstance(fields.get("options"), list):
        fields["options"] = json.dumps(fields["options"], ensure_ascii=False)
    if isinstance(fields.get("media_paths"), list):
        fields["media_paths"] = json.dumps(fields["media_paths"], ensure_ascii=False)
    with WRITE_LOCK:
        done = store.update_question(qid, fields)
    if not done:
        return err("题目不存在", 404)
    return ok({"qid": qid})


@app.route("/api/questions/<qid>", methods=["DELETE"])
@require_teacher
def question_del(user, qid):
    with WRITE_LOCK:
        store.delete_question(qid)
    return ok({"qid": qid})


@app.route("/api/tags")
@require_auth
def tags(user):
    s = sorted({str(q.get("tag", "")).strip()
                for q in store.load_questions() if str(q.get("tag", "")).strip()})
    return ok(s)


# ---------------- 试卷导入（教师） ----------------
ALLOWED_EXT = {".docx", ".doc", ".pdf", ".png", ".jpg", ".jpeg",
               ".bmp", ".gif", ".webp", ".txt"}


@app.route("/api/import/upload", methods=["POST"])
@require_teacher
def import_upload(user):
    f = request.files.get("file")
    if not f or not f.filename:
        return err("未选择文件")
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return err("不支持的文件格式：%s（支持 docx / pdf / 图片 / txt）" % ext)
    fid = secrets.token_hex(8)
    path = os.path.join(UPLOAD_DIR, fid + ext)
    f.save(path)
    UPLOADS[fid] = {"path": path, "filename": f.filename}
    return ok({"file_id": fid, "filename": f.filename})


@app.route("/api/import/recognize", methods=["POST"])
@require_teacher
def import_recognize(user):
    b = body()
    fid = str(b.get("file_id", ""))
    hint = str(b.get("subject_hint", "")).strip()
    up = UPLOADS.get(fid)
    if not up or not os.path.exists(up["path"]):
        return err("文件不存在或已过期，请重新上传", 404)
    jid = secrets.token_hex(8)
    job = {"status": "running", "logs": [], "message": "", "items": None,
           "error": "", "filename": up["filename"]}
    JOBS[jid] = job

    def progress(msg):
        job["logs"].append(str(msg))

    def stream(items_chunk, idx, total):
        job["logs"].append("第 %d/%d 段识别完成，本段 %d 题" % (idx, total, len(items_chunk)))

    def work():
        try:
            items, message = ai_recognize.recognize(
                up["path"], subject_hint=hint,
                progress_cb=progress, stream_cb=stream)
            job["items"] = items
            job["message"] = message
            job["status"] = "done"
        except Exception as e:
            job["error"] = str(e)
            job["status"] = "error"

    threading.Thread(target=work, daemon=True).start()
    return ok({"job_id": jid})


@app.route("/api/import/jobs/<jid>")
@require_teacher
def import_job(user, jid):
    job = JOBS.get(jid)
    if not job:
        return err("识别任务不存在或已过期", 404)
    data = {"status": job["status"], "logs": job["logs"][-300:],
            "message": job["message"], "error": job["error"],
            "filename": job["filename"]}
    if job["status"] == "done":
        data["items"] = [{k: _clean(v) for k, v in dict(it).items()}
                         for it in (job["items"] or [])]
    return ok(data)


@app.route("/api/import/commit", methods=["POST"])
@require_teacher
def import_commit(user):
    items = body().get("items") or []
    if not isinstance(items, list) or not items:
        return err("没有可入库的题目")
    clean = [_sanitize_q(it) for it in items if isinstance(it, dict)]
    with WRITE_LOCK:
        new_ids, skipped = store.add_questions(clean)
    return ok({"new_qids": new_ids, "skipped": skipped})


# ---------------- 知识点 ----------------
@app.route("/api/kp-tree")
@require_auth
def kp_tree(user):
    return ok(store.get_kp_tree())


@app.route("/api/kp", methods=["POST"])
@require_teacher
def kp_add(user):
    b = body()
    s = str(b.get("subject", "")).strip()
    c = str(b.get("chapter", "")).strip() or "未分章"
    k = str(b.get("knowledge_point", "")).strip()
    if not s or not k:
        return err("科目与知识点不能为空")
    with WRITE_LOCK:
        store.add_knowledge_point(s, c, k)
    return ok({"subject": s, "chapter": c, "knowledge_point": k})


# ---------------- 试卷 ----------------
@app.route("/api/papers")
@require_auth
def papers(user):
    rows = store.list_papers(subject=request.args.get("subject") or None)
    out = []
    for p in rows:
        d = {k: _clean(v) for k, v in dict(p).items()}
        d["q_count"] = len([x for x in str(p.get("qids", "")).split(",") if x.strip()])
        out.append(d)
    return ok(out)


@app.route("/api/papers/<pid>/questions")
@require_auth
def paper_questions(user, pid):
    paper = next((p for p in store.list_papers() if str(p["pid"]) == str(pid)), None)
    if not paper:
        return err("试卷不存在", 404)
    qids = [x.strip() for x in str(paper.get("qids", "")).split(",") if x.strip()]
    qmap = {str(q["qid"]): q for q in store.get_questions_by_ids(qids)}
    ordered = [qmap[qid] for qid in qids if qid in qmap]
    return ok({"paper": {k: _clean(v) for k, v in dict(paper).items()},
               "questions": [q_json(q) for q in ordered]})


# ---------------- 学生练习 ----------------
@app.route("/api/practice/kp", methods=["POST"])
@require_auth
def practice_kp(user):
    b = body()
    rows = store.load_questions(subject=str(b.get("subject", "")).strip() or None,
                                chapter=str(b.get("chapter", "")).strip() or None,
                                kp=str(b.get("kp", "")).strip() or None)
    random.shuffle(rows)
    return ok([q_json(r) for r in rows])


@app.route("/api/practice/mock", methods=["POST"])
@require_auth
def practice_mock(user):
    b = body()
    subject = str(b.get("subject", "")).strip()
    try:
        difficulty = max(1, min(5, int(float(b.get("difficulty", 3)))))
        count = max(1, min(50, int(float(b.get("count", 10)))))
    except (TypeError, ValueError):
        return err("难度或题量参数错误")
    qs = engine.generate_paper(subject, difficulty, count)
    if not qs:
        return err("题库中该科目题目不足，无法组卷")
    return ok([q_json(q) for q in qs])


@app.route("/api/practice/retest")
@require_auth
def practice_retest(user):
    pairs = study_loop.retest_questions(user["username"])
    return ok([{"kp": rec["knowledge_point"],
                "mastery": _clean(rec.get("mastery")),
                "question": q_json(q)} for rec, q in pairs])


# ---------------- 作答 ----------------
def _mastery_snapshot(username, kp):
    rec = store.get_mastery(username, kp)
    m = float(rec["mastery"]) if rec and rec.get("mastery") not in ("", None) else 0.0
    st = rec.get("status", "未练习") if rec else "未练习"
    return m, st


def _retest_adjust(user, res, b):
    """复测模式：on_answer 之后再按遗忘曲线规则调整一次，并回读最新掌握度。"""
    kp_rt = str(b.get("kp") or res["kp"])
    study_loop.apply_retest_result(user["username"], kp_rt, res["correct"])
    rec2 = store.get_mastery(user["username"], kp_rt)
    if rec2:
        res["mastery_after"] = float(rec2["mastery"] or 0)
        res["status"] = rec2.get("status", res["status"])
        res["passed"] = res["status"] == "已过关"


@app.route("/api/answer", methods=["POST"])
@require_auth
def answer(user):
    b = body()
    q = store.get_question(str(b.get("qid", "")))
    if not q:
        return err("题目不存在", 404)
    ua = str(b.get("answer", ""))
    try:
        dur = max(0, int(float(b.get("duration_sec") or 0)))
    except (TypeError, ValueError):
        dur = 0
    mode = str(b.get("mode", ""))
    std_answer = str(q.get("answer", ""))
    analysis = str(q.get("analysis", ""))

    if str(q.get("qtype", "")) == "解答":
        # 解答题不判分不落记录：先返回答案与解析，学生自评后走 /api/answer/self
        kp = str(q.get("knowledge_point", "")).split(",")[0].strip() or "未标注知识点"
        m, st = _mastery_snapshot(user["username"], kp)
        return ok({"correct": None, "kp": kp,
                   "mastery_before": m, "mastery_after": m, "status": st,
                   "passed": st == "已过关",
                   "std_answer": std_answer, "analysis": analysis})

    with WRITE_LOCK:
        res = study_loop.on_answer(user["username"], q, ua, correct=None,
                                   duration_sec=dur)
        if mode == "retest":
            _retest_adjust(user, res, b)
    res["std_answer"] = std_answer
    res["analysis"] = analysis
    return ok(res)


@app.route("/api/answer/self", methods=["POST"])
@require_auth
def answer_self(user):
    b = body()
    q = store.get_question(str(b.get("qid", "")))
    if not q:
        return err("题目不存在", 404)
    correct = bool(b.get("correct"))
    ua = str(b.get("answer", ""))
    try:
        dur = max(0, int(float(b.get("duration_sec") or 0)))
    except (TypeError, ValueError):
        dur = 0
    mode = str(b.get("mode", ""))
    with WRITE_LOCK:
        res = study_loop.on_answer(user["username"], q, ua, correct=correct,
                                   duration_sec=dur)
        if mode == "retest":
            _retest_adjust(user, res, b)
    res["std_answer"] = str(q.get("answer", ""))
    res["analysis"] = str(q.get("analysis", ""))
    return ok(res)


# ---------------- 相似题 / 错题本 / 学情 ----------------
@app.route("/api/similar")
@require_auth
def similar(user):
    kp = request.args.get("kp", "").strip()
    excl = request.args.get("exclude_qid", "").strip()
    subject = request.args.get("subject", "").strip() or None
    if not kp:
        return err("缺少知识点参数")
    q = study_loop.similar_question(user["username"], kp,
                                    [excl] if excl else [], subject)
    return ok(q_json(q) if q else None)


@app.route("/api/wrong")
@require_auth
def wrong(user):
    recs = store.load_records(username=user["username"])
    latest = {}
    for r in recs:                      # 记录按时间顺序追加，后者覆盖前者
        latest[str(r["qid"])] = r
    out = []
    for qid in reversed(list(latest.keys())):
        r = latest[qid]
        if int(r["correct"] or 0) != 0:
            continue
        q = store.get_question(qid)
        if not q:
            continue                    # 题目已删除
        d = q_json(q)
        d["wrong_ts"] = str(r.get("ts", ""))
        d["my_answer"] = str(r.get("user_answer", ""))
        out.append(d)
    return ok(out)


@app.route("/api/analytics/overall")
@require_auth
def ana_overall(user):
    return ok(analytics.overall(user["username"]))


@app.route("/api/analytics/kp")
@require_auth
def ana_kp(user):
    return ok(analytics.kp_summary(user["username"]))


@app.route("/api/analytics/category")
@require_auth
def ana_category(user):
    s = analytics.category_stats()
    for b in s["by_subject"]:
        b["tags"] = sorted(b["tags"])   # set 不可 JSON 序列化
    return ok(s)


# ---------------- 账号管理（教师） ----------------
@app.route("/api/users")
@require_teacher
def users(user):
    return ok([{"username": str(r["username"]), "role": r["role"],
                "name": r["name"], "status": r["status"]}
               for r in store.list_users()])


@app.route("/api/users/save", methods=["POST"])
@require_teacher
def user_save(user):
    b = body()
    username = str(b.get("username", "")).strip()
    if not username:
        return err("用户名不能为空")
    role = b.get("role") if b.get("role") in ("teacher", "student") else "student"
    name = str(b.get("name", "")).strip() or username
    status = b.get("status") if b.get("status") in ("启用", "停用") else "启用"
    password = str(b.get("password", ""))
    with WRITE_LOCK:
        store.save_user(username, password, role, name, status)
    return ok({"username": username})


@app.route("/api/users/toggle", methods=["POST"])
@require_teacher
def user_toggle(user):
    username = str(body().get("username", "")).strip()
    row = next((r for r in store.list_users() if str(r["username"]) == username), None)
    if not row:
        return err("用户不存在", 404)
    if username == user["username"]:
        return err("不能停用自己的账号")
    new_status = "停用" if row["status"] == "启用" else "启用"
    with WRITE_LOCK:
        store.save_user(username, "", row["role"], row["name"], new_status)
    return ok({"username": username, "status": new_status})


# ---------------- API 设置（教师） ----------------
def _mask(key):
    if not key:
        return ""
    if len(key) <= 8:
        return key[:2] + "****"
    return key[:4] + "****" + key[-4:]


@app.route("/api/settings/api")
@require_teacher
def settings_get(user):
    cfg = ai_recognize.load_settings()
    out = {}
    for kind in ("text", "vision"):
        c = cfg[kind]
        out[kind] = {"api_base": c.get("api_base", ""), "path": c.get("path", ""),
                     "model": c.get("model", ""), "has_key": bool(c.get("api_key")),
                     "api_key_mask": _mask(c.get("api_key", ""))}
    return ok(out)


@app.route("/api/settings/api", methods=["POST"])
@require_teacher
def settings_save(user):
    b = body()
    cfg = ai_recognize.load_settings()
    new = {}
    for kind in ("text", "vision"):
        sub = b.get(kind) or {}
        cur = cfg[kind]
        key = str(sub.get("api_key", "")).strip()
        new[kind] = {
            "api_base": str(sub.get("api_base", cur.get("api_base", ""))).strip(),
            "path": str(sub.get("path", cur.get("path", ""))).strip(),
            "model": str(sub.get("model", cur.get("model", ""))).strip(),
            "api_key": key if key else cur.get("api_key", ""),   # 留空 = 不修改
        }
    with WRITE_LOCK:
        ai_recognize.save_settings(new["text"], new["vision"])
    return ok({"saved": True})


@app.route("/api/settings/test", methods=["POST"])
@require_teacher
def settings_test(user):
    kind = body().get("kind")
    if kind not in ("text", "vision"):
        return err("kind 必须是 text 或 vision")
    success, msg = ai_recognize.test_api(kind)
    return ok({"success": success, "message": msg})


# ---------------- 启动 ----------------
def main():
    ap = argparse.ArgumentParser(description="高考学习训练系统 Web 版")
    ap.add_argument("--port", "-p", type=int,
                    default=int(os.environ.get("PORT", 7100)))
    ap.add_argument("--host", default=os.environ.get("HOST", "127.0.0.1"))
    args = ap.parse_args()
    print("=" * 56)
    print("  高考学习训练系统 Web 版")
    print("  访问地址: http://%s:%d" % (args.host, args.port))
    print("  内置账号: teacher/123456（教师）  student/123456（学生）")
    print("=" * 56)
    app.run(host=args.host, port=args.port, threaded=True)


if __name__ == "__main__":
    main()
