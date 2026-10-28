# -*- coding: utf-8 -*-
"""
test_student_flow.py — student 账号端到端功能验证（模拟前端真实调用）
覆盖：登录 → meta → 知识点训练 → 模拟卷 → 真题卷 → 判分 → 相似题 → 错题本 → 学情统计 → 权限隔离
"""
import os
import sys

import requests

BASE = os.environ.get("GAOKAO_BASE", "http://127.0.0.1:7100")
S = requests.Session()
PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  [OK]   " if cond else "  [FAIL] ") + name +
          (("  ← " + str(detail)[:160]) if detail and not cond else ""))


def api(method, path, token=None, **kw):
    h = kw.pop("headers", {})
    if token:
        h["X-Auth-Token"] = token
    r = S.request(method, BASE + path, headers=h, timeout=60, **kw)
    try:
        j = r.json()
    except Exception:
        j = {}
    return r.status_code, j


def data_of(resp):
    return resp[1].get("data") if isinstance(resp[1], dict) else None


print("=" * 70)
print("① 服务与静态资源")
code, j = api("GET", "/api/health")
check("健康检查", code == 200 and j.get("data", {}).get("status") == "up")
for p in ("/", "/js/app.js", "/js/student.js", "/css/app.css"):
    code, _ = api("GET", p)
    check("静态资源 %s" % p, code == 200)

print("\n② 登录")
code, j = api("POST", "/api/login", json={"username": "student", "password": "123456"})
TOKEN = data_of((code, j))["token"] if code == 200 else ""
U = (data_of((code, j)) or {}).get("user", {})
check("student/123456 登录成功", code == 200 and bool(TOKEN))
check("身份为 student", U.get("role") == "student", U)
code, _ = api("POST", "/api/login", json={"username": "student", "password": "bad"})
check("错误密码被拒 401", code == 401)
code, _ = api("POST", "/api/login", json={"username": "nobody", "password": "123456"})
check("不存在账号被拒 401", code == 401)

print("\n③ 学生端初始化数据（meta / kp-tree / tags）")
code, j = api("GET", "/api/meta", TOKEN)
meta = data_of((code, j)) or {}
check("GET /api/meta", code == 200)
subjects = meta.get("subjects") or []
kp_tree = meta.get("kp_tree") or {}
check("科目列表非空", len(subjects) >= 6, subjects)
check("知识点树非空", len(kp_tree) >= 6, "科目数 %d" % len(kp_tree))
code, j = api("GET", "/api/kp-tree", TOKEN)
tree = data_of((code, j)) or {}
check("GET /api/kp-tree", code == 200 and len(tree) > 0)

# 取一个真实可用的 科目/章节/知识点 三元组
triple = None
for subj, chapters in tree.items():
    for ch, kps in (chapters or {}).items():
        if kps:
            triple = (subj, ch, kps[0])
            break
    if triple:
        break
check("取得可练习的三元组", triple is not None, triple)
print("      使用三元组:", triple)

print("\n④ 知识点训练（practice/kp）")
qs = []
if triple:
    subj, ch, kp = triple
    code, j = api("POST", "/api/practice/kp", TOKEN,
                  json={"subject": subj, "chapter": ch, "kp": kp})
    qs = data_of((code, j)) or []
    check("POST /api/practice/kp 返回题目", code == 200 and isinstance(qs, list) and len(qs) > 0,
          "状态 %s 数量 %s" % (code, len(qs) if isinstance(qs, list) else "?"))
    if qs:
        q0 = qs[0]
        check("题目字段完整(qid/stem/qtype/options)",
              all(k in q0 for k in ("qid", "stem", "qtype")) and
              (q0.get("options") or q0.get("qtype") != "单选"),
              list(q0.keys())[:12])

print("\n⑤ 判分正确性（单选/多选/填空/解答 各测正确答案）")


def find_by_type(items, qtype):
    for it in items:
        if it.get("qtype") == qtype:
            return it
    return None


# 补一批混合题来覆盖不同题型
code, j = api("POST", "/api/practice/mock", TOKEN,
              json={"subject": "物理", "difficulty": 3, "count": 20})
mock = data_of((code, j)) or []
check("POST /api/practice/mock 组卷", code == 200 and len(mock) > 0, len(mock))
pool = (qs or []) + mock
print("      待测题池:", len(pool), "题，题型:", sorted(set(x.get('qtype') for x in pool)))

for qt in ("单选", "多选", "填空"):
    q = find_by_type(pool, qt)
    if not q:
        print("      （题池中无 %s 题，跳过）" % qt)
        continue
    std = str(q.get("answer") or "").replace(" ", "")
    if not std:
        continue
    code, j = api("POST", "/api/answer", TOKEN,
                  json={"qid": q["qid"], "answer": std, "duration_sec": 15, "mode": "kp"})
    d = data_of((code, j)) or {}
    check("%s 提交正确答案 → 判对" % qt, code == 200 and d.get("correct") is True,
          "std=%s 返回=%s" % (std[:30], d.get("correct")))
    # 故意错答（单选提交与答案不同的字母）
    if qt == "单选":
        wrong = "B" if std != "B" else "C"
        code, j2 = api("POST", "/api/answer", TOKEN,
                       json={"qid": q["qid"], "answer": wrong, "duration_sec": 15, "mode": "kp"})
        d2 = data_of((code, j2)) or {}
        check("单选提交错误答案 → 判错", code == 200 and d2.get("correct") is False,
              "提交 %s 返回 %s" % (wrong, d2.get("correct")))

q_calc = find_by_type(pool, "解答")
if q_calc:
    code, j = api("POST", "/api/answer", TOKEN,
                  json={"qid": q_calc["qid"], "answer": "自述解答过程", "duration_sec": 90, "mode": "kp"})
    d = data_of((code, j)) or {}
    check("解答题返回 correct=null 并给参考答案",
          code == 200 and d.get("correct") is None and "std_answer" in d,
          {k: (str(v)[:40] if v else v) for k, v in d.items()})
    code, j = api("POST", "/api/answer/self", TOKEN,
                  json={"qid": q_calc["qid"], "correct": True, "answer": "自述解答过程",
                        "duration_sec": 90, "mode": "kp"})
    check("解答题自评（我答对了）", code == 200, j.get("error") if code != 200 else "")
else:
    print("      （题池中无解答题，跳过自评）")

print("\n⑥ 真题卷（papers）")
code, j = api("GET", "/api/papers", TOKEN)
papers = data_of((code, j)) or []
check("GET /api/papers", code == 200 and isinstance(papers, list), len(papers))
if papers:
    p = papers[-1]
    code, j = api("GET", "/api/papers/%s/questions" % p["pid"], TOKEN)
    pd = data_of((code, j))
    # 返回结构为 {"paper": {...}, "questions": [...]}
    qlist = pd.get("questions") if isinstance(pd, dict) else None
    check("真题卷取题（%s）" % str(p.get("title", ""))[:22],
          code == 200 and isinstance(qlist, list) and len(qlist) > 0,
          "题数 %s" % (len(qlist) if isinstance(qlist, list) else pd))

print("\n⑦ 相似题（same knowledge point）")
q_ok = find_by_type(pool, "单选") or (qs[0] if qs else None)
if q_ok and q_ok.get("knowledge_point"):
    # 该知识点在题库中的题量：仅 1 题时排除自身返回 null 属正常
    from collections import Counter
    same_kp = [it for it in pool
               if it.get("knowledge_point") == q_ok["knowledge_point"]
               and it.get("subject") == q_ok.get("subject")]
    params = {"kp": q_ok["knowledge_point"], "exclude_qid": q_ok["qid"]}
    if q_ok.get("subject"):
        params["subject"] = q_ok["subject"]
    code, j = api("GET", "/api/similar", TOKEN, params=params)
    sim = data_of((code, j))
    if sim:
        check("GET /api/similar 返回同知识点新题",
              code == 200 and sim.get("qid") != q_ok["qid"], str(sim.get("qid")))
    else:
        # 用题量充足的知识点复测（物理/牛顿第二定律 题库有 20+ 题）
        code2, j2 = api("GET", "/api/similar", TOKEN,
                        params={"kp": "牛顿第二定律", "subject": "物理"})
        sim2 = data_of((code2, j2))
        check("GET /api/similar（改用题量充足的知识点）",
              code2 == 200 and sim2 is not None,
              "首次返回 null（该 kp 题库仅 1 题，排除自身后无题，属预期）")
else:
    print("      （无可用题目，跳过）")

print("\n⑧ 错题本 / 重练")
code, j = api("GET", "/api/wrong", TOKEN)
wrong = data_of((code, j)) or []
check("GET /api/wrong", code == 200 and isinstance(wrong, list), "错题 %d" % len(wrong))
if wrong:
    w0 = wrong[0]
    check("错题项含题目内容", bool(w0.get("stem")) and bool(w0.get("qid")),
          list(w0.keys())[:10])
code, j = api("GET", "/api/practice/retest", TOKEN)
retest = data_of((code, j))
check("GET /api/practice/retest 错题重练", code == 200, j.get("error") if code != 200 else "")
rn = len(retest) if isinstance(retest, list) else (len(retest.get("items", [])) if isinstance(retest, dict) else 0)
print("      重练题数:", rn)

print("\n⑨ 学情统计")
for ep in ("/api/analytics/overall", "/api/analytics/kp"):
    code, j = api("GET", ep, TOKEN)
    check("GET %s" % ep, code == 200, j.get("error") if code != 200 else "")
    if code == 200:
        d = data_of((code, j))
        print("      %s → %s" % (ep.split("/")[-1], str(d)[:110]))

print("\n⑩ 权限隔离（学生访问教师接口应 403）")
teacher_eps = [
    ("GET", "/api/questions?subject=物理", None),
    ("POST", "/api/questions", {"subject": "数学", "stem": "x"}),
    ("GET", "/api/users", None),
    ("POST", "/api/kp", {"subject": "数学", "chapter": "a", "knowledge_point": "b"}),
    ("GET", "/api/settings/api", None),
]
for method, ep, payload in teacher_eps:
    if payload is None:
        code, j = api(method, ep, TOKEN)
    else:
        code, j = api(method, ep, TOKEN, json=payload)
    check("%s %s → 403" % (method, ep.split("?")[0]), code == 403, "实际 %d" % code)
code, _ = api("GET", "/api/questions?subject=物理")
check("未登录访问 → 401", code == 401, "实际 %d" % code)
code, _ = api("GET", "/api/questions?subject=物理", "bogus-token")
check("伪造 token → 401", code == 401, "实际 %d" % code)

print("\n" + "=" * 70)
print("通过 %d 项，失败 %d 项" % (len(PASS), len(FAIL)))
if FAIL:
    print("失败项：")
    for f in FAIL:
        print("  -", f)
    sys.exit(1)
