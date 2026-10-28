# -*- coding: utf-8 -*-
"""
engine.py — 核心引擎
测评判分 · 掌握度模型 · 组卷算法（难度系数配比） · 复测调度（遗忘曲线）
"""
import datetime
import random
import excel_store as store

# 遗忘曲线复测间隔（天）：过关后第 1/2/4/7/15/30 天安排复测
RETEST_INTERVALS = [1, 2, 4, 7, 15, 30]

PASS_LINE = 80      # 掌握度 >= 80 判定过关
WEAK_LINE = 40      # 掌握度 < 40 判定未掌握


def _norm(s):
    return "".join(str(s).split()).upper().replace("，", ",").replace("；", ";")


def grade(question, user_answer):
    """
    判分。返回 True/False；解答题返回 None（由学生对照解析自评）。
    """
    qt = str(question.get("qtype", ""))
    if qt in ("单选", "多选"):
        return _norm(question.get("answer")) == _norm(user_answer)
    if qt == "填空":
        std = _norm(question.get("answer"))
        # 支持多个等价答案，用 | 分隔
        return any(_norm(user_answer) == a for a in std.split("|"))
    return None


def update_mastery(username, kp, correct, difficulty=3):
    """根据一次作答更新知识点掌握度，返回 (旧值, 新值, 状态)。"""
    rec = store.get_mastery(username, kp)
    old = float(rec["mastery"]) if rec and rec["mastery"] != "" else 0.0
    rc = int(rec["review_count"] or 0) if rec else 0
    last_review = rec.get("last_review", "") if rec else ""

    d = int(difficulty or 3)
    if correct:
        new = old + 6 + 2 * d          # 难题加分更多
    else:
        new = old - 8
    new = max(0.0, min(100.0, round(new, 1)))

    if new >= PASS_LINE:
        status = "已过关"
        last_review = str(datetime.date.today())
    elif new < WEAK_LINE:
        status = "未掌握"
    else:
        status = "学习中"

    store.upsert_mastery(username, kp, new, status,
                         last_review=last_review, review_count=rc)
    return old, new, status


def retest_due(username, only_producible=False):
    """返回需要复测的掌握度记录列表（已过关且到达遗忘曲线复测时间）。

    only_producible=True 时，只保留题库中仍能找到同知识点题目的记录，
    保证「待复测」数量与「一键复测」实际能出的题数一致。
    """
    due = []
    today = datetime.date.today()
    for r in store.get_mastery(username):
        if r.get("status") != "已过关" or not r.get("last_review"):
            continue
        try:
            last = datetime.date.fromisoformat(str(r["last_review"]))
        except ValueError:
            continue
        rc = int(r.get("review_count") or 0)
        interval = RETEST_INTERVALS[min(rc, len(RETEST_INTERVALS) - 1)]
        if (today - last).days >= interval:
            if only_producible and not questions_for_kp(r["knowledge_point"]):
                continue
            due.append(r)
    return due


def apply_retest(username, kp, correct):
    """复测结果处理：通过则 review_count+1 并顺延；不通过则打回学习闭环。"""
    rec = store.get_mastery(username, kp)
    if not rec:
        return None
    m = float(rec["mastery"] or 0)
    rc = int(rec["review_count"] or 0)
    if correct:
        m = min(100.0, m + 3)
        rc += 1
        status = "已过关"
        last = str(datetime.date.today())
    else:
        m = max(0.0, m - 15)
        status = "学习中"
        last = rec.get("last_review", "")
    store.upsert_mastery(username, kp, round(m, 1), status,
                         last_review=last, review_count=rc)
    return m, status


def generate_paper(subject, target_difficulty=3, count=10, kp=None):
    """
    按难度系数随机组卷：以目标难度为中心，按 25%/50%/25% 配比抽取
    (目标-1 / 目标 / 目标+1) 难度的题目，不足时用相邻难度补齐。
    """
    pool = store.load_questions(subject=subject, kp=kp)
    if not pool:
        return []
    t = max(1, min(5, int(round(float(target_difficulty)))))
    plan = [(t, 0.5), (t - 1, 0.25), (t + 1, 0.25)]
    random.shuffle(pool)
    by_diff = {}
    for q in pool:
        by_diff.setdefault(int(q["difficulty"] or 3), []).append(q)

    picked, used = [], set()
    for diff, ratio in plan:
        if diff < 1 or diff > 5:
            continue
        want = round(count * ratio)
        cand = [q for q in by_diff.get(diff, []) if q["qid"] not in used]
        for q in cand[:want]:
            picked.append(q)
            used.add(q["qid"])
    # 不足则从剩余池补齐
    if len(picked) < count:
        rest = [q for q in pool if q["qid"] not in used]
        random.shuffle(rest)
        for q in rest[: count - len(picked)]:
            picked.append(q)
            used.add(q["qid"])
    random.shuffle(picked)
    return picked[:count]


def questions_for_kp(kp, exclude_qids=None, subject=None):
    """取某知识点的题目（用于相似题推送/复测抽题）。"""
    exclude = {str(x) for x in (exclude_qids or [])}
    pool = store.load_questions(subject=subject, kp=kp)
    return [q for q in pool if str(q["qid"]) not in exclude]
