# -*- coding: utf-8 -*-
"""
study_loop.py — 学习闭环组件
作答处理（记录+掌握度更新）· 相似题推送 · 过关判定 · 复测
闭环：作答 → 判分归因知识点 → 未掌握给解析 → 再出类似题 → 做对过关 → 遗忘曲线复测
"""
import random
import excel_store as store
import engine


def on_answer(username, question, user_answer, correct=None, duration_sec=0):
    """
    处理一次作答：判分 → 写作答记录 → 按知识点更新掌握度。
    correct 为 None 时自动判分（解答题需由调用方自评后传入）。
    返回结果字典。
    """
    if correct is None:
        correct = engine.grade(question, user_answer)
    qid = question["qid"]
    store.add_record(username, qid, user_answer, bool(correct), duration_sec)

    kp = str(question.get("knowledge_point", "")).split(",")[0].strip() or "未标注知识点"
    old, new, status = engine.update_mastery(username, kp, bool(correct),
                                             question.get("difficulty", 3))
    return {
        "correct": bool(correct),
        "kp": kp,
        "mastery_before": old,
        "mastery_after": new,
        "status": status,
        "passed": status == "已过关",
    }


def similar_question(username, kp, exclude_qids=None, subject=None):
    """再出类似题：优先选学生未答对过的同知识点题目。"""
    pool = engine.questions_for_kp(kp, exclude_qids, subject)
    if not pool:
        return None
    # 学生此前答对过的题目降权
    right_ids = {str(r["qid"]) for r in store.load_records(username=username, correct=True)}
    fresh = [q for q in pool if str(q["qid"]) not in right_ids]
    return random.choice(fresh or pool)


def retest_questions(username):
    """为所有到期知识点各抽 1 道复测题。返回 [(mastery_rec, question), ...]"""
    out = []
    for rec in engine.retest_due(username):
        kp = rec["knowledge_point"]
        q = similar_question(username, kp)
        if q:
            out.append((rec, q))
    return out


def apply_retest_result(username, kp, correct):
    return engine.apply_retest(username, kp, correct)
