# -*- coding: utf-8 -*-
"""
analytics.py — 学情分析组件
掌握度画像 · 薄弱点清单 · 学习统计
"""
import excel_store as store
import engine


def kp_summary(username):
    """
    各知识点掌握情况：掌握度、状态、练题数、正确率。
    返回 list[dict]，按掌握度升序（最薄弱的排前面）。
    """
    records = store.load_records(username=username)
    stat = {}
    for r in records:
        q = store.get_question(r["qid"])
        if not q:
            continue
        kp = str(q.get("knowledge_point", "")).split(",")[0].strip() or "未标注知识点"
        s = stat.setdefault(kp, {"kp": kp, "total": 0, "right": 0,
                                 "subject": q.get("subject", "")})
        s["total"] += 1
        if int(r["correct"] or 0) == 1:
            s["right"] += 1

    mastery_map = {r["knowledge_point"]: r for r in store.get_mastery(username)}
    out = []
    for kp, s in stat.items():
        m = mastery_map.get(kp, {})
        out.append({
            "subject": s["subject"],
            "kp": kp,
            "mastery": float(m.get("mastery") or 0),
            "status": m.get("status", "未练习"),
            "total": s["total"],
            "accuracy": round(100.0 * s["right"] / s["total"], 1) if s["total"] else 0.0,
        })
    out.sort(key=lambda x: x["mastery"])
    return out


def weak_points(username, threshold=60):
    return [x for x in kp_summary(username) if x["mastery"] < threshold]


def overall(username):
    """总览统计。"""
    records = store.load_records(username=username)
    total = len(records)
    right = sum(1 for r in records if int(r["correct"] or 0) == 1)
    mastery_rows = store.get_mastery(username)
    passed = sum(1 for r in mastery_rows if r.get("status") == "已过关")
    due = len(engine.retest_due(username, only_producible=True))
    return {
        "total_answers": total,
        "accuracy": round(100.0 * right / total, 1) if total else 0.0,
        "kp_touched": len(mastery_rows),
        "kp_passed": passed,
        "retest_due": due,
    }


# ---------------- 看板：题库统计（各科 × 题型分类） ----------------
CATEGORIES = [("选择题", "单选/多选"), ("填空题", "填空"),
              ("阅读题", "解答含'阅读'"), ("大题", "解答其他")]

CAT_COLORS = {
    "选择题": "#3182CE",   # 蓝
    "填空题": "#38A169",   # 绿
    "阅读题": "#DD6B20",   # 橙
    "大题":   "#805AD5",   # 紫
}


def category_of(q):
    """根据题目 qtype + 知识点/章节文本判定所属大分类。"""
    qt = str(q.get("qtype", ""))
    if qt in ("单选", "多选"):
        return "选择题"
    if qt == "填空":
        return "填空题"
    if qt == "解答":
        text = (str(q.get("knowledge_point", "")) + " " +
                str(q.get("chapter", "")) + " " + str(q.get("stem", "")))
        if "阅读" in text or "完形" in text or "短文" in text or "cloze" in text.lower():
            return "阅读题"
        return "大题"
    return "其他"


SUBJECT_COLORS = {
    "语文": "#E53E3E", "数学": "#3182CE", "英语": "#38A169",
    "物理": "#DD6B20", "化学": "#805AD5", "生物": "#319795",
    "历史": "#B7791F", "地理": "#0987A0", "政治": "#D53F8C",
    "其他": "#718096",
}


def category_stats(subjects=None):
    """按科目 × 题型分类统计题库数量。
    返回 {'total': N, 'by_subject': [...], 'by_category': {cat: n}}。"""
    qs = store.load_questions()
    if subjects is None:
        subjects = list(store.SUBJECTS)
    stats = {s: {"subject": s, "total": 0,
                  "categories": {c: 0 for c, _ in CATEGORIES},
                  "tags": set()} for s in subjects}
    by_cat = {c: 0 for c, _ in CATEGORIES}
    other = 0
    for q in qs:
        sub = q.get("subject", "其他")
        if sub not in stats:
            stats[sub] = {"subject": sub, "total": 0,
                          "categories": {c: 0 for c, _ in CATEGORIES},
                          "tags": set()}
        cat = category_of(q)
        if cat in stats[sub]["categories"]:
            stats[sub]["categories"][cat] += 1
            by_cat[cat] += 1
        else:
            other += 1
        stats[sub]["total"] += 1
        tag = str(q.get("tag", "")).strip()
        if tag:
            stats[sub]["tags"].add(tag)
    ordered = [stats[s] for s in subjects if s in stats]
    others = [v for k, v in stats.items() if k not in subjects]
    ordered.extend(others)
    by_cat["其他"] = other
    return {"total": len(qs), "by_subject": ordered, "by_category": by_cat}
