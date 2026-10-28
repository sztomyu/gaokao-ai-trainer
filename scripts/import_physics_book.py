# -*- coding: utf-8 -*-
"""
import_physics_book.py — 把《2027广东高考物理考点精解.docx》的真题导入题库
文档结构（按段落顺序状态机解析）：
  Heading 2  第一模块　力学        → chapter
  Heading 3  考点 1　xxx（P01）    → knowledge_point
  Heading 4  真题精解               → 进入题目区
  Heading 5  1. 2006 广东 第 2 题（选择题） → 新题开始
  First Paragraph  题干
  Compact  A. / B. / C. / D. 选项
  First Paragraph  答案　C
  Body Text  解析
  Compact  解析条目

公式为 OMML 对象，按 XML 顺序拼接 w:t 与 m:t 还原文本；图片按 r:embed 导出到 media/。

用法：
  python scripts/import_physics_book.py --dry-run     # 只统计不写库
  python scripts/import_physics_book.py               # 写入题库
"""
import argparse
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "server"))
DOCX = os.path.join(BASE, "高考物理", "2027广东高考物理考点精解.docx")
MEDIA_DIR = os.path.join(BASE, "server", "data", "media")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# 考点标题：考点 1　运动的描述与匀变速直线运动（P01） / 考点 13　恒定电流（P15）
RE_KP = re.compile(r"^考点\s*(\d+)\s*[　 ]\s*(.+?)\s*[（(]\s*(P\d+)\s*[)）]\s*$")
# 题目标题：1. 2006 广东 第 2 题（选择题）
RE_Q = re.compile(r"^(\d+)\s*[.．]\s*(.+?)\s*[（(]\s*([^）)]+)\s*[)）]\s*$")
# 答案：答案　C / 答案　BD / 答案　（1）xxx
RE_ANS = re.compile(r"^答\s*案\s*[　 ]\s*(.*)$")
RE_OPT = re.compile(r"^([A-H])\s*[.．、]\s*(.*)$")


_MSUP = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵",
         "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻",
         "n": "ⁿ", "i": "ⁱ", "(": "⁽", ")": "⁾"}
_MSUB = {"0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅",
         "6": "₆", "7": "₇", "8": "₈", "9": "₉", "+": "₊", "-": "₋",
         "a": "ₐ", "e": "ₑ", "i": "ᵢ", "o": "ₒ", "x": "ₓ", "n": "ₙ",
         "k": "ₖ", "m": "ₘ"}


def _omml_children(el):
    if el is None:
        return ""
    return "".join(_omml_text(c) for c in el)


def _omml_text(el):
    """把 OMML 数学对象转成可读文本：分式 (a)/(b)、上标 ^ 或 Unicode、根式 √()。"""
    tag = el.tag
    if tag == "{%s}r" % M:                       # 数学 run
        return "".join(t.text or "" for t in el.findall("{%s}t" % M))
    if tag == "{%s}f" % M:                       # 分式
        num = _omml_children(el.find("{%s}num" % M))
        den = _omml_children(el.find("{%s}den" % M))
        if re.fullmatch(r"[\dA-Za-z]{1,3}", num) and re.fullmatch(r"[\dA-Za-z]{1,3}", den):
            return "%s/%s" % (num, den)
        return "(%s)/(%s)" % (num, den)
    if tag == "{%s}sSup" % M:                    # 上标
        base = _omml_children(el.find("{%s}e" % M))
        sup = _omml_children(el.find("{%s}sup" % M))
        if sup and all(c in _MSUP for c in sup):
            return base + "".join(_MSUP[c] for c in sup)
        return "%s^(%s)" % (base, sup)
    if tag == "{%s}sSub" % M:                    # 下标
        base = _omml_children(el.find("{%s}e" % M))
        sub = _omml_children(el.find("{%s}sub" % M))
        if sub and all(c in _MSUB for c in sub):
            return base + "".join(_MSUB[c] for c in sub)
        return "%s_%s" % (base, sub if len(sub) <= 3 else "(%s)" % sub)
    if tag == "{%s}sSubSup" % M:                 # 上下标
        base = _omml_children(el.find("{%s}e" % M))
        sub = _omml_children(el.find("{%s}sub" % M))
        sup = _omml_children(el.find("{%s}sup" % M))
        return "%s(%s)^(%s)" % (base, sub, sup)
    if tag == "{%s}rad" % M:                     # 根式
        body = _omml_children(el.find("{%s}e" % M))
        deg = _omml_children(el.find("{%s}deg" % M))
        return ("%s√(%s)" % (deg, body)) if deg.strip() else "√(%s)" % body
    if tag == "{%s}d" % M:                       # 定界符
        return "(" + _omml_children(el.find("{%s}e" % M)) + ")"
    return "".join(_omml_text(c) for c in el)


def _walk_para(node, out, img_saver):
    """按文档顺序遍历段落：w:t 文本 / OMML 公式 / 内嵌图片。"""
    tag = node.tag
    if tag in ("{%s}oMath" % M, "{%s}oMathPara" % M):
        out.append(_omml_text(node))
        return
    if tag == "{%s}t" % W:
        out.append(node.text or "")
        return
    if tag == "{%s}drawing" % W:
        blip = node.find(".//" + qn("a:blip"))
        if blip is not None:
            rid = blip.get(qn("r:embed"))
            fname = img_saver(rid) if rid else None
            if fname:
                out.append("[IMG:%s]" % fname)
        return
    for c in node:
        _walk_para(c, out, img_saver)


def _iter_para_text(p_el, doc, img_saver):
    """提取段落完整文本（含公式与配图标记）。"""
    out = []
    _walk_para(p_el, out, img_saver)
    s = "".join(out).replace("\u00a0", " ")
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


def parse(doc, save_images=True):
    os.makedirs(MEDIA_DIR, exist_ok=True)
    counter = {"img": 0}
    img_cache = {}

    def save_img(rid):
        if rid in img_cache:
            return img_cache[rid]
        try:
            part = doc.part.related_parts[rid]
        except Exception:
            return None
        ext = str(part.partname).rsplit(".", 1)[-1].lower()
        if ext not in ("png", "jpg", "jpeg", "gif", "bmp", "emf", "wmf"):
            ext = "png"
        counter["img"] += 1
        fname = "physbook_%03d.%s" % (counter["img"], ext)
        if save_images:
            with open(os.path.join(MEDIA_DIR, fname), "wb") as f:
                f.write(part.blob)
        img_cache[rid] = fname
        return fname

    questions = []
    cur_module, cur_kp, cur_kp_code = "", "", ""
    in_solution = False
    cur = None

    for p in doc.paragraphs:
        style = p.style.name
        txt = _iter_para_text(p._p, doc, save_img)
        if not txt:
            continue

        # 模块
        if style == "Heading 2":
            m = re.match(r"^第[一二三四五六七八九十]+模块\s*[　 ]\s*(.+)$", txt)
            if m:
                cur_module = m.group(1).strip()
            in_solution = False
            continue

        # 考点
        m = RE_KP.match(txt)
        if m and style.startswith("Heading"):
            cur_kp = m.group(2).strip()
            cur_kp_code = m.group(3)
            in_solution = False
            continue

        # 题目开始（Heading 5：1. 2006 广东 第 2 题（选择题））—— 必须先于通用 Heading 分支
        m = RE_Q.match(txt)
        if m and cur_kp:
            if cur:
                questions.append(cur)
            cur = {
                "no": m.group(1),
                "source": m.group(2).strip(),
                "qtype_raw": m.group(3).strip(),
                "kp": cur_kp, "kp_code": cur_kp_code, "module": cur_module,
                "stem_parts": [], "options": [], "answer": [],
                "analysis_parts": [], "figs": [],
            }
            in_solution = True   # 题目标题后进入题目内容
            continue

        # 其他 Heading（考纲要求/命题分析/常考模型/真题精解等）
        if style.startswith("Heading"):
            in_solution = "真题精解" in txt
            continue

        if cur is None:
            continue

        # 答案
        ma = RE_ANS.match(txt)
        if ma:
            cur["answer"].append(ma.group(1).strip())
            cur.setdefault("_phase", "answer")
            cur["_phase"] = "answer"
            continue

        # 选项
        mo = RE_OPT.match(txt)
        if mo and cur.get("_phase") != "analysis":
            cur["options"].append("%s. %s" % (mo.group(1), mo.group(2).strip()))
            continue

        # 解析标记
        if txt.strip() in ("解析", "【解析】", "详解", "【详解】"):
            cur["_phase"] = "analysis"
            continue

        # 其余按阶段归位
        phase = cur.get("_phase", "stem")
        if phase == "analysis":
            cur["analysis_parts"].append(txt)
        elif phase == "answer":
            # 答案之后、解析之前的零散文字并入解析
            cur["_phase"] = "analysis"
            cur["analysis_parts"].append(txt)
        else:
            cur["stem_parts"].append(txt)

    if cur:
        questions.append(cur)

    # 组装
    out = []
    for q in questions:
        stem = "\n".join(q["stem_parts"]).strip()
        answer = " ".join(q["answer"]).strip()
        analysis = "\n".join(q["analysis_parts"]).strip()
        figs = re.findall(r"\[IMG:([^\]]+)\]", stem) + re.findall(r"\[IMG:([^\]]+)\]", analysis)
        out.append({
            "subject": "物理",
            "chapter": q["module"] or "未分章",
            "knowledge_point": q["kp"],
            "qtype": _decide_qtype(q["qtype_raw"], answer, stem),
            "difficulty": _decide_difficulty(q["qtype_raw"], answer),
            "stem": stem,
            "options": q["options"],
            "answer": answer,
            "analysis": analysis,
            "source": q["source"],
            "tag": "高考真题",
            "media_paths": sorted(set(figs)),
            "media_type": "题图" if figs else "无",
        })
    return out


def _decide_qtype(raw, answer, stem=""):
    """题型判定：题干含填空线一律判填空（书中部分填空题错标为选择题）。"""
    if "计算" in raw:
        return "解答"
    if "实验" in raw:
        return "填空"
    if re.search(r"_{3,}|＿{3,}", stem):
        return "填空"
    if "选择" in raw:
        letters = re.sub(r"[^A-H]", "", answer.replace(" ", ""))
        return "多选" if len(letters) >= 2 else "单选"
    if "填空" in raw:
        return "填空"
    return "单选"


def _decide_difficulty(raw, answer):
    if "计算" in raw:
        return 4
    if "实验" in raw:
        return 3
    letters = re.sub(r"[^A-H]", "", answer.replace(" ", ""))
    return 4 if len(letters) >= 2 else 3


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    doc = Document(DOCX)
    items = parse(doc, save_images=not args.dry_run)

    n = len(items)
    no_ans = [it for it in items if not it["answer"].strip()]
    no_ana = [it for it in items if not it["analysis"].strip()]
    no_opt = [it for it in items if not it["options"] and it["qtype"] in ("单选", "多选")]
    with_img = [it for it in items if it["media_paths"]]
    kps = sorted(set((it["chapter"], it["knowledge_point"]) for it in items))

    print("=== 解析统计 ===")
    print("题目总数：%d" % n)
    print("模块 / 考点：%d 个模块，%d 个考点" % (len(set(i["chapter"] for i in items)), len(kps)))
    print("缺答案：%d　缺解析：%d　选择题缺选项：%d　含配图：%d" % (
        len(no_ans), len(no_ana), len(no_opt), len(with_img)))
    from collections import Counter
    print("题型分布：", dict(Counter(i["qtype"] for i in items)))
    print("平均题干长度：%.0f 字，平均解析长度：%.0f 字" % (
        sum(len(i["stem"]) for i in items) / max(1, n),
        sum(len(i["analysis"]) for i in items) / max(1, n)))
    print()
    print("=== 前 2 题样例 ===")
    for it in items[:2]:
        print("-" * 60)
        for k in ("chapter", "knowledge_point", "qtype", "source", "answer"):
            print("%s: %s" % (k, it[k]))
        print("stem: %s" % it["stem"][:200])
        print("options: %s" % it["options"])
        print("analysis: %s" % it["analysis"][:200])
        print("media: %s" % it["media_paths"])
    if no_ans:
        print()
        print("=== 缺答案样例 ===")
        for it in no_ans[:3]:
            print("[%s] %s | stem: %s" % (it["source"], it["knowledge_point"], it["stem"][:90]))

    if args.dry_run:
        print()
        print("（dry-run：未写入题库）")
        return

    import excel_store as store
    new_ids, skipped = store.add_questions(items)
    print()
    print("=== 写入题库 ===")
    print("新增 %d 题，去重跳过 %d 题" % (len(new_ids), skipped))
    if new_ids:
        print("首个 qid: %s  末个 qid: %s" % (new_ids[0], new_ids[-1]))


if __name__ == "__main__":
    main()
