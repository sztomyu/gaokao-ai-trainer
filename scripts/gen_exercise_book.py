# -*- coding: utf-8 -*-
"""
gen_exercise_book.py — 题库习题集 Word 生成器
从 server/data/questions.xlsx 读取指定科目的全部题目，按章节分组，
逐题排版（题干 + 选项 + 知识点 + 答案 + 解析），生成可打印 A4 习题集 docx。

用法：
  python gen_exercise_book.py --subject 数学 --output "数学高考练习.docx"
  python gen_exercise_book.py --subject 物理 --output "物理高考练习.docx"
"""
import argparse
import datetime
import json
import os
import re
import sys
from collections import Counter, OrderedDict

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE, "server", "data", "questions.xlsx")
MEDIA_DIR = os.path.join(BASE, "server", "data", "media")

# ---------------- LaTeX → Unicode 务实转换 ----------------

_SUP = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵",
        "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻",
        "=": "⁼", "n": "ⁿ", "i": "ⁱ", "(": "⁽", ")": "⁾"}
_SUB = {"0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅",
        "6": "₆", "7": "₇", "8": "₈", "9": "₉", "+": "₊", "-": "₋",
        "a": "ₐ", "e": "ₑ", "i": "ᵢ", "o": "ₒ", "x": "ₓ", "n": "ₙ",
        "k": "ₖ", "m": "ₘ"}

_SYMBOLS = {
    r"\times": "×", r"\cdot": "·", r"\div": "÷", r"\pm": "±", r"\mp": "∓",
    r"\leqslant": "≤", r"\leq": "≤", r"\le": "≤",
    r"\geqslant": "≥", r"\geq": "≥", r"\ge": "≥",
    r"\neq": "≠", r"\ne": "≠", r"\approx": "≈", r"\equiv": "≡",
    r"\infty": "∞", r"\propto": "∝",
    r"\in": "∈", r"\notin": "∉", r"\ni": "∋",
    r"\subset": "⊂", r"\subseteq": "⊆", r"\supset": "⊃",
    r"\cup": "∪", r"\cap": "∩", r"\emptyset": "∅", r"\varnothing": "∅",
    r"\forall": "∀", r"\exists": "∃",
    r"\rightarrow": "→", r"\to": "→", r"\Rightarrow": "⇒",
    r"\leftarrow": "←", r"\Leftarrow": "⇐", r"\leftrightarrow": "↔",
    r"\Leftrightarrow": "⇔", r"\longrightarrow": "⟶",
    r"\triangle": "△", r"\angle": "∠", r"\perp": "⊥", r"\parallel": "∥",
    r"\because": "∵", r"\therefore": "∴",
    r"\partial": "∂", r"\nabla": "∇",
    r"\sum": "Σ", r"\prod": "∏", r"\int": "∫",
    r"\degree": "°", r"\circ": "°", r"\prime": "′",
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\Delta": "Δ", r"\epsilon": "ε", r"\varepsilon": "ε",
    r"\zeta": "ζ", r"\eta": "η", r"\theta": "θ", r"\lambda": "λ",
    r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ", r"\pi": "π", r"\rho": "ρ",
    r"\sigma": "σ", r"\Sigma": "Σ", r"\tau": "τ", r"\phi": "φ",
    r"\varphi": "φ", r"\Phi": "Φ", r"\chi": "χ", r"\psi": "ψ",
    r"\omega": "ω", r"\Omega": "Ω", r"\lambda": "λ",
    r"\Omega": "Ω", r"\ell": "ℓ", r"\hbar": "ℏ",
    r"\ast": "∗", r"\star": "⋆", r"\bullet": "•",
    r"\vee": "∨", r"\wedge": "∧", r"\neg": "¬",
}

_FUNCS = [r"\sin", r"\cos", r"\tan", r"\cot", r"\sec", r"\csc",
          r"\log", r"\ln", r"\lim", r"\max", r"\min", r"\arcsin",
          r"\arccos", r"\arctan", r"\exp"]


def _grab_group(s, i):
    """从 s[i]=='{' 开始，返回 (花括号内容, '}' 之后的位置)。"""
    depth, j = 0, i
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
    return s[i + 1:], len(s)


def latex_to_unicode(text):
    """把 AI 题库里常见的 LaTeX 片段转换为可读 Unicode 文本（近似渲染）。"""
    if not text:
        return text
    s = text
    # 1) 去定界符
    s = re.sub(r"\$\$([\s\S]+?)\$\$", r"\1", s)
    s = re.sub(r"\\\[([\s\S]+?)\\\]", r"\1", s)
    s = re.sub(r"\\\(([\s\S]+?)\\\)", r"\1", s)
    s = re.sub(r"\$([^$\n]+?)\$", r"\1", s)
    s = s.replace(r"\left", "").replace(r"\right", "")
    s = s.replace(r"\!", "").replace(r"\,", " ").replace(r"\;", " ").replace(r"\ ", " ")

    # 2) \sqrt
    def _sqrt(m):
        idx = m.group(1)
        body, _ = _grab_group(s, m.start(2)) if False else (m.group(3), None)
        if idx:
            return "%s√(%s)" % (idx, body)
        return "√(%s)" % body
    prev = None
    while prev != s:
        prev = s
        m = re.search(r"\\sqrt\[(\w+)\]\s*(\{)", s)
        if m:
            body, end = _grab_group(s, m.start(2))
            s = s[:m.start()] + m.group(1) + "√(" + body + ")" + s[end:]
            continue
        m = re.search(r"\\sqrt\s*(\{)", s)
        if m:
            body, end = _grab_group(s, m.start(1))
            s = s[:m.start()] + "√(" + body + ")" + s[end:]
            continue
        m = re.search(r"\\sqrt\s*(\w)", s)
        if m:
            s = s[:m.start()] + "√" + m.group(1) + s[m.end():]

    # 3) \frac{a}{b} → (a)/(b)，再简化纯数字/单字母
    prev = None
    while prev != s:
        prev = s
        m = re.search(r"\\frac\s*(\{)", s)
        if not m:
            break
        num, after_num = _grab_group(s, m.start(1))
        # 跳过空白找分母
        k = after_num
        while k < len(s) and s[k] in " \t":
            k += 1
        if k < len(s) and s[k] == "{":
            den, after_den = _grab_group(s, k)
        else:
            den, after_den = (s[k] if k < len(s) else ""), k + 1
        s = s[:m.start()] + "(%s)/(%s)" % (num, den) + s[after_den:]
    s = re.sub(r"\((\d+)\)/\((\d+)\)", r"\1/\2", s)
    s = re.sub(r"\(([A-Za-z])\)/\((\d+)\)", r"\1/\2", s)
    s = re.sub(r"\((\d+)\)/\(([A-Za-z])\)", r"\1/\2", s)

    # 4) 符号替换（key 按长度降序，避免 \to 吃掉 \times 前缀）
    for key in sorted(_SYMBOLS, key=len, reverse=True):
        s = s.replace(key, _SYMBOLS[key])

    # 5) 函数名
    for f in _FUNCS:
        s = s.replace(f, f[1:])

    # 6) 上下标
    def _sup(m):
        t = m.group(1) if m.group(1) is not None else m.group(2)
        if t and all(c in _SUP for c in t):
            return "".join(_SUP[c] for c in t)
        return "^(" + (t or "") + ")"

    def _sub(m):
        t = m.group(1) if m.group(1) is not None else m.group(2)
        t_compact = (t or "").replace(" ", "")
        if t_compact and all(c in _SUB for c in t_compact):
            return "".join(_SUB[c] for c in t_compact)
        return "_(" + t_compact + ")"

    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\^\{([^{}]*)\}|\^([A-Za-z0-9+\-])", _sup, s, count=1)
        s = re.sub(r"_\{([^{}]*)\}|_([A-Za-z0-9+\-])", _sub, s, count=1)

    # 7) \vec{F} → F⃗（退化为删除 vec）
    s = re.sub(r"\\vec\s*", "", s)
    s = s.replace(r"\overline", "")

    # 8) 清理残留花括号与未知命令
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\[A-Za-z]+", "", s)     # 未识别命令丢弃
    s = re.sub(r"\\\\", "；", s)           # 换行
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


# ---------------- Word 工具 ----------------

def set_cn_font(style_or_run, cn="宋体", en="Times New Roman"):
    f = style_or_run.font
    f.name = en
    rpr = style_or_run._element.get_or_add_rPr() if hasattr(style_or_run, "_element") else None
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), cn)


def add_page_number(footer_para):
    run = footer_para.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1); run._r.append(it); run._r.append(f2)
    set_cn_font(run)


def para(doc, text="", size=10.5, bold=False, color=None, align=None,
         space_after=4, indent=None, cn="宋体", line=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    if color:
        r.font.color.rgb = RGBColor(*color)
    set_cn_font(r, cn=cn)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if line:
        p.paragraph_format.line_spacing = line
    if indent:
        p.paragraph_format.first_line_indent = Pt(indent)
    return p, r


def add_stem_with_images(doc, stem):
    """题干文本；[IMG:xxx] 标记替换为内嵌图片（文件存在时），否则剔除标记。"""
    parts = re.split(r"\[IMG:([^\]\s]+)\]", stem)
    # parts: [text, fname, text, fname, ...]
    for i, part in enumerate(parts):
        if i % 2 == 0:
            txt = latex_to_unicode(part)
            if txt.strip():
                p = doc.add_paragraph()
                r = p.add_run(txt)
                r.font.size = Pt(11)
                set_cn_font(r)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.25
        else:
            fpath = os.path.join(MEDIA_DIR, part)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists(fpath):
                try:
                    p.add_run().add_picture(fpath, width=Cm(6.5))
                    p.paragraph_format.space_after = Pt(4)
                    continue
                except Exception:
                    pass
            r = p.add_run("（本题原配图缺失，请结合题干文字理解）")
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            set_cn_font(r)


# ---------------- 主生成逻辑 ----------------

_EMBED_TAGS = ("【答案】", "【详解】", "【分析】", "【解析】", "【解答】")


def _split_stem_embedded_answer(it):
    """部分题目 stem 内嵌了「【答案】…【详解】…」（AI 识别时混入）。
    截断：正文只留题干；【答案】内容并入 answer（原 answer 为空时），
    【详解】/【分析】/【解析】/【解答】内容追加到 analysis，保证附录不丢内容。"""
    s = it["stem"]
    if not any(t in s for t in _EMBED_TAGS):
        return
    m = re.search(r"【(?:答案|详解|分析|解析|解答)】", s)
    head, tail = s[:m.start()].strip(), s[m.start():]
    it["stem"] = head
    # 按标签拆分 tail，归类答案段与解析段
    ans_txt, extra_segs = "", []
    cur = None
    for seg in re.split(r"(【答案】|【详解】|【分析】|【解析】|【解答】)", tail):
        if seg in _EMBED_TAGS:
            cur = seg
            continue
        if not seg:
            continue
        if cur == "【答案】":
            ans_txt += seg
        else:
            extra_segs.append(seg.strip())
    if not it["answer"] and ans_txt.strip():
        it["answer"] = ans_txt.strip()
    elif ans_txt.strip() and ans_txt.strip() != it["answer"]:
        extra_segs.append("（原题内嵌答案）" + ans_txt.strip())
    extra = "\n".join(x for x in extra_segs if x).strip()
    if extra:
        it["analysis"] = (it["analysis"] + "\n" + extra).strip() if it["analysis"] else extra


STAR = "★☆"


def build(subject, output):
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["questions"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}

    items = []
    for r in rows[1:]:
        if not r or not r[idx.get("qid", 0)]:
            continue
        if str(r[idx["subject"]] or "").strip() != subject:
            continue
        try:
            options = json.loads(r[idx["options"]] or "[]")
            if not isinstance(options, list):
                options = []
        except Exception:
            options = []
        try:
            media = json.loads(r[idx["media_paths"]] or "[]")
            if not isinstance(media, list):
                media = []
        except Exception:
            media = []
        it = {
            "qid": str(r[idx["qid"]]),
            "chapter": str(r[idx["chapter"]] or "未分章").strip() or "未分章",
            "kp": str(r[idx["knowledge_point"]] or "").strip(),
            "qtype": str(r[idx["qtype"]] or "").strip(),
            "difficulty": r[idx["difficulty"]] or 3,
            "stem": str(r[idx["stem"]] or "").strip(),
            "options": options,
            "answer": str(r[idx["answer"]] or "").strip(),
            "analysis": str(r[idx["analysis"]] or "").strip(),
            "source": str(r[idx["source"]] or "").strip(),
            "media": media,
        }
        _split_stem_embedded_answer(it)
        items.append(it)
    wb.close()
    if not items:
        raise SystemExit("题库中没有科目为「%s」的题目" % subject)

    # 按章节分组，保持首次出现顺序；章内按 qid
    chapters = OrderedDict()
    for it in items:
        chapters.setdefault(it["chapter"], []).append(it)

    doc = Document()
    # 页面 A4 + 页边距
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)

    # 页脚页码
    footer_p = sec.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(footer_p)

    # 页眉
    head_p = sec.header.paragraphs[0]
    head_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = head_p.add_run("高考%s 练习习题集（按知识点汇编）" % subject)
    hr.font.size = Pt(9)
    hr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    set_cn_font(hr)

    # Normal 样式
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10.5)
    set_cn_font(normal)

    # ---------- 封面 ----------
    for _ in range(6):
        para(doc, "")
    para(doc, "高考%s 练习习题集" % subject, size=30, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, cn="黑体", space_after=10)
    para(doc, "—— 按知识点分类汇编 · 含答案与解析 ——", size=13,
         color=(0x66, 0x66, 0x66), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)

    qtype_cnt = Counter(it["qtype"] for it in items)
    qtype_line = "　|　".join("%s %d 题" % (k, v) for k, v in qtype_cnt.most_common())
    stat_lines = [
        "科目：%s" % subject,
        "题量：共 %d 题，按 %d 个章节 / 知识点模块编排" % (len(items), len(chapters)),
        "题型分布：%s" % qtype_line,
        "编排：正文只含题目（供线下作答）；答案与解题步骤统一编排在文末附录，做完再对照",
        "生成日期：%s" % datetime.date.today().strftime("%Y 年 %m 月 %d 日"),
        "来源：%s" % "高考真题与模拟题题库（系统导入）",
    ]
    for ln in stat_lines:
        para(doc, ln, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    for _ in range(5):
        para(doc, "")
    para(doc, "（打印后可线下练习；先在正文完成作答，再翻至文末附录核对答案与解题步骤）",
         size=10.5, color=(0x99, 0x99, 0x99), align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ---------- 目录（章节清单） ----------
    para(doc, "目　录", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         cn="黑体", space_after=12)
    for i, (ch, ch_items) in enumerate(chapters.items(), 1):
        p, r = para(doc, "第 %d 章　%s ……（%d 题）" % (i, ch, len(ch_items)),
                    size=11, space_after=3)
    para(doc, "附　录　答案与解题步骤 ……", size=11, bold=True, space_after=3)
    doc.add_page_break()

    # ---------- 正文 ----------
    q_no = 0
    for i, (ch, ch_items) in enumerate(chapters.items(), 1):
        h = doc.add_heading("", level=1)
        hr2 = h.add_run("第 %d 章　%s（%d 题）" % (i, ch, len(ch_items)))
        hr2.font.size = Pt(15)
        hr2.font.color.rgb = RGBColor(0x1F, 0x3B, 0x66)
        hr2.font.bold = True
        set_cn_font(hr2, cn="黑体")

        for it in ch_items:
            q_no += 1
            # 题号 + 题型 + 难度
            d = it["difficulty"]
            try:
                d = int(d)
            except Exception:
                d = 3
            d = max(1, min(5, d))
            meta = "%s　难度 %s%s" % (it["qtype"] or "题目", STAR[:d], STAR[d:])
            p = doc.add_paragraph()
            r1 = p.add_run("%d. " % q_no)
            r1.font.bold = True
            r1.font.size = Pt(11)
            set_cn_font(r1)
            r2 = p.add_run(meta)
            r2.font.size = Pt(9)
            r2.font.color.rgb = RGBColor(0x2E, 0x6C, 0xE6)
            set_cn_font(r2)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = True

            # 题干（含图片）
            add_stem_with_images(doc, it["stem"])

            # 选项
            if it["options"]:
                letters = "ABCDEFGH"
                for j, opt in enumerate(it["options"][:8]):
                    opt_txt = latex_to_unicode(str(opt))
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Cm(0.8)
                    p.paragraph_format.space_after = Pt(2)
                    ra = p.add_run("%s. " % letters[j])
                    ra.font.bold = True
                    ra.font.size = Pt(10.5)
                    set_cn_font(ra)
                    ro = p.add_run(opt_txt)
                    ro.font.size = Pt(10.5)
                    set_cn_font(ro)

            # media_paths 附加图（题干标记之外）
            stem_raw = it["stem"]
            for fname in it["media"]:
                if not fname or ("[IMG:" + fname + "]") in stem_raw:
                    continue
                fpath = os.path.join(MEDIA_DIR, str(fname))
                if os.path.exists(fpath):
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        p.add_run().add_picture(fpath, width=Cm(6.5))
                    except Exception:
                        pass

            # 知识点
            if it["kp"]:
                p, r = para(doc, "知识点：%s" % it["kp"], size=9,
                            color=(0x8A, 0x8A, 0x8A), space_after=2)
                p.paragraph_format.left_indent = Cm(0.2)
                p.paragraph_format.space_before = Pt(2)

    # ---------- 附录：答案与解题步骤 ----------
    doc.add_page_break()
    ah = doc.add_heading("", level=1)
    ahr = ah.add_run("附录　答案与解题步骤")
    ahr.font.size = Pt(16)
    ahr.font.color.rgb = RGBColor(0x1F, 0x3B, 0x66)
    ahr.font.bold = True
    set_cn_font(ahr, cn="黑体")
    para(doc, "（题号与正文一致：先在正文完成作答，再按题号在此核对）",
         size=9.5, color=(0x99, 0x99, 0x99), space_after=10)

    a_no = 0
    for i, (ch, ch_items) in enumerate(chapters.items(), 1):
        p = doc.add_paragraph()
        chr_ = p.add_run("第 %d 章　%s" % (i, ch))
        chr_.font.size = Pt(12)
        chr_.font.bold = True
        chr_.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        set_cn_font(chr_, cn="黑体")
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True

        for it in ch_items:
            a_no += 1
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Cm(0.2)
            n1 = p.add_run("第 %d 题" % a_no)
            n1.font.bold = True
            n1.font.size = Pt(10.5)
            set_cn_font(n1)
            la = p.add_run("　【答案】")
            la.font.bold = True
            la.font.size = Pt(10.5)
            la.font.color.rgb = RGBColor(0x1B, 0x7A, 0x43)
            set_cn_font(la)
            va = p.add_run(latex_to_unicode(it["answer"]) if it["answer"] else "（待补充）")
            va.font.size = Pt(10.5)
            if not it["answer"]:
                va.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
            set_cn_font(va)

            if it["analysis"]:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.left_indent = Cm(0.2)
                lp = p.add_run("【解析】")
                lp.font.bold = True
                lp.font.size = Pt(10.5)
                lp.font.color.rgb = RGBColor(0x8A, 0x51, 0x0A)
                set_cn_font(lp)
                vap = p.add_run(latex_to_unicode(it["analysis"]))
                vap.font.size = Pt(10.5)
                set_cn_font(vap)

    doc.save(output)
    return len(items), len(chapters)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    n, c = build(args.subject, args.output)
    print("OK：%s → %s（%d 题 / %d 章）" % (args.subject, args.output, n, c))


if __name__ == "__main__":
    main()
