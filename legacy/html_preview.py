# -*- coding: utf-8 -*-
"""
html_preview.py — 题目富媒体 HTML 预览
用 KaTeX（数学公式）+ mhchem（化学方程式）在浏览器中渲染题目，
图片以 base64 内嵌（单文件自包含），默认浏览器打开即可。
适用：题干含 LaTeX 公式 / 化学方程式 / 图片的题目在桌面端无法直接展示时。
"""
import os
import re
import json
import base64
import html
import tempfile
import webbrowser

import excel_store as store


_HEAD = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>题目预览</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
<script src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/mhchem.min.js"></script>
<style>
  body { font-family: "Microsoft YaHei UI", "PingFang SC", sans-serif;
         max-width: 880px; margin: 24px auto; padding: 0 20px; color: #1A202C;
         background: #F7FAFC; }
  .q-card { background: #fff; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 18px 20px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
  .q-head { color: #4A5568; font-size: 13px; margin-bottom: 8px; }
  .badge { display: inline-block; background: #EDF2F7; color: #2D3748; border-radius: 4px;
           padding: 1px 8px; margin-right: 6px; font-size: 12px; }
  .badge.media { background: #EBF8FF; color: #2B6CB0; }
  .stem { font-size: 15px; line-height: 1.9; white-space: pre-wrap; }
  .opts { margin: 10px 0 0 0; padding-left: 22px; line-height: 1.9; }
  .ans { background: #F0FFF4; border-left: 4px solid #38A169; padding: 8px 12px;
         margin-top: 10px; font-size: 14px; }
  .ana { background: #FFFAF0; border-left: 4px solid #DD6B20; padding: 8px 12px;
         margin-top: 8px; font-size: 14px; line-height: 1.8; white-space: pre-wrap; }
  .q-img { max-width: 100%; border: 1px solid #E2E8F0; border-radius: 6px;
           margin: 8px 0; display: block; }
  h3 { color: #2D3748; }
</style>
</head>
<body>
"""


def _esc(s):
    return html.escape(str(s or ""), quote=False)


def _img_html(fname):
    """把 data/media/ 下的图片转成 base64 内嵌 <img>。"""
    path = fname if os.path.isabs(fname) else os.path.join(store.media_dir(), fname)
    if not os.path.exists(path):
        return '<div style="color:#C53030;font-size:12px">[图片缺失: %s]</div>' % _esc(os.path.basename(fname))
    ext = os.path.splitext(path)[1].lower().strip(".") or "png"
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif",
            "bmp": "bmp", "webp": "webp"}.get(ext, "png")
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return '<img class="q-img" src="data:image/%s;base64,%s">' % (mime, b64)


def _options_html(options):
    if isinstance(options, str):
        try:
            options = json.loads(options)
        except Exception:
            options = [options] if options else []
    if not options:
        return ""
    lis = "".join("<li>%s</li>" % _esc(o) for o in options)
    return '<ol class="opts" type="A">%s</ol>' % lis


def render_question_html(q):
    """渲染单道题（dict）为 HTML 卡片片段。"""
    media_type = q.get("media_type", "无")
    media_paths = store.media_paths_of(q)
    head = ('<span class="badge">%s</span><span class="badge">%s</span>'
            '<span class="badge">%s</span><span class="badge">难度 %s</span>'
            '<span class="badge">%s</span>'
            % (_esc(q.get("subject", "")), _esc(q.get("chapter", "")),
               _esc(q.get("knowledge_point", "")), _esc(q.get("difficulty", "")),
               _esc(q.get("tag", "训练题"))))
    if media_type and media_type != "无":
        head += '<span class="badge media">图形方案: %s</span>' % _esc(media_type)
    imgs = "".join(_img_html(p) for p in media_paths)
    body = '<div class="stem">%s</div>' % _esc(q.get("stem", ""))
    body += imgs
    body += _options_html(q.get("options", []))
    if str(q.get("answer", "")).strip():
        body += '<div class="ans"><b>答案：</b>%s</div>' % _esc(q.get("answer", ""))
    if str(q.get("analysis", "")).strip():
        body += '<div class="ana"><b>解析：</b>%s</div>' % _esc(q.get("analysis", ""))
    return '<div class="q-card"><div class="q-head">%s</div>%s</div>' % (head, body)


_TAIL = """
<script>
  document.addEventListener("DOMContentLoaded", function() {
    renderMathInElement(document.body, {
      delimiters: [
        {left: "$$", right: "$$", display: true},
        {left: "\\\\(", right: "\\\\)", display: false},
        {left: "\\\\[", right: "\\\\]", display: true},
        {left: "$", right: "$", display: false}
      ],
      throwOnError: false
    });
  });
</script>
</body>
</html>
"""


def render_questions_html(questions, title="题目预览"):
    """把多道题渲染成完整 HTML 文档字符串。"""
    cards = "\n".join(render_question_html(q) for q in questions)
    return _HEAD + ("<h3>%s（共 %d 题）</h3>" % (_esc(title), len(questions))) + cards + _TAIL


def preview_questions(questions, title="题目预览"):
    """生成 HTML 文件并在默认浏览器打开，返回文件路径。"""
    doc = render_questions_html(questions, title)
    out = os.path.join(tempfile.gettempdir(), "gk_preview.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    webbrowser.open("file://" + out.replace("\\", "/"))
    return out
