# -*- coding: utf-8 -*-
"""
ai_recognize.py — 大模型识别模块
上传试卷（扫描件 / Word / PDF）→ 文件预处理 → 长文本按题号分段 → 调用云端大模型 API →
题目结构化（题干/选项/答案/解析）+ 科目/题型/知识点自动分类 + 置信度。
未配置 API 时进入“演示模式”：按题号规则切分文本，方便无网环境走通流程。
"""
import os
import re
import json
import base64
import tempfile
import requests

import excel_store as store

SETTINGS_PATH = os.path.join(store.DATA_DIR, "settings.json")

FIELDS = ["subject", "chapter", "knowledge_point", "qtype", "difficulty",
          "stem", "options", "answer", "analysis", "confidence"]


# ---------------- 设置 ----------------
# 两套独立 API：
#   text   — 用于 Word / 文本型 PDF 题目识别（普通大模型即可）
#   vision — 用于扫描件 / 图片型 PDF / 阅卷（需要多模态大模型，如 qwen-vl-max、gpt-4o 等）
def _default_settings():
    return {
        "text":   {"api_base": "https://api.deepseek.com",  "path": "/v1/chat/completions",
                    "api_key": "", "model": "deepseek-v4-pro"},
        "vision": {"api_base": "https://dashscope.aliyuncs.com/compatible-mode",
                    "path": "/v1/chat/completions",
                    "api_key": "", "model": "qwen-vl-max"},
    }


def load_settings():
    """读取并兼容旧版 settings.json。返回 {'text': {...}, 'vision': {...}}。"""
    cfg = _default_settings()
    if not os.path.exists(SETTINGS_PATH):
        return cfg
    try:
        raw = json.load(open(SETTINGS_PATH, encoding="utf-8"))
        # 旧版平铺字段自动迁移
        if "text" not in raw and ("api_base" in raw or "api_key_enc" in raw or "model" in raw):
            text = {"api_base": raw.get("api_base", cfg["text"]["api_base"]),
                    "path": "/v1/chat/completions",
                    "model": raw.get("model", cfg["text"]["model"])}
            if "api_key_enc" in raw:
                text["api_key"] = base64.b64decode(raw["api_key_enc"]).decode("utf-8")
            cfg["text"] = text
        if "vision" not in raw:
            base = raw.get("api_base", cfg["text"]["api_base"])
            key = base64.b64decode(raw["api_key_enc"]).decode("utf-8") if raw.get("api_key_enc") else ""
            cfg["vision"] = {"api_base": base, "path": "/v1/chat/completions",
                              "api_key": key, "model": raw.get("vision_model", "")}
        # 新版节段覆盖默认值
        for k in ("text", "vision"):
            sub = raw.get(k, {})
            for kk in ("api_base", "api_key", "model", "path"):
                if kk in sub and sub[kk]:
                    cfg[k][kk] = sub[kk]
        return cfg
    except Exception:
        return cfg


def save_settings(text, vision):
    """text/vision 各为 dict：{'api_base','path','api_key','model'}。"""
    raw = {}
    for k, d in (("text", text), ("vision", vision)):
        key = d.get("api_key", "").strip()
        path = d.get("path", "").strip() or "/v1/chat/completions"
        if not path.startswith("/"):
            path = "/" + path
        raw[k] = {
            "api_base": d.get("api_base", "").strip().rstrip("/"),
            "path": path,
            "api_key": key,
            "api_key_enc": base64.b64encode(key.encode("utf-8")).decode("ascii") if key else "",
            "model": d.get("model", "").strip(),
        }
    json.dump(raw, open(SETTINGS_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def api_configured(kind="text"):
    return bool(load_settings()[kind]["api_key"])


# ---------------- 文件预处理 ----------------
def extract_content(file_path):
    """
    返回 (kind, payload)：
      kind="text"  → payload 为抽取到的文本
      kind="images"→ payload 为图片路径列表（扫描件 / 图片型PDF，需视觉模型）
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".docx", ".doc"):
        from docx import Document
        doc = Document(file_path)
        # 带图片原位提取：段落内的图片转为 [IMG:文件名] 标记跟随题目流走
        text = _extract_docx_with_images(doc, file_path)
        # 清理 python-docx 对下标/上标的 HTML 标签碎片（如 <sub>3</sub>）
        text = _TAG_CLEAN.sub("", text)
        return "text", text
    if ext == ".pdf":
        import fitz
        doc = fitz.open(file_path)
        text = "\n".join(page.get_text() for page in doc)
        if len(text.strip()) >= 50:          # 文本型 PDF
            doc.close()
            return "text", text
        # 图片型 PDF → 逐页渲染 + 严格缩压（GLM-4V 等对单图大小敏感）
        imgs = []
        tmp = tempfile.mkdtemp(prefix="gk_pdf_")
        n_pages = doc.page_count
        for i, page in enumerate(doc):
            # 降低 dpi：100 dpi 足以 OCR，长边通常约 1000-1300px
            pix = page.get_pixmap(dpi=100)
            p = os.path.join(tmp, "page_%03d.png" % (i + 1))
            pix.save(p)
            imgs.append(p)
        doc.close()
        # 每页都过 _prep_image_for_vision 缩压到 < 800KB，避免请求过大
        prepped = []
        for idx, p in enumerate(imgs):
            try:
                p2 = _prep_image_for_vision(p, max_side=1100, max_bytes=700_000)
                if p2 and p2[0]:
                    prepped.append(p2[0])
                else:
                    prepped.append(p)
            except Exception:
                prepped.append(p)
        return "images", prepped
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"):
        return "images", _prep_image_for_vision(file_path)
    if ext == ".txt":
        return "text", open(file_path, encoding="utf-8", errors="ignore").read()
    raise ValueError("不支持的文件格式：%s（支持 docx / pdf / 图片 / txt）" % ext)


# ---------------- 图形图表：媒体目录与 docx 图片提取 ----------------
def media_dir():
    """题目图形/图片统一存放目录：data/media/（单独目录管理）。"""
    d = os.path.join(store.DATA_DIR, "media")
    os.makedirs(d, exist_ok=True)
    return d


_IMG_MARK = re.compile(r"\[IMG:([^\]]+)\]")


def _extract_docx_with_images(doc, file_path):
    """从 docx 提取文本，段落内嵌图片保存到 data/media/ 并在原位置插入 [IMG:文件名] 标记。
    标记会跟随题号切分进入对应题目块，最终映射到 media_paths 字段。"""
    from docx.oxml.ns import qn
    base = os.path.splitext(os.path.basename(file_path))[0]
    safe = re.sub(r"[^\w一-鿿]+", "_", base)[:40]
    img_i = 0
    lines = []
    for p in doc.paragraphs:
        seg = []
        for run in p.runs:
            if run.text:
                seg.append(run.text)
            # 查找 run 内的 drawing（内嵌图片）
            for d in run._element.findall(".//" + qn("w:drawing")):
                blip = d.find(".//" + qn("a:blip"))
                if blip is None:
                    continue
                rid = blip.get(qn("r:embed"))
                if not rid:
                    continue
                try:
                    part = doc.part.related_parts[rid]
                except KeyError:
                    continue
                img_i += 1
                ext2 = os.path.splitext(str(part.partname))[1] or ".png"
                fname = "%s_img%02d%s" % (safe, img_i, ext2)
                with open(os.path.join(media_dir(), fname), "wb") as f:
                    f.write(part.blob)
                seg.append("[IMG:%s]" % fname)
        line = "".join(seg).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def _strip_img_marks(text):
    """剥离 [IMG:xxx] 标记，返回 (纯文本, 图片文件名列表)。"""
    imgs = _IMG_MARK.findall(text or "")
    clean = _IMG_MARK.sub(" ", text or "").strip()
    clean = re.sub(r"[ \t]+", " ", clean)
    return clean, imgs


# 图形类型关键词 → 渲染方案标记
_MEDIA_RULES = [
    ("物理示意图", ("电路", "电流表", "电压表", "受力", "光路", "磁场", "磁感线", "滑轮",
                 "杠杆", "浮力", "压强计", "透镜", "弹簧", "运动图像", "v-t", "x-t")),
    ("分子结构式", ("结构式", "分子结构", "苯环", "同分异构", "有机", "官能团", "羟基", "羧基")),
    ("化学方程式", ("化学方程式", "反应式", "离子方程式", "化学平衡", "电极反应")),
    ("函数图像", ("函数图像", "函数图象", "y=f(x)", "y＝f(x)", "抛物线", "双曲线", "椭圆",
                "曲线C", "图像如图", "图象如图", "如图所示的曲线")),
    ("几何图形", ("三角形", "正方形", "矩形", "菱形", "平行四边形", "梯形", "正方体", "长方体",
                "圆锥", "圆柱", "球", "棱锥", "棱柱", "如图，在", "如图所示")),
]
_FORMULA_PAT = re.compile(
    r"\\(frac|sqrt|sum|prod|int|lim|sin|cos|tan|log|ln|alpha|beta|Delta|pi|theta|ce)"
    r"|\$[^$]+\$"                 # $...$ 行内公式
    r"|\\\(|\\\["                 # \( 或 \[ 公式定界
    r"|[A-Za-z0-9]\^\d|[A-Za-z]_\d|[A-Za-z]\^[A-Za-z]"   # x^2、x_1、a^n 下上标
)


def classify_media_type(text, has_images=False):
    """根据题目文本判定图形渲染方案标记。返回中文类型名。"""
    t = text or ""
    for name, kws in _MEDIA_RULES:
        if any(k in t for k in kws):
            return name
    if _FORMULA_PAT.search(t):
        return "数学公式"
    if has_images:
        return "其他图表"
    return "无"


def _prep_image_for_vision(file_path, max_side=1100, max_bytes=700_000, quality=82, force=False):
    """对单张图片做缩压。force=True 时即使已存在也重新处理（用于重试更小图）。"""
    from PIL import Image
    try:
        im = Image.open(file_path)
        im.load()
    except Exception:
        return [file_path]
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGB")
    elif im.mode != "RGB":
        im = im.convert("RGB")
    w, h = im.size
    if max(w, h) > max_side:
        ratio = max_side / max(w, h)
        im = im.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    # 临时文件用 attempt 参数避免重试时缓存
    suffix = "_s%d_q%d" % (max_side, quality)
    tmp = os.path.join(tempfile.gettempdir(),
                       "gk_vision_%s%s.jpg" % (os.path.basename(file_path), suffix))
    try:
        im.save(tmp, "JPEG", quality=quality, optimize=True)
    except Exception:
        return [file_path]
    if os.path.getsize(tmp) > max_bytes and quality > 60:
        try:
            im.save(tmp, "JPEG", quality=max(60, quality - 10), optimize=True)
        except Exception:
            pass
    if force and file_path != tmp and os.path.exists(file_path) is False:
        pass
    return [tmp]


# ---------------- 大模型调用 ----------------
_TAG_CLEAN = re.compile(r"</?(?:sub|sup|SUB|SUP)\s*>")  # docx 下标/上标标签碎片清理
_PROMPT = (
    "你是高考题库录入助手。请从用户提供的试卷内容中抽取所有题目，"
    "输出严格的 JSON 数组，不要输出任何其他文字。每个元素字段：\n"
    "subject(科目，从 语文/数学/英语/物理/化学/生物 中选)、"
    "chapter(章节)、knowledge_point(知识点，尽量具体，多个用英文逗号分隔)、"
    "qtype(题型，从 单选/多选/填空/解答 中选)、difficulty(难度，1-5的整数)、"
    "stem(题干原文)、options(选项文本数组，不含A.B.C.D.标号；非选择题为空数组)、"
    "answer(答案)、analysis(详细解析，含解题步骤)、confidence(识别置信度0-1小数)、"
    "tag(题目标签)。\n"
    "tag 规则：能从题面或试卷信息中确定年份/地区/考试类型则填 '年份/地区/类型'，"
    "例如 '2023/北京/高考真题'、'2024/全国Ⅰ卷/高考真题'、'2025/广东/高考模拟'；"
    "无法确定则填 '训练题'。\n"
    "格式要求：1) 字符串内不要出现未转义的英文双引号，引用请用中文引号「」；"
    "2) 字符串内不要出现真实换行，多行内容用 \\n 表示；"
    "3) 数学公式中的反斜杠必须写成 \\\\（如 \\\\frac{a}{b}、\\\\( x \\\\)）；"
    "4) 不要使用尾逗号；5) 只输出 JSON 数组，不要 markdown 代码块标记；"
    "6) 题目中如出现 [IMG:xxx] 标记（表示该处有一张图片），请原样保留在 stem 字段中，不要删除、翻译或改写。"
)


def _escape_bad_backslash(s):
    """把不是合法 JSON 转义的反斜杠（如 LaTeX 的 \\( \\frac）补成 \\\\。"""
    return re.sub(r'\\(?![\\"/bfnrtu])', r'\\\\', s)


def _loads_loose(raw):
    """宽松解析 JSON 文本，依次尝试：标准 → 允许控制字符 → 去尾逗号 → 修复反斜杠 → 组合。"""
    cands = [
        lambda s: json.loads(s),
        lambda s: json.loads(s, strict=False),
        lambda s: json.loads(re.sub(r",\s*([}\]])", r"\1", s), strict=False),
        lambda s: json.loads(_escape_bad_backslash(s), strict=False),
        lambda s: json.loads(re.sub(r",\s*([}\]])", r"\1", _escape_bad_backslash(s)),
                             strict=False),
    ]
    last_err = None
    for fn in cands:
        try:
            return fn(raw)
        except Exception as e:
            last_err = e
    raise last_err if last_err else ValueError("JSON 解析失败")


def _parse_json_array(text):
    """解析大模型输出为题目对象列表。
    优先整体解析 JSON 数组（含宽松修复）；数组被截断/混有解释文字时，
    逐个抢救完整的题目对象（raw_decode 支持嵌套）。"""
    text = text.strip()
    m = re.search(r"\[.*\]", text, re.S)
    if m:
        try:
            data = _loads_loose(m.group(0))
            if isinstance(data, list):
                return data
        except Exception:
            pass  # 数组被截断或含坏字符，转抢救模式
    # 抢救模式：逐个解析完整的 {...} 对象，保留含 stem 的
    dec = json.JSONDecoder(strict=False)
    objs, i = [], 0
    while True:
        j = text.find("{", i)
        if j < 0:
            break
        try:
            d, end = dec.raw_decode(_escape_bad_backslash(text[j:]))
            if isinstance(d, dict) and str(d.get("stem", "")).strip():
                objs.append(d)
            i = j + max(end, 1)
        except Exception:
            i = j + 1
    if objs:
        return objs
    snippet = text[:120].replace("\n", " ")
    raise ValueError("大模型返回内容不是有效 JSON（开头片段：%s…）" % snippet)


# ---------------- 长文本分段 ----------------
CHUNK_CHARS = 2500          # 每段最大字符数（约 5~8 道题，保证输出不被截断）


def _split_question_blocks(text):
    """按题号边界把试卷文本切成完整题目块，返回 (卷头, 题目块列表)。
    支持格式：1.  1、  1．  1:  1：  题1：  第1题（选择题，3分）  等。"""
    parts = re.split(r"(?m)^\s*(?:题\s*|第\s*)?(\d{1,3})(?:\s*[\.、．：:]|\s*题)", text)
    prefix = parts[0].strip()
    blocks = []
    for i in range(1, len(parts) - 1, 2):
        body = parts[i + 1].strip()
        # 丢弃太短的块（页码/纯目录条目等误切片段）
        if len(body) < 12:
            prefix += "\n" + parts[i] + "题 " + body
            continue
        blocks.append(parts[i] + "题 " + body)
    if not blocks and text.strip():
        blocks = [text.strip()]
    return prefix, blocks


def _chunk_blocks(blocks, max_chars=CHUNK_CHARS):
    """把题目块按长度装箱为若干段，保证不切断任何一道题。"""
    chunks, cur, cur_len = [], [], 0
    for b in blocks:
        if cur and cur_len + len(b) > max_chars:
            chunks.append("\n\n".join(cur))
            cur, cur_len = [], 0
        cur.append(b)
        cur_len += len(b)
    if cur:
        chunks.append("\n\n".join(cur))
    return chunks


def _normalize_items(data, source):
    items = []
    for d in data:
        if not isinstance(d, dict) or not str(d.get("stem", "")).strip():
            continue
        try:
            difficulty = int(d.get("difficulty", 3) or 3)
        except (TypeError, ValueError):
            difficulty = 3
        try:
            confidence = float(d.get("confidence", 0.8) or 0.8)
        except (TypeError, ValueError):
            confidence = 0.8
        # 剥离 [IMG:xxx] 标记 → media_paths（题目图片单独存 data/media/）
        stem_clean, imgs1 = _strip_img_marks(str(d.get("stem", "")).strip())
        ana_clean, imgs2 = _strip_img_marks(str(d.get("analysis", "")).strip())
        opts_raw = d.get("options") if isinstance(d.get("options"), list) else []
        opts_clean, imgs3 = [], []
        for o in opts_raw:
            oc, oi = _strip_img_marks(str(o))
            opts_clean.append(oc)
            imgs3.extend(oi)
        media_paths = imgs1 + imgs2 + imgs3
        if not stem_clean:
            continue
        it = {
            "subject": str(d.get("subject", "")).strip() or "数学",
            "chapter": str(d.get("chapter", "")).strip() or "未分章",
            "knowledge_point": str(d.get("knowledge_point", "")).strip() or "待标注",
            "qtype": str(d.get("qtype", "")).strip() or "单选",
            "difficulty": difficulty,
            "stem": stem_clean,
            "options": opts_clean,
            "answer": str(d.get("answer", "")).strip(),
            "analysis": ana_clean,
            "confidence": confidence,
            "source": source,
            "tag": _derive_tag(d, source),
            # 图形方案标记：根据题目内容判定渲染方案
            "media_type": classify_media_type(stem_clean + " " + ana_clean,
                                              has_images=bool(media_paths)),
            "media_paths": media_paths,
        }
        if it["qtype"] not in ("单选", "多选", "填空", "解答"):
            it["qtype"] = "单选" if it["options"] else "填空"
        it["difficulty"] = max(1, min(5, it["difficulty"]))
        items.append(it)
    return items


# 常见省份/试卷简称 → 地区名（识别时优先匹配）
_REGIONS = ("北京", "上海", "天津", "重庆", "广东", "江苏", "浙江", "山东", "福建", "湖北",
            "湖南", "河南", "河北", "安徽", "江西", "辽宁", "吉林", "黑龙江", "内蒙古", "新疆",
            "西藏", "青海", "甘肃", "宁夏", "陕西", "山西", "四川", "贵州", "云南", "广西",
            "海南", "台湾", "香港", "澳门", "全国")
_EXAM_TYPES = ("高考真题", "高考模拟", "校考", "联考", "竞赛", "模拟卷", "期中", "期末", "一模", "二模")


def _derive_tag(item, source):
    """根据 AI 输出的 tag 字段 / 题干 / 源文件名 推断标签，格式 '年份/地区/类型' 或 '训练题'。"""
    raw = str(item.get("tag", "")).strip()
    if raw and raw != "训练题" and "/" in raw:
        return raw
    # 拼接所有可读信息
    text = " ".join([str(item.get("stem", "")), str(item.get("analysis", "")),
                     source or ""])
    yrs = re.findall(r"(?:19|20)\d{2}", text)
    year = yrs[0] if yrs else None
    region = next((r for r in _REGIONS if r in text), None)
    exam = next((t for t in _EXAM_TYPES if t in text), None)
    if year and region and exam:
        return "%s/%s/%s" % (year, region, exam)
    if year and region:
        return "%s/%s/高考题" % (year, region)
    if region and exam:
        return "%s/%s" % (region, exam)
    if year and exam:
        return "%s/全国/%s" % (year, exam)
    return "训练题"


def _chat(messages, kind="text", model=None, max_tokens=8000, timeout=180):
    """kind='text' 用文本 API；'vision' 用视觉 API。两套配置完全独立。"""
    cfg = load_settings()[kind]
    if not cfg["api_key"]:
        kind_name = "文本" if kind == "text" else "视觉"
        raise ValueError("未配置%s API，请在教师端“API设置”中填写 api_key" % kind_name)
    path = cfg.get("path") or "/v1/chat/completions"
    if not path.startswith("/"):
        path = "/" + path
    url = cfg["api_base"].rstrip("/") + path
    headers = {"Authorization": "Bearer " + cfg["api_key"],
               "Content-Type": "application/json"}
    body = {"model": model or cfg["model"], "messages": messages,
            "temperature": 0.2, "max_tokens": max_tokens}
    try:
        resp = requests.post(url, headers=headers, json=body, timeout=timeout)
    except requests.exceptions.ConnectionError as e:
        # 包含 Windows 10053 (软件中止已建连接)，通常为图片过大/防火墙/杀毒
        raise ConnectionError(_fmt_network_error(kind, e)) from e
    except requests.exceptions.Timeout as e:
        raise ConnectionError(_fmt_network_error(kind, e, "请求超时")) from e
    except requests.exceptions.RequestException as e:
        raise ConnectionError(_fmt_network_error(kind, e)) from e
    if resp.status_code == 401:
        raise ValueError("[%s API] 401 未授权：API Key 无效或过期" % ("文本" if kind=="text" else "视觉"))
    if resp.status_code == 404:
        raise ValueError("[%s API] 404 地址错误：base_url 或模型名写错（%s）"
                          % ("文本" if kind=="text" else "视觉", url))
    if resp.status_code in (400, 413):
        # 请求过大（图片超大等）
        raise ValueError("[%s API] 请求被拒绝（%d）：可能是图片过大或模型名不支持。\n"
                          "建议：换更小的扫描件、降低dpi，或确认模型名（当前：%s）"
                          % ("文本" if kind=="text" else "视觉", resp.status_code, body["model"]))
    try:
        resp.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise ValueError("[%s API] HTTP %s：%s" % ("文本" if kind=="text" else "视觉",
                                                   resp.status_code, e)) from e
    try:
        return resp.json()["choices"][0]["message"]["content"]
    except (KeyError, IndexError, ValueError) as e:
        raise ValueError("[%s API] 返回内容格式异常：%s" % ("文本" if kind=="text" else "视觉", str(e)[:200]))


def _fmt_network_error(kind, exc, prefix=""):
    kind_name = "文本" if kind == "text" else "视觉"
    msg = ("[%s API] 网络连接失败（%s）" % (kind_name, prefix or type(exc).__name__))
    s = str(exc)
    if "10053" in s or "ConnectionAborted" in s:
        msg += "\n可能是图片太大、防火墙/杀毒拦截、或本地代理导致连接被中止。"
    elif "timeout" in s.lower() or "timed out" in s.lower():
        msg += "\n请求超时，建议缩小图片或换更快的网络。"
    elif "refused" in s.lower():
        msg += "\n服务器拒绝连接（地址错误/端口不通/服务未启动）。"
    msg += "\n  建议：\n  ① 在“API设置”点“测试连接”验证 Key 有效；\n  ② 若是视觉 API 且图片 > 1.5MB，本系统已自动压缩到 1600px/JPEG，确认源图本身可读；\n  ③ 检查防火墙/杀毒/代理是否拦截了 HTTPS 出站；\n  ④ 尝试换其他 base_url（如官方直连）。"
    return msg


def test_api(kind="text"):
    """测试指定 API 是否可用：发送一个最简请求，返回 (ok, message)。"""
    cfg = load_settings()[kind]
    if not cfg["api_key"]:
        return False, "未配置 API Key"
    if not cfg["api_base"]:
        return False, "未配置 base_url"
    if not cfg["model"]:
        return False, "未配置模型名"
    path = cfg.get("path") or "/v1/chat/completions"
    if not path.startswith("/"):
        path = "/" + path
    url = cfg["api_base"].rstrip("/") + path
    headers = {"Authorization": "Bearer " + cfg["api_key"],
               "Content-Type": "application/json"}
    body = {"model": cfg["model"], "messages": [{"role": "user", "content": "ping"}],
            "max_tokens": 8, "temperature": 0}
    try:
        r = requests.post(url, headers=headers, json=body, timeout=20)
    except Exception as e:
        return False, _fmt_network_error(kind, e)
    if r.status_code == 200:
        return True, "连接成功 ✔（模型：%s）" % cfg["model"]
    if r.status_code == 401:
        return False, "401 未授权：API Key 无效或过期"
    if r.status_code == 404:
        return False, "404 找不到接口：base_url 或模型名错误"
    if r.status_code in (400, 413):
        return False, "%d 请求错误：%s" % (r.status_code, r.text[:120])
    return False, "HTTP %d：%s" % (r.status_code, r.text[:120])


def _call_chunk(chunk_text, context, subject_hint, strict=False):
    """识别一个文本段；strict=True 时追加更严格的输出约束（用于重试）。"""
    hint = ("本试卷科目提示：" + subject_hint + "。") if subject_hint else ""
    user = hint
    if context:
        user += "试卷信息：" + context[:200] + "。"
    user += "以下是试卷的部分题目，请全部抽取：\n" + chunk_text
    if strict:
        user += "\n\n重要：只输出 JSON 数组本身，不要输出任何解释、标题或 markdown 标记。"
    content = _chat([{"role": "system", "content": _PROMPT},
                     {"role": "user", "content": user}], kind="text")
    return _parse_json_array(content)


def _recognize_text(text, subject_hint="", log=None, stream_cb=None):
    """长文本分段识别：按题号边界切段 → 逐段调用 → 合并结果。
    返回 (data, 段数, 失败段数, 最后一次错误信息)。
    单段失败自动重试一次，仍失败则跳过该段继续。
    stream_cb(items, idx, total) 在每段成功解析后回调，便于 UI 实时入库/显示进度。"""
    prefix, blocks = _split_question_blocks(text)
    chunks = _chunk_blocks(blocks)
    data, failed, last_err = [], 0, ""
    for idx, chunk in enumerate(chunks, 1):
        if log:
            log("正在识别第 %d/%d 段（约 %d 题）……" % (idx, len(chunks), chunk.count("\n\n") + 1))
        items = None
        try:
            items = _call_chunk(chunk, prefix, subject_hint)
        except Exception as e:
            last_err = str(e)
        if items is None:
            try:
                items = _call_chunk(chunk, prefix, subject_hint, strict=True)
            except Exception as e:
                last_err = str(e)
                failed += 1
                if log:
                    log("第 %d/%d 段识别失败，已跳过" % (idx, len(chunks)))
                continue
        # 立即回调：让调用方实时入库/刷新 UI，不等所有段都完成
        if stream_cb and items:
            try:
                stream_cb(items, idx, len(chunks))
            except Exception as e:
                if log:
                    log("第 %d 段回调异常：%s" % (idx, e))
        data.extend(items)
    return data, len(chunks), failed, last_err


def _recognize_images(img_paths, subject_hint="", log=None, stream_cb=None):
    """扫描件/图片型PDF：逐页调用视觉模型并合并（单页请求更稳）。
    自动对过大的图进行 PIL 缩压；失败时按 70%/50% 自动重试更小尺寸，
    仍失败则明确标注跳过此页。"""
    if not api_configured("vision"):
        raise ValueError("未配置视觉 API。请在教师端“API设置” → 视觉 API 区域填写"
                         "base_url、API Key 与多模态模型名（如 qwen-vl-max、gpt-4o、glm-4v-plus 等）。")
    # 入口处统一再压一次（默认更小尺寸，适配 GLM-4V 这类 token 敏感模型）
    norm = []
    for p in img_paths:
        try:
            r = _prep_image_for_vision(p, max_side=1100, max_bytes=700_000)
            norm.append(r[0] if r else p)
        except Exception:
            norm.append(p)
    data, failed, last_err = [], 0, ""
    total = len(norm)
    for idx, p in enumerate(norm, 1):
        if log:
            log("正在识别第 %d/%d 页（图片）……" % (idx, total))
        items, err = _try_recognize_one_page(p, idx, total, subject_hint, log)
        if items is None:
            failed += 1
            if log:
                log("第 %d 页识别失败，已跳过" % idx)
            last_err = "第 %d 页：%s" % (idx, err or "未知错误")
            continue
        if stream_cb and items:
            try:
                stream_cb(items, idx, total)
            except Exception as e:
                if log:
                    log("第 %d 页回调异常：%s" % (idx, e))
        data.extend(items)
    return data, total, failed, last_err


def _try_recognize_one_page(p, idx, total, subject_hint, log):
    """对单页图片最多尝试 3 次：原始 -> 70% -> 50%。返回 (items, err_str)。"""
    attempts = [
        (1100, 700_000, 80),       # 原始：<=1100px、<=700KB、JPEG q=80
        (800,  500_000, 75),       # 70%：<=800px、<=500KB、q=75
        (640,  350_000, 70),       # 50%：<=640px、<=350KB、q=70
    ]
    last_err = ""
    for attempt_idx, (max_side, max_bytes, q) in enumerate(attempts):
        try:
            tmp = _prep_image_for_vision(p, max_side=max_side, max_bytes=max_bytes,
                                         quality=q, force=True)
            if not tmp:
                tmp = [p]
            img_path = tmp[0]
            with open(img_path, "rb") as f:
                raw = f.read()
            b64 = base64.b64encode(raw).decode("ascii")
            # 严格去换行（部分服务对 base64 中的换行敏感）
            b64 = b64.replace("\n", "").replace("\r", "")
            content_parts = [
                {"type": "text", "text": _PROMPT +
                    "\n这是试卷第 %d/%d 页，请抽取本页全部题目。\n只输出 JSON 数组，不要其他文字。" % (idx, total)},
                {"type": "image_url",
                 "image_url": {"url": "data:image/jpeg;base64," + b64}}
            ]
            content = _chat([{"role": "user", "content": content_parts}],
                            kind="vision", max_tokens=4000)
            return _parse_json_array(content), ""
        except Exception as e:
            last_err = str(e)
            # 400/413/图像过大 → 进入下一次更小重试
            if "400" in last_err or "413" in last_err or "image" in last_err.lower() or "too large" in last_err.lower() or "token" in last_err.lower():
                if log:
                    log("第 %d 页第 %d 次重试（原图过大）……" % (idx, attempt_idx + 2))
                continue
            # 其他错误（鉴权/地址错）不必重试
            break
    return None, last_err


# ---------------- 演示模式（未配置API） ----------------
_OPT_PAT = re.compile(r"([A-D])[\.、．]\s*(.+?)(?=\s+[A-D][\.、．]|$)")


def _demo_parse(text, source):
    """按题号切分文本，生成待校对草稿（置信度低，强制人工完善）。"""
    blocks = re.split(r"(?m)^\s*(?:题\s*|第\s*)?(\d{1,3})(?:\s*[\.、．：:]|\s*题)", text)
    items = []
    # blocks: [前缀, '1', 内容, '2', 内容, ...]
    for i in range(1, len(blocks) - 1, 2):
        body = blocks[i + 1].strip()
        if len(body) < 5:
            continue
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]
        stem_lines, opts = [], []
        for ln in lines:
            # 选项可一行一个，也可多个挤在同一行（如 A. xx  B. xx）
            if re.match(r"^[A-D][\.、．]", ln):
                opts.extend(m.group(2).strip() for m in _OPT_PAT.finditer(ln))
            elif not opts:
                stem_lines.append(ln)
        items.append({
            "subject": "数学", "chapter": "未分章", "knowledge_point": "待标注",
            "qtype": "单选" if opts else "填空",
            "difficulty": 3, "stem": "\n".join(stem_lines)[:500],
            "options": opts[:6], "answer": "",
            "analysis": "（演示模式：未配置大模型API，请教师补全答案与解析）",
            "confidence": 0.4, "source": source,
            "tag": _derive_tag({}, source),
        })
        if len(items) >= 50:
            break
    return items


# ---------------- 主入口 ----------------
def recognize(file_path, subject_hint="", progress_cb=None, stream_cb=None):
    """
    识别试卷文件，返回 (items, message)。
    items: 题目 dict 列表（含 confidence），message: 过程说明。
    stream_cb(items, idx, total)：可选，每段成功识别后立即回调，便于 UI 实时入库/显示进度。
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)

    source = os.path.basename(file_path)
    log("正在解析文件：%s" % source)
    kind, payload = extract_content(file_path)

    if not api_configured("text"):
        log("未配置大模型 API，进入演示模式（按题号切分，需人工补全）")
        if kind == "images":
            raise ValueError("扫描件必须配置大模型 API（含视觉模型）才能识别。\n"
                             "请先在“API设置”中配置 api_base / api_key / 视觉模型。")
        items = _demo_parse(payload, source)
        if not items:
            raise ValueError("未能从文件中切分出题目（演示模式需要带题号的文本，如 1. 2. 3.）")
        return items, "演示模式识别出 %d 道草稿题（未调用大模型，请逐题校对补全）" % len(items)

    log("正在调用大模型 API 识别题目……")
    if kind == "text":
        data, n_seg, n_fail, last_err = _recognize_text(
            payload, subject_hint, log=log, stream_cb=stream_cb)
    else:
        data, n_seg, n_fail, last_err = _recognize_images(
            payload, subject_hint, log=log, stream_cb=stream_cb)
    items = _normalize_items(data, source)
    # 按题干去重（防止个别段重复返回）
    seen, uniq = set(), []
    for it in items:
        key = it["stem"][:80]
        if key not in seen:
            seen.add(key)
            uniq.append(it)
    items = uniq
    if not items:
        detail = ("最后一次错误：%s。" % last_err) if last_err else ""
        raise ValueError("大模型未能识别出有效题目（共 %d 段均失败）。%s"
                         "请检查 API 配置（地址/Key/模型名）后重试。" % (n_seg, detail))
    low = sum(1 for it in items if it["confidence"] < 0.7)
    msg = "大模型分 %d 段识别，共识别出 %d 道题" % (n_seg, len(items))
    if n_fail:
        msg += "（%d 段识别失败已跳过，可单独重传）" % n_fail
    if low:
        msg += "，其中 %d 道置信度较低（<0.7），请重点校对" % low
    return items, msg
