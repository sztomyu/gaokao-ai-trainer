/* render.js — 富文本渲染管线（LaTeX / mhchem / 题目配图原位渲染）+ 通用 UI 组件
   这是 Web 版改造的核心：题干中的数学公式、化学方程式、配图全部在浏览器端渲染。 */
"use strict";

/* ---------------- 基础工具 ---------------- */
function escapeHtml(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

/* 图片文件名白名单：中英文/数字/下划线/横线/点（docx 提取图可能含中文名） */
const IMG_NAME_RE = /^[\w.\-一-鿿]+$/;

function imgUrl(fname) {
  return "/media/" + encodeURIComponent(fname);
}

/* 截断摘要 */
function truncate(s, n) {
  s = String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  return s.length > (n || 60) ? s.slice(0, n || 60) + "…" : s;
}

/* ---------------- 富文本渲染管线 ---------------- */
/*
 * renderRich(container, text) 五步：
 * 1. 整段 HTML 转义
 * 2. [IMG:文件名] → <img class="q-img">
 * 3. 数界定界符之外的 \ce{...} 自动包成 $\ce{...}$（mhchem 化学方程式）
 * 4. 写入 innerHTML 后调 KaTeX auto-render
 * 5. 换行由容器 CSS white-space: pre-wrap 保留
 */

/* 数界定界符切分正则：$$...$$ / \(...\) / \[...\] / $...$ */
const _MATH_SEG_RE = /(\$\$[\s\S]+?\$\$|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$[^$\n]+?\$)/g;
const _CE_RE = /\\ce\{((?:[^{}]|\{[^{}]*\})*)\}/g;

function _wrapCeOutsideMath(text) {
  return text.split(_MATH_SEG_RE).map((seg, i) => {
    if (i % 2 === 1) return seg;            // 数界定界符内部，原样保留
    return seg.replace(_CE_RE, m => "$" + m + "$");
  }).join("");
}

function renderRich(container, text) {
  if (!container) return;
  let t = escapeHtml(text == null ? "" : text);
  // [IMG:文件名] → 图片标签（白名单校验防目录穿越）
  t = t.replace(/\[IMG:([^\]\s]+)\]/g, (m, fname) => {
    if (!IMG_NAME_RE.test(fname) || fname.indexOf("..") >= 0) return m;
    return '<img class="q-img" src="' + imgUrl(fname) + '" alt="题目配图">';
  });
  // \ce{...} → $\ce{...}$（仅数界定界符之外）
  t = _wrapCeOutsideMath(t);
  container.innerHTML = t;
  if (window.renderMathInElement) {
    try {
      renderMathInElement(container, {
        delimiters: [
          { left: "$$", right: "$$", display: true },
          { left: "\\(", right: "\\)", display: false },
          { left: "\\[", right: "\\]", display: true },
          { left: "$", right: "$", display: false }
        ],
        throwOnError: false
      });
    } catch (e) { console.warn("KaTeX 渲染异常：", e); }
  }
}

/* ---------------- 题目卡片渲染 ---------------- */
/**
 * renderQuestionInto(el, q, {showAnswer, hideOptions})
 * 徽章行 → 题干(富文本) → media_paths 图片(去重) → 选项 → 答案/解析(可选)
 */
function renderQuestionInto(el, q, opts) {
  opts = opts || {};
  const showAnswer = !!opts.showAnswer;
  el.innerHTML = "";

  // 徽章行
  const badges = document.createElement("div");
  badges.className = "q-badges";
  [q.subject, q.chapter, q.knowledge_point,
   q.qtype, "难度 " + (q.difficulty || 3), q.tag]
    .filter(x => x && String(x).trim())
    .forEach(t => {
      const b = document.createElement("span");
      b.className = "badge";
      b.textContent = t;
      badges.appendChild(b);
    });
  if (q.media_type && q.media_type !== "无") {
    const b = document.createElement("span");
    b.className = "badge badge-media";
    b.textContent = "🖼 " + q.media_type;
    badges.appendChild(b);
  }
  // 直接检测是否含图片/图形
  const hasMedia = (q.media_paths && q.media_paths.length) ||
                   (q.has_figure === true) ||
                   (q.has_figure === "true") ||
                   (q.has_figure === 1) ||
                   (q.has_figure === "1");
  if (hasMedia) {
    const b = document.createElement("span");
    b.className = "badge badge-figure";
    b.textContent = "📐 含图";
    badges.appendChild(b);
  }
  el.appendChild(badges);

  // 题干
  const stem = document.createElement("div");
  stem.className = "q-stem";
  renderRich(stem, q.stem || "");
  el.appendChild(stem);

  // media_paths 图片（题干中已含同名 [IMG:] 则跳过，防重复展示）
  const stemText = String(q.stem || "");
  (q.media_paths || []).forEach(fname => {
    if (!fname || stemText.indexOf("[IMG:" + fname + "]") >= 0) return;
    if (!IMG_NAME_RE.test(fname) || fname.indexOf("..") >= 0) return;
    const img = document.createElement("img");
    img.className = "q-img";
    img.src = imgUrl(fname);
    img.alt = "题目配图";
    el.appendChild(img);
  });

  // 选项列表
  const options = Array.isArray(q.options) ? q.options : [];
  if (options.length && !opts.hideOptions) {
    const ol = document.createElement("ol");
    ol.className = "q-opts";
    ol.type = "A";
    options.forEach(o => {
      const li = document.createElement("li");
      renderRich(li, o);
      ol.appendChild(li);
    });
    el.appendChild(ol);
  }

  // 答案与解析
  if (showAnswer) {
    if (String(q.answer || "").trim()) {
      const ans = document.createElement("div");
      ans.className = "q-ans";
      const lb = document.createElement("b");
      lb.textContent = "答案：";
      ans.appendChild(lb);
      const span = document.createElement("span");
      renderRich(span, q.answer);
      ans.appendChild(span);
      el.appendChild(ans);
    }
    if (String(q.analysis || "").trim()) {
      const ana = document.createElement("div");
      ana.className = "q-ana";
      const lb = document.createElement("b");
      lb.textContent = "解析：";
      ana.appendChild(lb);
      const span = document.createElement("span");
      renderRich(span, q.analysis);
      ana.appendChild(span);
      el.appendChild(ana);
    }
  }
}

/* ---------------- Toast ---------------- */
function toast(msg, type, ms) {
  const root = document.getElementById("toast-root");
  const d = document.createElement("div");
  d.className = "toast " + (type || "");
  d.textContent = msg;
  root.appendChild(d);
  setTimeout(() => { d.style.opacity = "0"; d.style.transition = "opacity .3s"; }, (ms || 2600) - 300);
  setTimeout(() => d.remove(), ms || 2600);
}

/* ---------------- 弹窗 ---------------- */
function openModal(opt) {
  const root = document.getElementById("modal-root");
  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";
  const modal = document.createElement("div");
  modal.className = "modal" + (opt.wide ? " wide" : "");
  const head = document.createElement("div");
  head.className = "modal-head";
  const title = document.createElement("span");
  title.textContent = opt.title || "";
  const x = document.createElement("button");
  x.className = "modal-x";
  x.textContent = "×";
  head.appendChild(title); head.appendChild(x);
  const bodyEl = document.createElement("div");
  bodyEl.className = "modal-body";
  if (typeof opt.body === "string") bodyEl.innerHTML = opt.body;
  else if (opt.body) bodyEl.appendChild(opt.body);
  modal.appendChild(head); modal.appendChild(bodyEl);
  let foot = null;
  if (opt.foot) {
    foot = document.createElement("div");
    foot.className = "modal-foot";
    modal.appendChild(foot);
  }
  backdrop.appendChild(modal);
  root.appendChild(backdrop);
  const close = () => { backdrop.remove(); if (opt.onClose) opt.onClose(); };
  x.onclick = close;
  backdrop.addEventListener("mousedown", e => { if (e.target === backdrop) close(); });
  return { close, body: bodyEl, foot };
}

/* ---------------- 表单辅助 ---------------- */
function h(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

function fieldRow(label, input) {
  const wrap = h("div");
  const lb = h("label", "f-label", label);
  wrap.appendChild(lb);
  wrap.appendChild(input);
  return wrap;
}

function makeSelect(options, value, withEmpty) {
  const sel = document.createElement("select");
  if (withEmpty) {
    const o = document.createElement("option");
    o.value = ""; o.textContent = withEmpty === true ? "（全部）" : withEmpty;
    sel.appendChild(o);
  }
  (options || []).forEach(v => {
    const o = document.createElement("option");
    o.value = v; o.textContent = v;
    sel.appendChild(o);
  });
  if (value !== undefined && value !== null) sel.value = value;
  return sel;
}

/* 按钮 loading 态包装：执行期间禁用并显示“处理中” */
async function withBtn(btn, text, fn) {
  const old = btn.textContent;
  btn.disabled = true;
  btn.textContent = text || "处理中…";
  try { return await fn(); }
  finally { btn.disabled = false; btn.textContent = old; }
}
