/* teacher.js — 教师端五个标签页：试卷导入 / 题库管理 / 数据看板 / 知识点 / 账号与API */
"use strict";

const MEDIA_TYPES = ["无", "数学公式", "函数图像", "几何图形", "化学方程式",
                     "分子结构式", "物理示意图", "其他图表"];

const CAT_COLORS = { "选择题": "#3182CE", "填空题": "#38A169",
                     "阅读题": "#DD6B20", "大题": "#805AD5", "其他": "#718096" };

/* ================= 题目编辑弹窗（导入结果 & 题库管理共用） ================= */
function openQuestionEditModal(q, onSave, title) {
  const bodyEl = h("div");
  const inputs = {};

  const subjInput = h("input");
  subjInput.type = "text"; subjInput.value = q.subject || "";
  subjInput.setAttribute("list", "dl-subjects");
  const dl = h("datalist"); dl.id = "dl-subjects";
  Meta.subjects().forEach(s => { const o = h("option"); o.value = s; dl.appendChild(o); });
  bodyEl.appendChild(dl);

  inputs.subject = subjInput;
  inputs.chapter = h("input"); inputs.chapter.type = "text"; inputs.chapter.value = q.chapter || "";
  inputs.kp = h("input"); inputs.kp.type = "text"; inputs.kp.value = q.knowledge_point || "";
  inputs.qtype = makeSelect(["单选", "多选", "填空", "解答"], q.qtype || "单选");
  inputs.difficulty = makeSelect(["1", "2", "3", "4", "5"], String(q.difficulty || 3));
  inputs.stem = h("textarea"); inputs.stem.rows = 5; inputs.stem.value = q.stem || "";
  inputs.stem.style.width = "100%";
  inputs.options = h("textarea"); inputs.options.rows = 4;
  inputs.options.value = (Array.isArray(q.options) ? q.options : []).join("\n");
  inputs.options.placeholder = "每行一个选项（不含 A. B. C. 标号）；非选择题留空";
  inputs.options.style.width = "100%";
  inputs.answer = h("input"); inputs.answer.type = "text"; inputs.answer.value = q.answer || "";
  inputs.analysis = h("textarea"); inputs.analysis.rows = 4;
  inputs.analysis.value = q.analysis || ""; inputs.analysis.style.width = "100%";
  inputs.tag = h("input"); inputs.tag.type = "text"; inputs.tag.value = q.tag || "";
  inputs.mediaType = makeSelect(MEDIA_TYPES, q.media_type || "无");
  inputs.mediaPaths = h("input"); inputs.mediaPaths.type = "text";
  inputs.mediaPaths.value = (q.media_paths || []).join(",");
  inputs.mediaPaths.placeholder = "多个文件名用英文逗号分隔";

  const grid1 = h("div", "row");
  const w1 = h("div", "grow"); w1.appendChild(fieldRow("科目", inputs.subject));
  const w2 = h("div", "grow"); w2.appendChild(fieldRow("章节", inputs.chapter));
  grid1.appendChild(w1); grid1.appendChild(w2);
  bodyEl.appendChild(grid1);
  bodyEl.appendChild(fieldRow("知识点", inputs.kp));
  const grid2 = h("div", "row");
  const w3 = h("div", "grow"); w3.appendChild(fieldRow("题型", inputs.qtype));
  const w4 = h("div", "grow"); w4.appendChild(fieldRow("难度", inputs.difficulty));
  const w5 = h("div", "grow"); w5.appendChild(fieldRow("标签", inputs.tag));
  grid2.appendChild(w3); grid2.appendChild(w4); grid2.appendChild(w5);
  bodyEl.appendChild(grid2);
  bodyEl.appendChild(fieldRow("题干（支持 LaTeX 公式、\\ce{化学方程式}、[IMG:文件名]）", inputs.stem));
  bodyEl.appendChild(fieldRow("选项", inputs.options));
  bodyEl.appendChild(fieldRow("答案", inputs.answer));
  bodyEl.appendChild(fieldRow("解析", inputs.analysis));
  const grid3 = h("div", "row");
  const w6 = h("div", "grow"); w6.appendChild(fieldRow("图形类型", inputs.mediaType));
  grid3.appendChild(w6);
  bodyEl.appendChild(grid3);
  bodyEl.appendChild(fieldRow("配图文件（media_paths）", inputs.mediaPaths));

  // media_paths 缩略图实时预览
  const thumbs = h("div", "thumb-row");
  const refreshThumbs = () => {
    thumbs.innerHTML = "";
    inputs.mediaPaths.value.split(",").map(s => s.trim()).filter(Boolean)
      .forEach(f => {
        if (!IMG_NAME_RE.test(f) || f.indexOf("..") >= 0) return;
        const img = h("img");
        img.src = imgUrl(f); img.title = f;
        thumbs.appendChild(img);
      });
  };
  inputs.mediaPaths.addEventListener("input", refreshThumbs);
  refreshThumbs();
  bodyEl.appendChild(thumbs);

  const m = openModal({ title: title || "编辑题目", body: bodyEl, wide: true, foot: true });
  const cancel = h("button", "btn btn-ghost", "取消");
  cancel.onclick = m.close;
  const save = h("button", "btn", "保存");
  save.onclick = async () => {
    const updated = Object.assign({}, q, {
      subject: inputs.subject.value.trim(),
      chapter: inputs.chapter.value.trim(),
      knowledge_point: inputs.kp.value.trim(),
      qtype: inputs.qtype.value,
      difficulty: parseInt(inputs.difficulty.value, 10) || 3,
      stem: inputs.stem.value,
      options: inputs.options.value.split("\n").map(s => s.trim()).filter(Boolean),
      answer: inputs.answer.value,
      analysis: inputs.analysis.value,
      tag: inputs.tag.value.trim(),
      media_type: inputs.mediaType.value,
      media_paths: inputs.mediaPaths.value.split(",").map(s => s.trim()).filter(Boolean)
    });
    if (!updated.stem.trim()) { toast("题干不能为空", "err"); return; }
    save.disabled = true;
    try { await onSave(updated); m.close(); }
    catch (e) { toast(e.message, "err"); }
    finally { save.disabled = false; }
  };
  m.foot.appendChild(cancel); m.foot.appendChild(save);
}

/* 题目预览弹窗 */
function openQuestionPreview(q, title) {
  const bodyEl = h("div");
  renderQuestionInto(bodyEl, q, { showAnswer: true });
  openModal({ title: title || "题目预览", body: bodyEl, wide: true });
}

/* ================= 1. 试卷导入 ================= */
const ImportView = {
  fileId: "", filename: "", jobId: "", timer: null, items: [],

  async render(root) {
    this.stopPoll();
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "试卷导入（docx / pdf / 图片 / txt）"));

    const row = h("div", "row");
    const fileInput = h("input");
    fileInput.type = "file";
    fileInput.accept = ".docx,.doc,.pdf,.png,.jpg,.jpeg,.bmp,.gif,.webp,.txt";
    const upBtn = h("button", "btn btn-ghost", "上传文件");
    const upStatus = h("span", "muted small", "未上传");
    row.appendChild(fileInput); row.appendChild(upBtn); row.appendChild(upStatus);

    upBtn.onclick = () => withBtn(upBtn, "上传中…", async () => {
      if (!fileInput.files.length) { toast("请先选择文件", "err"); return; }
      const fd = new FormData();
      fd.append("file", fileInput.files[0]);
      const r = await api("/api/import/upload", { method: "POST", body: fd });
      this.fileId = r.file_id; this.filename = r.filename;
      upStatus.textContent = "已上传：" + r.filename;
      upStatus.className = "small";
      toast("上传成功", "ok");
      startBtn.disabled = false;
    }).catch(e => toast(e.message, "err"));

    const row2 = h("div", "row mt8");
    row2.appendChild(h("span", null, "科目提示（可选）："));
    const hintSel = makeSelect(Meta.subjects(), "", "（不指定）");
    row2.appendChild(hintSel);
    const startBtn = h("button", "btn", "开始识别");
    startBtn.disabled = true;
    row2.appendChild(startBtn);
    const stopNote = h("span", "muted small",
      "识别在后台进行，大试卷可能需要几分钟，可切换标签稍后回来查看");
    row2.appendChild(stopNote);

    card.appendChild(row); card.appendChild(row2);

    // 进度日志
    this.logBox = h("div", "log-box mt16", "等待开始…");
    this.logBox.style.display = "none";
    card.appendChild(this.logBox);

    // 结果区
    this.resultBox = h("div", "mt16");
    card.appendChild(this.resultBox);
    root.appendChild(card);

    startBtn.onclick = () => withBtn(startBtn, "启动中…", async () => {
      const r = await api("/api/import/recognize", {
        method: "POST",
        body: { file_id: this.fileId, subject_hint: hintSel.value }
      });
      this.jobId = r.job_id;
      this.items = [];
      this.resultBox.innerHTML = "";
      this.logBox.style.display = "block";
      this.logBox.textContent = "";
      this.startPoll();
    }).catch(e => toast(e.message, "err"));

    // 已有进行中的任务则恢复
    if (this.jobId) {
      this.logBox.style.display = "block";
      this.startPoll();
    }
  },

  startPoll() {
    this.stopPoll();
    this.timer = setInterval(() => this.poll(), 1500);
    this.poll();
  },
  stopPoll() { if (this.timer) { clearInterval(this.timer); this.timer = null; } },

  async poll() {
    if (!this.jobId) return;
    let j;
    try { j = await api("/api/import/jobs/" + this.jobId); }
    catch (e) { this.stopPoll(); return; }
    this.logBox.textContent = (j.logs || []).join("\n") ||
      (j.status === "running" ? "识别中…" : "");
    this.logBox.scrollTop = this.logBox.scrollHeight;
    if (j.status === "done") {
      this.stopPoll();
      this.items = j.items || [];
      this.logBox.textContent += "\n" + (j.message || "识别完成");
      this.renderResult();
      toast(j.message || "识别完成", "ok", 4000);
    } else if (j.status === "error") {
      this.stopPoll();
      this.logBox.textContent += "\n识别失败：" + (j.error || "未知错误");
      toast("识别失败：" + (j.error || "未知错误"), "err", 5000);
    }
  },

  renderResult() {
    const box = this.resultBox;
    box.innerHTML = "";
    if (!this.items.length) { box.appendChild(h("div", "empty-tip", "未识别出题目")); return; }

    const head = h("div", "row mb16");
    head.appendChild(h("b", null, "识别结果（共 " + this.items.length + " 题）"));
    const selAll = h("button", "btn btn-ghost btn-sm", "全选");
    const selNone = h("button", "btn btn-ghost btn-sm", "全不选");
    const commitBtn = h("button", "btn btn-accent", "写入选中题入库");
    head.appendChild(selAll); head.appendChild(selNone); head.appendChild(commitBtn);
    box.appendChild(head);

    const wrap = h("div", "tbl-wrap");
    const tbl = h("table", "tbl");
    tbl.innerHTML = "<thead><tr><th></th><th>#</th><th>科目</th><th>题型</th>" +
      "<th>知识点</th><th>难度</th><th>置信度</th><th>图形</th><th>题干摘要</th><th>操作</th></tr></thead>";
    const tbody = h("tbody");
    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    box.appendChild(wrap);

    const checks = [];
    this.items.forEach((it, i) => {
      const tr = h("tr");
      const tdCk = h("td");
      const ck = h("input"); ck.type = "checkbox"; ck.checked = true;
      checks.push(ck); tdCk.appendChild(ck);
      tr.appendChild(tdCk);
      tr.appendChild(h("td", null, String(i + 1)));
      tr.appendChild(h("td", null, it.subject || ""));
      tr.appendChild(h("td", null, it.qtype || ""));
      tr.appendChild(h("td", null, it.knowledge_point || ""));
      tr.appendChild(h("td", null, String(it.difficulty || 3)));
      const conf = (it.confidence == null ? 0.8 : Number(it.confidence));
      const tdConf = h("td", conf < 0.7 ? "conf-low" : "conf-ok", conf.toFixed(2));
      if (conf < 0.7) tdConf.title = "置信度较低，请重点校对";
      tr.appendChild(tdConf);
      const mt = it.media_type && it.media_type !== "无" ? "🖼" + it.media_type : "—";
      tr.appendChild(h("td", null, mt));
      const tdStem = h("td", null, truncate(it.stem, 60));
      tdStem.style.maxWidth = "320px";
      tr.appendChild(tdStem);
      const tdOp = h("td");
      const pv = h("button", "btn btn-ghost btn-sm", "预览");
      const ed = h("button", "btn btn-ghost btn-sm", "编辑");
      const del = h("button", "btn btn-danger btn-sm", "删除");
      pv.onclick = () => openQuestionPreview(this.items[i], "识别预览 #" + (i + 1));
      ed.onclick = () => openQuestionEditModal(this.items[i], async (upd) => {
        this.items[i] = upd;
        this.renderResult();
        toast("已修改（尚未入库）", "ok");
      }, "编辑识别结果 #" + (i + 1));
      del.onclick = () => {
        if (!confirm("确认删除第 " + (i + 1) + " 题？")) return;
        this.items.splice(i, 1);
        this.renderResult();
      };
      tdOp.appendChild(pv); tdOp.appendChild(ed); tdOp.appendChild(del);
      tr.appendChild(tdOp);
      tbody.appendChild(tr);
    });

    selAll.onclick = () => checks.forEach(c => c.checked = true);
    selNone.onclick = () => checks.forEach(c => c.checked = false);
    commitBtn.onclick = () => withBtn(commitBtn, "入库中…", async () => {
      const selected = this.items.filter((_, i) => checks[i] && checks[i].checked);
      if (!selected.length) { toast("请先勾选要入库的题目", "err"); return; }
      const r = await api("/api/import/commit", {
        method: "POST", body: { items: selected }
      });
      toast("入库 " + r.new_qids.length + " 题，跳过重复 " + r.skipped + " 题", "ok", 5000);
      Meta.load(true);
    }).catch(e => toast(e.message, "err"));
  }
};

/* ================= 2. 题库管理 ================= */
const BankView = {
  page: 1, size: 20,

  async render(root) {
    this.page = 1;
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "题库管理"));

    const filters = h("div", "row");
    this.fSubject = makeSelect(Meta.subjects(), "", "全部科目");
    this.fQtype = makeSelect(Meta.qtypes(), "", "全部题型");
    this.fTag = h("input"); this.fTag.type = "text"; this.fTag.placeholder = "标签";
    this.fTag.style.width = "140px";
    this.fKeyword = h("input"); this.fKeyword.type = "text";
    this.fKeyword.placeholder = "题干关键词"; this.fKeyword.style.width = "180px";
    const searchBtn = h("button", "btn", "查询");
    this.totalSpan = h("span", "muted");
    filters.appendChild(this.fSubject); filters.appendChild(this.fQtype);
    filters.appendChild(this.fTag); filters.appendChild(this.fKeyword);
    filters.appendChild(searchBtn); filters.appendChild(this.totalSpan);
    card.appendChild(filters);

    this.tblBox = h("div", "mt16");
    card.appendChild(this.tblBox);
    this.pagerBox = h("div", "pager");
    card.appendChild(this.pagerBox);
    root.appendChild(card);

    searchBtn.onclick = () => { this.page = 1; this.load(); };
    this.fKeyword.addEventListener("keydown", e => {
      if (e.key === "Enter") { this.page = 1; this.load(); }
    });
    await this.load();
  },

  async load() {
    this.tblBox.innerHTML = "<div class='empty-tip loading-dot'>加载中</div>";
    const qs = new URLSearchParams({
      subject: this.fSubject.value, qtype: this.fQtype.value,
      tag: this.fTag.value.trim(), keyword: this.fKeyword.value.trim(),
      page: this.page, size: this.size
    });
    let r;
    try { r = await api("/api/questions?" + qs.toString()); }
    catch (e) { this.tblBox.innerHTML = ""; toast(e.message, "err"); return; }
    this.totalSpan.textContent = "共 " + r.total + " 题";
    this.renderTable(r);
  },

  renderTable(r) {
    const box = this.tblBox;
    box.innerHTML = "";
    if (!r.items.length) { box.appendChild(h("div", "empty-tip", "没有符合条件的题目")); }
    else {
      const wrap = h("div", "tbl-wrap");
      const tbl = h("table", "tbl");
      tbl.innerHTML = "<thead><tr><th>qid</th><th>科目</th><th>章节</th><th>知识点</th>" +
        "<th>题型</th><th>难度</th><th>图形</th><th>标签</th><th>题干摘要</th><th>操作</th></tr></thead>";
      const tbody = h("tbody");
      r.items.forEach(q => {
        const tr = h("tr");
        tr.appendChild(h("td", null, q.qid));
        tr.appendChild(h("td", null, q.subject));
        tr.appendChild(h("td", null, q.chapter));
        const kpTd = h("td", null, q.knowledge_point);
        kpTd.style.maxWidth = "160px";
        tr.appendChild(kpTd);
        tr.appendChild(h("td", null, q.qtype));
        tr.appendChild(h("td", null, String(q.difficulty)));
        tr.appendChild(h("td", null,
          q.media_type && q.media_type !== "无" ? "🖼" + q.media_type
            : (q.media_count > 0 ? "🖼×" + q.media_count : "—")));
        tr.appendChild(h("td", null, q.tag || ""));
        const st = h("td", null, truncate(q.stem, 60));
        st.style.maxWidth = "300px";
        tr.appendChild(st);
        const op = h("td");
        op.style.whiteSpace = "nowrap";
        const pv = h("button", "btn btn-ghost btn-sm", "预览");
        const ed = h("button", "btn btn-ghost btn-sm", "编辑");
        const del = h("button", "btn btn-danger btn-sm", "删除");
        pv.onclick = () => openQuestionPreview(q, "题目预览 " + q.qid);
        ed.onclick = () => openQuestionEditModal(q, async (upd) => {
          await api("/api/questions/" + encodeURIComponent(q.qid),
                    { method: "PUT", body: upd });
          toast("已保存", "ok");
          this.load();
        }, "编辑题目 " + q.qid);
        del.onclick = async () => {
          if (!confirm("确认删除题目 " + q.qid + "？此操作不可恢复。")) return;
          try {
            await api("/api/questions/" + encodeURIComponent(q.qid), { method: "DELETE" });
            toast("已删除", "ok");
            this.load();
          } catch (e) { toast(e.message, "err"); }
        };
        op.appendChild(pv); op.appendChild(ed); op.appendChild(del);
        tr.appendChild(op);
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      wrap.appendChild(tbl);
      box.appendChild(wrap);
    }
    // 分页
    const pages = Math.max(1, Math.ceil(r.total / this.size));
    this.pagerBox.innerHTML = "";
    const prev = h("button", "btn btn-ghost btn-sm", "上一页");
    const next = h("button", "btn btn-ghost btn-sm", "下一页");
    prev.disabled = this.page <= 1;
    next.disabled = this.page >= pages;
    prev.onclick = () => { this.page--; this.load(); };
    next.onclick = () => { this.page++; this.load(); };
    this.pagerBox.appendChild(prev);
    this.pagerBox.appendChild(h("span", null, "第 " + this.page + " / " + pages + " 页"));
    this.pagerBox.appendChild(next);
  }
};

/* ================= 3. 数据看板 ================= */
const DashView = {
  charts: [],
  async render(root) {
    this.charts.forEach(c => { try { c.dispose(); } catch (e) {} });
    this.charts = [];
    root.innerHTML = "<div class='empty-tip loading-dot'>加载中</div>";
    let r;
    try { r = await api("/api/analytics/category"); }
    catch (e) { root.innerHTML = ""; toast(e.message, "err"); return; }
    root.innerHTML = "";

    // 各科题量卡片
    const cards = h("div", "stat-cards");
    cards.appendChild(this.statCard(r.total, "题库总量"));
    r.by_subject.filter(s => s.total > 0).forEach(s =>
      cards.appendChild(this.statCard(s.total, s.subject)));
    root.appendChild(cards);

    // 堆叠柱状图 + 饼图
    const row = h("div", "row");
    row.style.alignItems = "stretch";
    const c1 = h("div", "card grow"); c1.style.minWidth = "420px";
    c1.appendChild(h("h3", null, "各科目题型分布"));
    const bar = h("div", "chart-box"); c1.appendChild(bar);
    const c2 = h("div", "card grow"); c2.style.minWidth = "320px";
    c2.appendChild(h("h3", null, "题型分类占比"));
    const pie = h("div", "chart-box"); c2.appendChild(pie);
    row.appendChild(c1); row.appendChild(c2);
    root.appendChild(row);

    const subjects = r.by_subject.filter(s => s.total > 0);
    const cats = ["选择题", "填空题", "阅读题", "大题"];
    const barChart = echarts.init(bar);
    barChart.setOption({
      tooltip: { trigger: "axis" },
      legend: { bottom: 0 },
      grid: { left: 40, right: 20, top: 30, bottom: 50 },
      xAxis: { type: "category", data: subjects.map(s => s.subject) },
      yAxis: { type: "value" },
      series: cats.map(c => ({
        name: c, type: "bar", stack: "总量",
        itemStyle: { color: CAT_COLORS[c] },
        data: subjects.map(s => s.categories[c] || 0)
      }))
    });
    const pieChart = echarts.init(pie);
    pieChart.setOption({
      tooltip: { trigger: "item" },
      legend: { bottom: 0 },
      series: [{
        type: "pie", radius: ["35%", "65%"],
        label: { formatter: "{b}: {c}" },
        data: Object.keys(r.by_category)
          .filter(k => r.by_category[k] > 0)
          .map(k => ({ name: k, value: r.by_category[k],
                       itemStyle: { color: CAT_COLORS[k] || "#999" } }))
      }]
    });
    this.charts.push(barChart, pieChart);

    // 标签云
    const tagCard = h("div", "card");
    tagCard.appendChild(h("h3", null, "标签云（按科目）"));
    subjects.forEach(s => {
      if (!s.tags || !s.tags.length) return;
      const sec = h("div", "mb16");
      sec.appendChild(h("div", null, s.subject + "（" + s.tags.length + " 个标签）"));
      const cloud = h("div", "mt8");
      s.tags.forEach(t => cloud.appendChild(h("span", "tag-chip", t)));
      sec.appendChild(cloud);
      tagCard.appendChild(sec);
    });
    root.appendChild(tagCard);
  },
  statCard(num, label) {
    const c = h("div", "stat-card");
    c.appendChild(h("div", "num", String(num)));
    c.appendChild(h("div", "lbl", label));
    return c;
  }
};

/* ================= 4. 知识点 ================= */
const KpView = {
  async render(root) {
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "知识点树"));
    this.treeBox = h("div");
    card.appendChild(this.treeBox);

    const formCard = h("div", "card");
    formCard.appendChild(h("h3", null, "新增知识点"));
    const row = h("div", "row");
    const subj = h("input"); subj.type = "text"; subj.placeholder = "科目（如 数学）";
    subj.setAttribute("list", "dl-kp-subjects");
    const dl = h("datalist"); dl.id = "dl-kp-subjects";
    Meta.subjects().forEach(s => { const o = h("option"); o.value = s; dl.appendChild(o); });
    const chap = h("input"); chap.type = "text"; chap.placeholder = "章节（如 函数与导数）";
    const kp = h("input"); kp.type = "text"; kp.placeholder = "知识点（如 导数的几何意义）";
    const addBtn = h("button", "btn", "添加");
    row.appendChild(subj); row.appendChild(dl); row.appendChild(chap);
    row.appendChild(kp); row.appendChild(addBtn);
    formCard.appendChild(row);
    root.appendChild(formCard);
    root.appendChild(card);

    addBtn.onclick = () => withBtn(addBtn, "添加中…", async () => {
      await api("/api/kp", { method: "POST", body: {
        subject: subj.value.trim(), chapter: chap.value.trim(),
        knowledge_point: kp.value.trim() } });
      toast("已添加", "ok");
      kp.value = "";
      await Meta.load(true);
      this.loadTree();
    }).catch(e => toast(e.message, "err"));

    await this.loadTree();
  },

  async loadTree() {
    let tree;
    try { tree = await api("/api/kp-tree"); }
    catch (e) { toast(e.message, "err"); return; }
    const box = this.treeBox;
    box.innerHTML = "";
    const subjects = Object.keys(tree);
    if (!subjects.length) { box.appendChild(h("div", "empty-tip", "暂无知识点")); return; }
    subjects.forEach(s => {
      const sec = h("div", "kp-subject");
      sec.appendChild(h("h4", null, s));
      Object.keys(tree[s]).forEach(c => {
        const ch = h("div", "kp-chapter");
        ch.appendChild(h("span", "cname", "▸ " + c));
        const list = h("div", "kp-list");
        tree[s][c].forEach(k => list.appendChild(h("span", "badge", k)));
        ch.appendChild(list);
        sec.appendChild(ch);
      });
      box.appendChild(sec);
    });
  }
};

/* ================= 5. 账号与 API ================= */
const AccountView = {
  async render(root) {
    root.innerHTML = "";
    // 账号管理
    const card = h("div", "card");
    const head = h("div", "row");
    head.appendChild(h("h3", null, "账号管理"));
    head.appendChild(h("div", "grow"));
    const addBtn = h("button", "btn btn-ghost", "新增账号");
    head.appendChild(addBtn);
    card.appendChild(head);
    this.userBox = h("div", "mt8");
    card.appendChild(this.userBox);
    root.appendChild(card);
    addBtn.onclick = () => this.editUser(null);

    // API 设置
    const apiCard = h("div", "card");
    apiCard.appendChild(h("h3", null, "大模型 API 设置"));
    this.apiBox = h("div", "api-grid");
    apiCard.appendChild(this.apiBox);
    apiCard.appendChild(h("div", "muted small mt8",
      "文本 API 用于 Word / 文本型 PDF 识别；视觉 API 用于扫描件与图片型试卷。"));
    root.appendChild(apiCard);

    await Promise.all([this.loadUsers(), this.loadApi()]);
  },

  async loadUsers() {
    let users;
    try { users = await api("/api/users"); }
    catch (e) { toast(e.message, "err"); return; }
    const box = this.userBox;
    box.innerHTML = "";
    const tbl = h("table", "tbl");
    tbl.innerHTML = "<thead><tr><th>用户名</th><th>角色</th><th>姓名</th>" +
      "<th>状态</th><th>操作</th></tr></thead>";
    const tbody = h("tbody");
    users.forEach(u => {
      const tr = h("tr");
      tr.appendChild(h("td", null, u.username));
      tr.appendChild(h("td", null, u.role === "teacher" ? "教师" : "学生"));
      tr.appendChild(h("td", null, u.name));
      const st = h("span", "badge " + (u.status === "启用" ? "badge-green" : "badge-red"),
                   u.status);
      const tdSt = h("td"); tdSt.appendChild(st); tr.appendChild(tdSt);
      const op = h("td");
      const ed = h("button", "btn btn-ghost btn-sm", "编辑");
      const tg = h("button", "btn btn-ghost btn-sm", u.status === "启用" ? "停用" : "启用");
      const rp = h("button", "btn btn-ghost btn-sm", "重置密码");
      ed.onclick = () => this.editUser(u);
      tg.onclick = async () => {
        try {
          const r = await api("/api/users/toggle",
                              { method: "POST", body: { username: u.username } });
          toast("已" + r.status, "ok");
          this.loadUsers();
        } catch (e) { toast(e.message, "err"); }
      };
      rp.onclick = () => {
        const pwd = prompt("为 " + u.username + " 设置新密码：");
        if (!pwd) return;
        api("/api/users/save", { method: "POST", body: {
          username: u.username, password: pwd, role: u.role,
          name: u.name, status: u.status } })
          .then(() => toast("密码已重置", "ok"))
          .catch(e => toast(e.message, "err"));
      };
      op.appendChild(ed); op.appendChild(tg); op.appendChild(rp);
      tr.appendChild(op);
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
    box.appendChild(tbl);
  },

  editUser(u) {
    const isNew = !u;
    u = u || { username: "", role: "student", name: "", status: "启用" };
    const bodyEl = h("div");
    const un = h("input"); un.type = "text"; un.value = u.username;
    un.disabled = !isNew;
    const nm = h("input"); nm.type = "text"; nm.value = u.name;
    const role = makeSelect([["student", "学生"], ["teacher", "教师"]].map(x => x[0]),
                            u.role);
    // 用中文显示
    role.innerHTML = "";
    [["student", "学生"], ["teacher", "教师"]].forEach(([v, t]) => {
      const o = h("option"); o.value = v; o.textContent = t; role.appendChild(o);
    });
    role.value = u.role;
    const st = makeSelect(["启用", "停用"], u.status);
    const pwd = h("input"); pwd.type = "password";
    pwd.placeholder = isNew ? "默认 123456（留空即用默认）" : "留空则不修改密码";
    bodyEl.appendChild(fieldRow("用户名", un));
    bodyEl.appendChild(fieldRow("姓名", nm));
    bodyEl.appendChild(fieldRow("角色", role));
    bodyEl.appendChild(fieldRow("状态", st));
    bodyEl.appendChild(fieldRow("密码", pwd));
    const m = openModal({ title: isNew ? "新增账号" : "编辑账号", body: bodyEl, foot: true });
    const cancel = h("button", "btn btn-ghost", "取消");
    cancel.onclick = m.close;
    const save = h("button", "btn", "保存");
    save.onclick = () => withBtn(save, "保存中…", async () => {
      await api("/api/users/save", { method: "POST", body: {
        username: un.value.trim(), password: pwd.value,
        role: role.value, name: nm.value.trim(), status: st.value } });
      toast("已保存", "ok");
      m.close();
      this.loadUsers();
    }).catch(e => toast(e.message, "err"));
    m.foot.appendChild(cancel); m.foot.appendChild(save);
  },

  async loadApi() {
    let cfg;
    try { cfg = await api("/api/settings/api"); }
    catch (e) { toast(e.message, "err"); return; }
    const box = this.apiBox;
    box.innerHTML = "";
    [["text", "文本 API（题目识别）"], ["vision", "视觉 API（扫描件识别）"]]
      .forEach(([kind, title]) => box.appendChild(this.apiCard(kind, title, cfg[kind])));
  },

  apiCard(kind, title, c) {
    const card = h("div", "card");
    card.style.marginBottom = "0";
    card.appendChild(h("h3", null, title));
    const base = h("input"); base.type = "text"; base.value = c.api_base || "";
    base.style.width = "100%";
    const path = h("input"); path.type = "text"; path.value = c.path || "";
    path.style.width = "100%";
    const key = h("input"); key.type = "password";
    key.placeholder = c.has_key ? "已配置（" + c.api_key_mask + "），留空不修改" : "未配置";
    key.style.width = "100%";
    const model = h("input"); model.type = "text"; model.value = c.model || "";
    model.style.width = "100%";
    card.appendChild(fieldRow("API 地址（api_base）", base));
    card.appendChild(fieldRow("接口路径（path）", path));
    card.appendChild(fieldRow("API Key", key));
    card.appendChild(fieldRow("模型名（model）", model));
    const row = h("div", "row mt8");
    const testBtn = h("button", "btn btn-ghost btn-sm", "测试连接");
    const saveBtn = h("button", "btn btn-sm", "保存");
    const result = h("span", "test-result");
    row.appendChild(testBtn); row.appendChild(saveBtn); row.appendChild(result);
    card.appendChild(row);

    testBtn.onclick = () => withBtn(testBtn, "测试中…", async () => {
      result.textContent = "";
      const r = await api("/api/settings/test", { method: "POST", body: { kind } });
      result.textContent = r.message;
      result.className = "test-result " + (r.success ? "ok" : "err");
    }).catch(e => { result.textContent = e.message; result.className = "test-result err"; });

    saveBtn.onclick = () => withBtn(saveBtn, "保存中…", async () => {
      await api("/api/settings/api", { method: "POST", body: {
        [kind]: { api_base: base.value, path: path.value,
                  api_key: key.value, model: model.value } } });
      toast("已保存", "ok");
      key.value = "";
      this.loadApi();
    }).catch(e => toast(e.message, "err"));
    return card;
  }
};

const TeacherViews = {
  import: ImportView, bank: BankView, dash: DashView,
  kp: KpView, account: AccountView
};
