/* student.js — 学生端五个标签页 + 答题全屏视图
   知识点训练 / 模拟卷 / 真题卷 / 错题本 / 学情看板 */
"use strict";

/* ================= 答题视图（核心） ================= */
const AnswerView = {
  questions: [], idx: 0, mode: "", title: "",
  results: [],           // {qid, correct: true/false/"self-right"/"self-wrong"}
  qStart: 0, answered: false, lastRes: null,

  /**
   * questions: 题目数组；opts.title 标题；opts.mode "retest" 复测模式
   * 每题可带 _retest_kp（复测来源知识点，随作答请求回传）
   */
  start(questions, opts) {
    opts = opts || {};
    this.questions = questions || [];
    this.idx = 0; this.results = [];
    this.mode = opts.mode || "";
    this.title = opts.title || "练习";
    if (!this.questions.length) { toast("没有可练习的题目", "err"); return; }
    showAnswerScreen();
    this.renderQuestion();
  },

  exit() {
    showMainScreen();
    StudentHome.refresh();
  },

  renderQuestion() {
    const q = this.questions[this.idx];
    this.answered = false;
    this.lastRes = null;
    this.qStart = Date.now();
    const root = document.getElementById("answer-root");
    root.innerHTML = "";

    // 顶部：进度 + 退出
    const head = h("div", "answer-head");
    head.appendChild(h("span", "prog",
      this.title + " · 第 " + (this.idx + 1) + "/" + this.questions.length + " 题"));
    const bar = h("div", "prog-bar");
    const fill = h("div");
    fill.style.width = Math.round(100 * this.idx / this.questions.length) + "%";
    bar.appendChild(fill);
    head.appendChild(bar);
    if (this.mode === "retest") head.appendChild(h("span", "badge badge-blue", "复测"));
    const exitBtn = h("button", "btn btn-ghost btn-sm", "退出");
    exitBtn.onclick = () => { if (confirm("确定退出本次练习？已作答成绩会保留。")) this.exit(); };
    head.appendChild(exitBtn);
    root.appendChild(head);

    const body = h("div", "answer-body");
    root.appendChild(body);

    // 题目卡片（不含选项列表，选项做成可点按钮）
    const card = h("div", "card");
    this.qCard = h("div");
    renderQuestionInto(this.qCard, q, { showAnswer: false, hideOptions: true });
    card.appendChild(this.qCard);
    body.appendChild(card);

    // 作答区
    this.inputCard = h("div", "card");
    body.appendChild(this.inputCard);
    this.renderInputArea(q);

    // 结果区（提交后填充）
    this.resultCard = h("div");
    body.appendChild(this.resultCard);
  },

  renderInputArea(q) {
    const box = this.inputCard;
    box.innerHTML = "";
    const qt = q.qtype;
    this.selSingle = "";
    this.selMulti = new Set();

    if (qt === "单选" || qt === "多选") {
      const opts = Array.isArray(q.options) ? q.options : [];
      box.appendChild(h("div", "muted small mb16",
        qt === "单选" ? "请选择一个选项：" : "请选择所有正确选项（多选）："));
      opts.forEach((o, i) => {
        const letter = String.fromCharCode(65 + i);
        const btn = h("button", "opt-btn");
        const lb = h("span", "opt-letter", letter);
        const tx = h("span", "opt-text");
        renderRich(tx, o);
        btn.appendChild(lb); btn.appendChild(tx);
        btn.onclick = () => {
          if (this.answered) return;
          if (qt === "单选") {
            this.selSingle = letter;
            box.querySelectorAll(".opt-btn").forEach(b => b.classList.remove("sel"));
            btn.classList.add("sel");
          } else {
            if (this.selMulti.has(letter)) { this.selMulti.delete(letter); btn.classList.remove("sel"); }
            else { this.selMulti.add(letter); btn.classList.add("sel"); }
          }
        };
        box.appendChild(btn);
      });
    } else if (qt === "填空") {
      box.appendChild(h("div", "muted small mb16", "请输入答案："));
      this.fillInput = h("input");
      this.fillInput.type = "text";
      this.fillInput.style.width = "100%";
      this.fillInput.placeholder = "在此填写答案";
      this.fillInput.addEventListener("keydown", e => {
        if (e.key === "Enter") submitBtn.click();
      });
      box.appendChild(this.fillInput);
    } else {
      box.appendChild(h("div", "muted small mb16",
        "解答题：请在下方作答，提交后对照标准答案自评。"));
      this.textArea = h("textarea");
      this.textArea.rows = 6;
      this.textArea.style.width = "100%";
      this.textArea.placeholder = "写下你的解题过程…";
      box.appendChild(this.textArea);
    }

    const row = h("div", "row mt16");
    const submitBtn = h("button", "btn btn-lg", "提交");
    row.appendChild(submitBtn);
    box.appendChild(row);
    submitBtn.onclick = () => this.submit(q);
  },

  currentAnswer(q) {
    if (q.qtype === "单选") return this.selSingle;
    if (q.qtype === "多选") return Array.from(this.selMulti).sort().join("");
    if (q.qtype === "填空") return this.fillInput.value.trim();
    return this.textArea.value;
  },

  async submit(q) {
    if (this.answered) return;
    const ans = this.currentAnswer(q);
    if (q.qtype !== "解答" && !ans) { toast("请先作答再提交", "err"); return; }
    const dur = Math.round((Date.now() - this.qStart) / 1000);
    const payload = { qid: q.qid, answer: ans, duration_sec: dur, mode: this.mode };
    if (this.mode === "retest" && q._retest_kp) payload.kp = q._retest_kp;
    let res;
    try { res = await api("/api/answer", { method: "POST", body: payload }); }
    catch (e) { toast(e.message, "err"); return; }
    this.answered = true;
    this.lockInputs();
    if (res.correct === null) {
      // 解答题：展示答案解析 + 自评
      this.renderSelfGrade(q, res, ans, dur);
    } else {
      this.results.push({ qid: q.qid, correct: !!res.correct });
      this.renderResult(q, res, false);
    }
  },

  lockInputs() {
    this.inputCard.querySelectorAll("button.opt-btn").forEach(b => b.disabled = true);
    this.inputCard.querySelectorAll("input,textarea").forEach(i => i.disabled = true);
    const sb = this.inputCard.querySelector(".btn-lg");
    if (sb) sb.disabled = true;
  },

  renderSelfGrade(q, res, ans, dur) {
    const box = this.resultCard;
    box.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "参考答案与解析"));
    const aq = Object.assign({}, q, { answer: res.std_answer, analysis: res.analysis });
    const holder = h("div");
    // 只渲染答案/解析部分
    if (String(res.std_answer || "").trim()) {
      const ans2 = h("div", "q-ans");
      ans2.appendChild(h("b", null, "答案："));
      const sp = h("span"); renderRich(sp, res.std_answer); ans2.appendChild(sp);
      holder.appendChild(ans2);
    }
    if (String(res.analysis || "").trim()) {
      const ana = h("div", "q-ana");
      ana.appendChild(h("b", null, "解析："));
      const sp = h("span"); renderRich(sp, res.analysis); ana.appendChild(sp);
      holder.appendChild(ana);
    }
    card.appendChild(holder);
    card.appendChild(h("div", "muted mt8", "请对照标准答案，诚实自评："));
    const row = h("div", "row mt8");
    const right = h("button", "btn btn-green", "我答对了");
    const wrong = h("button", "btn btn-danger", "我答错了");
    row.appendChild(right); row.appendChild(wrong);
    card.appendChild(row);
    box.appendChild(card);

    const doSelf = async (correct) => {
      right.disabled = wrong.disabled = true;
      const payload = { qid: q.qid, correct, answer: ans, duration_sec: dur, mode: this.mode };
      if (this.mode === "retest" && q._retest_kp) payload.kp = q._retest_kp;
      let r2;
      try { r2 = await api("/api/answer/self", { method: "POST", body: payload }); }
      catch (e) { toast(e.message, "err"); right.disabled = wrong.disabled = false; return; }
      this.results.push({ qid: q.qid, correct });
      card.remove();
      this.renderResult(q, r2, true);
    };
    right.onclick = () => doSelf(true);
    wrong.onclick = () => doSelf(false);
  },

  renderResult(q, res, selfGraded) {
    const box = this.resultCard;
    box.innerHTML = "";
    const card = h("div", "card");
    const banner = h("div", "result-banner");
    const icon = h("div", "result-icon " + (res.correct ? "right" : "wrong"),
                   res.correct ? "✔" : "✘");
    const info = h("div");
    info.appendChild(h("div", null,
      (res.correct ? "回答正确！" : "回答错误") + (selfGraded ? "（自评）" : "")));
    const ml = h("div", "mastery-line");
    const stCls = res.status === "已过关" ? "badge-green"
                : res.status === "未掌握" ? "badge-red" : "badge-yellow";
    ml.innerHTML = "";
    ml.appendChild(document.createTextNode(
      "知识点「" + res.kp + "」掌握度 " + res.mastery_before + " → "));
    const nb = h("b", null, String(res.mastery_after));
    ml.appendChild(nb);
    ml.appendChild(document.createTextNode(" "));
    const stBadge = h("span", "badge " + stCls, res.status + (res.passed ? " 🎉" : ""));
    ml.appendChild(stBadge);
    info.appendChild(ml);
    banner.appendChild(icon); banner.appendChild(info);
    card.appendChild(banner);

    if (!selfGraded) {
      if (String(res.std_answer || "").trim()) {
        const ans = h("div", "q-ans");
        ans.appendChild(h("b", null, "答案："));
        const sp = h("span"); renderRich(sp, res.std_answer); ans.appendChild(sp);
        card.appendChild(ans);
      }
      if (String(res.analysis || "").trim()) {
        const ana = h("div", "q-ana");
        ana.appendChild(h("b", null, "解析："));
        const sp = h("span"); renderRich(sp, res.analysis); ana.appendChild(sp);
        card.appendChild(ana);
      }
    }

    const row = h("div", "row mt16");
    const simBtn = h("button", "btn btn-ghost", "再出类似题");
    const nextBtn = h("button", "btn",
      this.idx + 1 < this.questions.length ? "下一题" : "完成");
    row.appendChild(simBtn); row.appendChild(nextBtn);
    card.appendChild(row);
    box.appendChild(card);
    card.scrollIntoView({ behavior: "smooth", block: "start" });

    simBtn.onclick = () => withBtn(simBtn, "查找中…", async () => {
      const qs = new URLSearchParams({ kp: res.kp, exclude_qid: q.qid });
      if (q.subject) qs.set("subject", q.subject);
      const sim = await api("/api/similar?" + qs.toString());
      if (!sim) {
        simBtn.disabled = true;
        simBtn.textContent = "暂无更多类似题";
        return;
      }
      if (this.mode === "retest" && q._retest_kp) sim._retest_kp = q._retest_kp;
      this.questions[this.idx] = sim;
      this.renderQuestion();
      toast("已换成同知识点的类似题", "ok");
    }).catch(e => toast(e.message, "err"));

    nextBtn.onclick = () => {
      if (this.idx + 1 < this.questions.length) {
        this.idx++;
        this.renderQuestion();
        window.scrollTo({ top: 0 });
      } else {
        this.renderSummary();
      }
    };
  },

  renderSummary() {
    const root = document.getElementById("answer-root");
    root.innerHTML = "";
    const body = h("div", "answer-body");
    const card = h("div", "card");
    card.style.textAlign = "center";
    card.appendChild(h("h3", null, "本次练习完成！"));
    const right = this.results.filter(r => r.correct).length;
    const wrong = this.results.length - right;
    const acc = this.results.length ? Math.round(100 * right / this.results.length) : 0;
    const stats = h("div", "stat-cards mt16");
    stats.appendChild(this._stat(this.results.length, "总题数"));
    stats.appendChild(this._stat(right, "答对"));
    stats.appendChild(this._stat(wrong, "答错"));
    stats.appendChild(this._stat(acc + "%", "正确率"));
    card.appendChild(stats);
    const back = h("button", "btn btn-lg mt16", "返回首页");
    back.onclick = () => this.exit();
    card.appendChild(h("div"));
    card.appendChild(back);
    body.appendChild(card);
    root.appendChild(body);
    window.scrollTo({ top: 0 });
  },

  _stat(num, label) {
    const c = h("div", "stat-card");
    c.appendChild(h("div", "num", String(num)));
    c.appendChild(h("div", "lbl", label));
    return c;
  }
};

/* ================= 1. 知识点训练 ================= */
const StuKpView = {
  async render(root) {
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "知识点训练"));
    card.appendChild(h("div", "muted small mb16",
      "选择科目 → 章节 → 知识点，系统会调出该知识点的全部题目。"));
    const tree = Meta.kpTree();
    const sSel = makeSelect(Object.keys(tree), "", "请选择科目");
    const cSel = makeSelect([], "", "请选择章节");
    const kSel = makeSelect([], "", "请选择知识点");
    cSel.disabled = true; kSel.disabled = true;
    sSel.onchange = () => {
      const chapters = tree[sSel.value] ? Object.keys(tree[sSel.value]) : [];
      cSel.innerHTML = ""; cSel.appendChild(new Option("请选择章节", ""));
      chapters.forEach(c => cSel.appendChild(new Option(c, c)));
      cSel.disabled = !chapters.length;
      kSel.innerHTML = ""; kSel.appendChild(new Option("请选择知识点", ""));
      kSel.disabled = true;
    };
    cSel.onchange = () => {
      const kps = (tree[sSel.value] && tree[sSel.value][cSel.value]) || [];
      kSel.innerHTML = ""; kSel.appendChild(new Option("请选择知识点", ""));
      kps.forEach(k => kSel.appendChild(new Option(k, k)));
      kSel.disabled = !kps.length;
    };
    const row = h("div", "row");
    row.appendChild(sSel); row.appendChild(cSel); row.appendChild(kSel);
    const startBtn = h("button", "btn", "开始训练");
    row.appendChild(startBtn);
    card.appendChild(row);
    root.appendChild(card);

    startBtn.onclick = () => withBtn(startBtn, "加载中…", async () => {
      if (!sSel.value || !cSel.value || !kSel.value) {
        toast("请完整选择科目、章节与知识点", "err"); return;
      }
      const qs = await api("/api/practice/kp", { method: "POST", body: {
        subject: sSel.value, chapter: cSel.value, kp: kSel.value } });
      if (!qs.length) { toast("该知识点暂无题目", "err"); return; }
      AnswerView.start(qs, { title: "知识点训练 · " + kSel.value });
    }).catch(e => toast(e.message, "err"));
  }
};

/* ================= 2. 模拟卷 ================= */
const StuMockView = {
  async render(root) {
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "模拟卷组卷"));
    card.appendChild(h("div", "muted small mb16",
      "按目标难度智能配比抽题（相邻难度各占 25%，目标难度占 50%）。"));
    const row = h("div", "row");
    row.appendChild(h("span", null, "科目："));
    const sSel = makeSelect(Meta.subjects(), "数学");
    row.appendChild(sSel);
    row.appendChild(h("span", null, "难度："));
    const range = h("input");
    range.type = "range"; range.min = "1"; range.max = "5"; range.value = "3";
    const dLbl = h("b", null, "3");
    range.oninput = () => dLbl.textContent = range.value;
    row.appendChild(range); row.appendChild(dLbl);
    row.appendChild(h("span", null, "题量："));
    const cSel = makeSelect(["5", "10", "15", "20"], "10");
    row.appendChild(cSel);
    const startBtn = h("button", "btn", "开始组卷");
    row.appendChild(startBtn);
    card.appendChild(row);
    root.appendChild(card);

    startBtn.onclick = () => withBtn(startBtn, "组卷中…", async () => {
      const qs = await api("/api/practice/mock", { method: "POST", body: {
        subject: sSel.value, difficulty: parseInt(range.value, 10),
        count: parseInt(cSel.value, 10) } });
      AnswerView.start(qs, { title: "模拟卷 · " + sSel.value });
    }).catch(e => toast(e.message, "err"));
  }
};

/* ================= 3. 真题卷 ================= */
const StuPaperView = {
  async render(root) {
    root.innerHTML = "";
    const card = h("div", "card");
    card.appendChild(h("h3", null, "真题卷"));
    const listBox = h("div");
    card.appendChild(listBox);
    root.appendChild(card);
    let papers;
    try { papers = await api("/api/papers"); }
    catch (e) { toast(e.message, "err"); return; }
    if (!papers.length) { listBox.appendChild(h("div", "empty-tip", "暂无试卷")); return; }
    papers.forEach(p => {
      const row = h("div", "paper-card card");
      const info = h("div", "info");
      info.appendChild(h("div", "title", p.title));
      info.appendChild(h("div", "muted small",
        p.subject + " · " + (p.year || "") + " · " + (p.ptype || "") +
        " · 共 " + p.q_count + " 题"));
      const startBtn = h("button", "btn", "开始作答");
      row.appendChild(info); row.appendChild(startBtn);
      listBox.appendChild(row);
      startBtn.onclick = () => withBtn(startBtn, "加载中…", async () => {
        const r = await api("/api/papers/" + encodeURIComponent(p.pid) + "/questions");
        if (!r.questions.length) { toast("该试卷暂无题目", "err"); return; }
        AnswerView.start(r.questions, { title: p.title });
      }).catch(e => toast(e.message, "err"));
    });
  }
};

/* ================= 4. 错题本 ================= */
const StuWrongView = {
  items: [],
  async render(root) {
    root.innerHTML = "";
    const card = h("div", "card");
    const head = h("div", "row");
    head.appendChild(h("h3", null, "错题本"));
    head.appendChild(h("div", "grow"));
    const redoBtn = h("button", "btn", "重做这些题");
    head.appendChild(redoBtn);
    card.appendChild(head);
    this.listBox = h("div", "mt8");
    card.appendChild(this.listBox);
    root.appendChild(card);

    let items;
    try { items = await api("/api/wrong"); }
    catch (e) { toast(e.message, "err"); return; }
    this.items = items;
    redoBtn.disabled = !items.length;
    if (!items.length) {
      this.listBox.appendChild(h("div", "empty-tip", "暂无错题，继续保持！"));
      return;
    }
    items.forEach(q => {
      const row = h("div", "card");
      row.style.padding = "14px 16px";
      const badges = h("div", "q-badges");
      [q.subject, q.knowledge_point, q.qtype].filter(Boolean).forEach(t =>
        badges.appendChild(h("span", "badge", t)));
      if (q.media_type && q.media_type !== "无")
        badges.appendChild(h("span", "badge badge-media", "🖼 " + q.media_type));
      row.appendChild(badges);
      row.appendChild(h("div", null, truncate(q.stem, 60)));
      row.appendChild(h("div", "muted small mt8",
        "答错时间：" + (q.wrong_ts || "未知") +
        (q.my_answer ? " · 我的答案：" + truncate(q.my_answer, 20) : "")));
      this.listBox.appendChild(row);
    });
    redoBtn.onclick = () => {
      AnswerView.start(this.items, { title: "错题重做" });
    };
  }
};

/* ================= 5. 学情看板 ================= */
const StuDashView = {
  async render(root) {
    root.innerHTML = "<div class='empty-tip loading-dot'>加载中</div>";
    let overall, kps, retests;
    try {
      [overall, kps, retests] = await Promise.all([
        api("/api/analytics/overall"),
        api("/api/analytics/kp"),
        api("/api/practice/retest")
      ]);
    } catch (e) { root.innerHTML = ""; toast(e.message, "err"); return; }
    root.innerHTML = "";

    // 四张统计卡
    const cards = h("div", "stat-cards");
    const mk = (num, lbl) => {
      const c = h("div", "stat-card");
      c.appendChild(h("div", "num", String(num)));
      c.appendChild(h("div", "lbl", lbl));
      return c;
    };
    cards.appendChild(mk(overall.total_answers, "总答题数"));
    cards.appendChild(mk(overall.accuracy + "%", "正确率"));
    cards.appendChild(mk(overall.kp_passed + "/" + overall.kp_touched, "已过关知识点"));
    cards.appendChild(mk(overall.retest_due, "待复测"));
    root.appendChild(cards);

    // 待复测区
    if (retests.length) {
      const rc = h("div", "card");
      rc.appendChild(h("h3", null, "待复测知识点（按遗忘曲线到期）"));
      retests.forEach(it => {
        const badge = h("span", "badge badge-blue",
          it.kp + "（当前掌握度 " + it.mastery + "）");
        rc.appendChild(badge);
      });
      const startBtn = h("button", "btn mt8", "一键开始复测");
      rc.appendChild(h("div"));
      rc.appendChild(startBtn);
      root.appendChild(rc);
      startBtn.onclick = () => {
        const qs = retests.map(it => {
          const q = it.question;
          q._retest_kp = it.kp;
          return q;
        });
        AnswerView.start(qs, { title: "遗忘曲线复测", mode: "retest" });
      };
    }

    // 掌握度表格
    const card = h("div", "card");
    card.appendChild(h("h3", null, "知识点掌握度（最薄弱排前）"));
    if (!kps.length) {
      card.appendChild(h("div", "empty-tip", "还没有练习记录，去「知识点训练」开始吧"));
    } else {
      const wrap = h("div", "tbl-wrap");
      const tbl = h("table", "tbl");
      tbl.innerHTML = "<thead><tr><th>知识点</th><th>科目</th><th>掌握度</th>" +
        "<th>状态</th><th>练题数</th><th>正确率</th></tr></thead>";
      const tbody = h("tbody");
      kps.forEach(k => {
        const tr = h("tr");
        tr.appendChild(h("td", null, k.kp));
        tr.appendChild(h("td", null, k.subject));
        const tdM = h("td");
        const mCls = k.mastery >= 80 ? "m-green" : k.mastery >= 40 ? "m-yellow" : "m-red";
        const bar = h("span", "mastery-bar");
        const fill = h("div", mCls);
        fill.style.width = Math.round(k.mastery) + "%";
        bar.appendChild(fill);
        tdM.appendChild(bar);
        tdM.appendChild(document.createTextNode(" " + k.mastery));
        tr.appendChild(tdM);
        const stCls = k.status === "已过关" ? "badge-green"
                    : k.status === "未掌握" ? "badge-red" : "badge-yellow";
        const tdS = h("td");
        tdS.appendChild(h("span", "badge " + stCls, k.status));
        tr.appendChild(tdS);
        tr.appendChild(h("td", null, String(k.total)));
        tr.appendChild(h("td", null, k.accuracy + "%"));
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      wrap.appendChild(tbl);
      card.appendChild(wrap);
    }
    root.appendChild(card);
  }
};

const StudentViews = {
  kp: StuKpView, mock: StuMockView, papers: StuPaperView,
  wrong: StuWrongView, dash: StuDashView
};

/* 学生端各标签重新进入时的刷新（答题退出后调用） */
const StudentHome = {
  current: null,
  refresh() {
    if (this.current && typeof this.current.render === "function") {
      const page = document.getElementById("page");
      if (page) this.current.render(page);
    }
  }
};
