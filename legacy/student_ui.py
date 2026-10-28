# -*- coding: utf-8 -*-
"""
student_ui.py — 学生端界面
知识点训练 · 模拟卷训练（难度系数） · 高考真题 · 错题本 · 学情看板（含复测）
"""
import random
import tkinter as tk
from tkinter import ttk, messagebox

import excel_store as store
import engine
import study_loop
import analytics

FONT = ("Microsoft YaHei UI", 10)
FONT_B = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 12, "bold")


# ================= 答题面板（三种训练模式共用） =================
class QuizFrame(ttk.Frame):
    """按题目列表逐题作答：即时判分 → 解析 → 相似题 → 过关提示。"""

    def __init__(self, parent, username, questions, title, on_close=None):
        super().__init__(parent)
        self.username = username
        self.questions = list(questions)
        self.title = title
        self.on_close = on_close
        self.idx = 0
        self.right_count = 0
        self.asked_ids = []
        self._build()
        self._show_question()

    # ---------- UI ----------
    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=6)
        ttk.Label(top, text=self.title, font=FONT_TITLE).pack(side="left")
        self.lbl_prog = ttk.Label(top, text="", font=FONT)
        self.lbl_prog.pack(side="right")

        self.lbl_tag = ttk.Label(self, text="", font=FONT, foreground="#2B6CB0")
        self.lbl_tag.pack(anchor="w", padx=10)

        self.txt_stem = tk.Text(self, height=6, wrap="word", font=FONT,
                                bg="#FAFAFA", relief="solid", bd=1)
        self.txt_stem.pack(fill="x", padx=10, pady=4)
        self.txt_stem.config(state="disabled")

        self.frm_imgs = tk.Frame(self)      # 题目图片展示区
        self.frm_imgs.pack(fill="x", padx=10)
        self._img_refs = []

        self.frm_answer = ttk.Frame(self)
        self.frm_answer.pack(fill="x", padx=10, pady=4)

        self.btn_submit = ttk.Button(self, text="提交", command=self._submit)
        self.btn_submit.pack(pady=4)

        self.frm_feedback = ttk.Frame(self)
        self.frm_feedback.pack(fill="both", expand=True, padx=10, pady=4)

    def _clear(self, frame):
        for w in frame.winfo_children():
            w.destroy()

    def _show_question_images(self):
        """显示题目关联图片（data/media/），并提供 HTML 富媒体预览（公式渲染）。"""
        self._clear(self.frm_imgs)
        self._img_refs = []
        media_paths = store.media_paths_of(self.q)
        if not media_paths:
            return
        from PIL import Image, ImageTk
        for fname in media_paths[:3]:
            p = fname if os.path.isabs(fname) else os.path.join(store.media_dir(), fname)
            if not os.path.exists(p):
                continue
            try:
                im = Image.open(p)
                im.thumbnail((360, 260))
                photo = ImageTk.PhotoImage(im)
                self._img_refs.append(photo)
                tk.Label(self.frm_imgs, image=photo, bd=1, relief="solid",
                         bg="white").pack(side="left", padx=4, pady=4)
            except Exception:
                pass
        ttk.Button(self.frm_imgs, text="🖼 高清预览(公式渲染)",
                   command=self._open_html_preview).pack(side="right", padx=6)

    def _open_html_preview(self):
        import html_preview
        html_preview.preview_questions([self.q], title="题目预览")

    def _show_question(self):
        if self.idx >= len(self.questions):
            self._finish()
            return
        self.q = self.questions[self.idx]
        self.asked_ids.append(self.q["qid"])
        self._clear(self.frm_answer)
        self._clear(self.frm_feedback)
        self.btn_submit.config(state="normal", text="提交")

        self.lbl_prog.config(text="第 %d / %d 题    已答对 %d 题"
                                  % (self.idx + 1, len(self.questions), self.right_count))
        self.lbl_tag.config(text="【%s】%s · %s · 难度%s"
                                 % (self.q["qtype"], self.q["subject"],
                                    self.q["knowledge_point"], self.q["difficulty"]))
        self.txt_stem.config(state="normal")
        self.txt_stem.delete("1.0", "end")
        self.txt_stem.insert("1.0", self.q["stem"])
        self.txt_stem.config(state="disabled")
        self._show_question_images()

        self.options = store.parse_options(self.q)
        qt = self.q["qtype"]
        if qt == "单选":
            self.var = tk.StringVar(value="")
            for i, opt in enumerate(self.options):
                ttk.Radiobutton(self.frm_answer, text="%s. %s" % (chr(65 + i), opt),
                                variable=self.var, value=chr(65 + i),
                                style="TRadiobutton").pack(anchor="w", pady=2)
        elif qt == "多选":
            self.vars = []
            for i, opt in enumerate(self.options):
                v = tk.StringVar(value="")
                self.vars.append((v, chr(65 + i)))
                ttk.Checkbutton(self.frm_answer, text="%s. %s" % (chr(65 + i), opt),
                                variable=v, onvalue=chr(65 + i),
                                offvalue="").pack(anchor="w", pady=2)
        elif qt == "填空":
            ttk.Label(self.frm_answer, text="请输入答案：", font=FONT).pack(anchor="w")
            self.var = tk.StringVar()
            ttk.Entry(self.frm_answer, textvariable=self.var, font=FONT,
                      width=50).pack(anchor="w", pady=4)
        else:  # 解答
            ttk.Label(self.frm_answer, text="请在纸上作答后，点击提交查看解析并自评：",
                      font=FONT).pack(anchor="w")
            self.btn_submit.config(text="提交并查看解析")

    def _user_answer(self):
        qt = self.q["qtype"]
        if qt == "多选":
            return "".join(sorted(v.get() for v, _ in self.vars if v.get()))
        if qt == "解答":
            return "（自评）"
        return self.var.get().strip()

    # ---------- 判分 ----------
    def _submit(self):
        ua = self._user_answer()
        if self.q["qtype"] in ("单选", "填空") and not ua:
            messagebox.showinfo("提示", "请先作答再提交")
            return
        if self.q["qtype"] == "多选" and not ua:
            messagebox.showinfo("提示", "请先选择答案再提交")
            return

        judge = engine.grade(self.q, ua)
        if judge is None:      # 解答题：展示解析，自评
            self._show_analysis_panel(need_self_mark=True, user_answer=ua)
            self.btn_submit.config(state="disabled")
            return
        self._after_judge(bool(judge), ua)

    def _after_judge(self, correct, user_answer):
        self.btn_submit.config(state="disabled")
        result = study_loop.on_answer(self.username, self.q, user_answer, correct=correct)
        if correct:
            self.right_count += 1
        self._show_analysis_panel(result=result, user_answer=user_answer)

    def _show_analysis_panel(self, result=None, need_self_mark=False, user_answer=""):
        self._clear(self.frm_feedback)
        f = self.frm_feedback
        if need_self_mark:
            ttk.Label(f, text="请对照下面的解析进行自评：", font=FONT_B).pack(anchor="w")
        else:
            ok = result["correct"]
            ttk.Label(f, text="✔ 回答正确" if ok else "✘ 回答错误",
                      font=("Microsoft YaHei UI", 11, "bold"),
                      foreground="#2F855A" if ok else "#C53030").pack(anchor="w")
            ttk.Label(f, text="知识点「%s」掌握度：%.0f → %.0f（%s）"
                              % (result["kp"], result["mastery_before"],
                                 result["mastery_after"], result["status"]),
                      font=FONT).pack(anchor="w")
            if result["passed"]:
                ttk.Label(f, text="🎉 该知识点已过关！系统将在遗忘曲线节点提醒你复测。",
                          font=FONT_B, foreground="#2F855A").pack(anchor="w")

        ans = tk.Text(f, height=7, wrap="word", font=FONT, bg="#FFFBEA",
                      relief="solid", bd=1)
        ans.pack(fill="both", expand=True, pady=4)
        ans.insert("1.0", "【参考答案】%s\n\n【详细解析】\n%s"
                          % (self.q["answer"], self.q["analysis"] or "（暂无解析）"))
        ans.config(state="disabled")

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=4)
        if need_self_mark:
            ttk.Button(btns, text="我答对了",
                       command=lambda: self._after_judge(True, user_answer)).pack(side="left", padx=4)
            ttk.Button(btns, text="我答错了",
                       command=lambda: self._after_judge(False, user_answer)).pack(side="left", padx=4)
        else:
            if not result["correct"]:
                ttk.Button(btns, text="再出类似题",
                           command=self._similar).pack(side="left", padx=4)
            ttk.Button(btns, text="下一题", command=self._next).pack(side="left", padx=4)

    def _similar(self):
        kp = str(self.q["knowledge_point"]).split(",")[0].strip()
        q = study_loop.similar_question(self.username, kp,
                                        exclude_qids=self.asked_ids,
                                        subject=self.q["subject"])
        if not q:
            messagebox.showinfo("提示", "该知识点暂无更多题目，可让老师通过“试卷识别录入”补充题库。")
            return
        self.questions.insert(self.idx + 1, q)
        self._next()

    def _next(self):
        self.idx += 1
        self._show_question()

    def _finish(self):
        self._clear(self.frm_answer)
        self._clear(self.frm_feedback)
        self.btn_submit.config(state="disabled")
        n = len(self.questions)
        acc = (100.0 * self.right_count / n) if n else 0
        self.txt_stem.config(state="normal")
        self.txt_stem.delete("1.0", "end")
        self.txt_stem.insert("1.0", "本次训练完成！\n\n共 %d 题，答对 %d 题，正确率 %.1f%%。"
                                    "\n\n可到“学情看板”查看知识点掌握情况。"
                                    % (n, self.right_count, acc))
        self.txt_stem.config(state="disabled")
        self.lbl_prog.config(text="已完成")


# ================= 学生端主窗口 =================
class StudentWindow(tk.Toplevel):
    def __init__(self, master, user):
        super().__init__()
        self.master = master
        self.user = user
        self.username = user["username"]
        self.title("高考学习训练系统（简化版）— 学生端 · %s" % user["name"])
        self.geometry("980x720")
        self.protocol("WM_DELETE_WINDOW", self._quit)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb

        self.tab_kp = ttk.Frame(nb); nb.add(self.tab_kp, text="知识点训练")
        self.tab_mock = ttk.Frame(nb); nb.add(self.tab_mock, text="模拟卷训练")
        self.tab_real = ttk.Frame(nb); nb.add(self.tab_real, text="高考真题")
        self.tab_wrong = ttk.Frame(nb); nb.add(self.tab_wrong, text="错题本")
        self.tab_dash = ttk.Frame(nb); nb.add(self.tab_dash, text="学情看板")

        self._build_kp_tab()
        self._build_mock_tab()
        self._build_real_tab()
        self._build_wrong_tab()
        self._build_dash_tab()
        self.bind("<Visibility>", lambda e: self._refresh_dash())

    def _quit(self):
        self.master.destroy()

    # ---------- 知识点训练 ----------
    def _build_kp_tab(self):
        bar = ttk.LabelFrame(self.tab_kp, text="选择知识点")
        bar.pack(fill="x", padx=10, pady=8)
        self.tree = store.get_kp_tree()
        subjects = list(self.tree.keys())

        self.cb_subject = ttk.Combobox(bar, values=subjects, state="readonly", font=FONT, width=10)
        self.cb_chapter = ttk.Combobox(bar, state="readonly", font=FONT, width=18)
        self.cb_kp = ttk.Combobox(bar, state="readonly", font=FONT, width=24)
        ttk.Label(bar, text="科目", font=FONT).grid(row=0, column=0, padx=4, pady=8)
        self.cb_subject.grid(row=0, column=1, padx=4)
        ttk.Label(bar, text="章节", font=FONT).grid(row=0, column=2, padx=4)
        self.cb_chapter.grid(row=0, column=3, padx=4)
        ttk.Label(bar, text="知识点", font=FONT).grid(row=0, column=4, padx=4)
        self.cb_kp.grid(row=0, column=5, padx=4)
        ttk.Button(bar, text="开始训练", command=self._start_kp).grid(row=0, column=6, padx=10)

        self.cb_subject.bind("<<ComboboxSelected>>", self._on_subject)
        self.cb_chapter.bind("<<ComboboxSelected>>", self._on_chapter)
        if subjects:
            self.cb_subject.current(0)
            self._on_subject()

        self.frm_kp_quiz = ttk.Frame(self.tab_kp)
        self.frm_kp_quiz.pack(fill="both", expand=True)

    def _on_subject(self, event=None):
        s = self.cb_subject.get()
        chapters = list(self.tree.get(s, {}).keys())
        self.cb_chapter.config(values=chapters)
        if chapters:
            self.cb_chapter.current(0)
        self._on_chapter()

    def _on_chapter(self, event=None):
        s, c = self.cb_subject.get(), self.cb_chapter.get()
        kps = self.tree.get(s, {}).get(c, [])
        self.cb_kp.config(values=kps)
        if kps:
            self.cb_kp.current(0)

    def _start_kp(self):
        kp = self.cb_kp.get()
        if not kp:
            messagebox.showinfo("提示", "请先选择知识点")
            return
        qs = store.load_questions(subject=self.cb_subject.get(), kp=kp)
        if not qs:
            messagebox.showinfo("提示", "该知识点下暂无题目，请老师先录入。")
            return
        random.shuffle(qs)
        self._mount_quiz(self.frm_kp_quiz, qs, "知识点训练 · %s" % kp)

    # ---------- 模拟卷 ----------
    def _build_mock_tab(self):
        bar = ttk.LabelFrame(self.tab_mock, text="随机组卷（可设置难度系数 1~5）")
        bar.pack(fill="x", padx=10, pady=8)
        self.cb_mock_subject = ttk.Combobox(bar, values=store.SUBJECTS,
                                            state="readonly", font=FONT, width=8)
        self.cb_mock_subject.set("数学")
        ttk.Label(bar, text="科目", font=FONT).grid(row=0, column=0, padx=4, pady=8)
        self.cb_mock_subject.grid(row=0, column=1, padx=4)
        ttk.Label(bar, text="难度系数", font=FONT).grid(row=0, column=2, padx=4)
        self.sp_diff = ttk.Spinbox(bar, from_=1, to=5, increment=1, width=5, font=FONT)
        self.sp_diff.set(3)
        self.sp_diff.grid(row=0, column=3, padx=4)
        ttk.Label(bar, text="题量", font=FONT).grid(row=0, column=4, padx=4)
        self.sp_count = ttk.Spinbox(bar, from_=5, to=30, increment=1, width=5, font=FONT)
        self.sp_count.set(10)
        self.sp_count.grid(row=0, column=5, padx=4)
        ttk.Button(bar, text="生成模拟卷", command=self._start_mock).grid(row=0, column=6, padx=10)

        self.frm_mock_quiz = ttk.Frame(self.tab_mock)
        self.frm_mock_quiz.pack(fill="both", expand=True)

    def _start_mock(self):
        try:
            diff = float(self.sp_diff.get())
            count = int(self.sp_count.get())
        except ValueError:
            messagebox.showinfo("提示", "难度系数/题量格式不正确")
            return
        qs = engine.generate_paper(self.cb_mock_subject.get(), diff, count)
        if not qs:
            messagebox.showinfo("提示", "题库中该科目暂无题目，请老师先录入。")
            return
        self._mount_quiz(self.frm_mock_quiz, qs,
                         "模拟卷 · %s · 难度%s（共%d题）"
                         % (self.cb_mock_subject.get(), self.sp_diff.get(), len(qs)))

    # ---------- 真题 ----------
    def _build_real_tab(self):
        bar = ttk.LabelFrame(self.tab_real, text="选择真题卷")
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Button(bar, text="刷新列表", command=self._refresh_papers).pack(side="left", padx=6, pady=6)

        cols = ("title", "subject", "year", "ptype", "difficulty", "count")
        self.tv_papers = ttk.Treeview(self.tab_real, columns=cols, show="headings", height=8)
        heads = ["试卷名称", "科目", "年份", "类型", "难度", "题数"]
        widths = [320, 70, 60, 60, 60, 60]
        for c, h, w in zip(cols, heads, widths):
            self.tv_papers.heading(c, text=h)
            self.tv_papers.column(c, width=w, anchor="center")
        self.tv_papers.pack(fill="x", padx=10)
        ttk.Button(self.tab_real, text="开始作答", command=self._start_real).pack(pady=6)

        self.frm_real_quiz = ttk.Frame(self.tab_real)
        self.frm_real_quiz.pack(fill="both", expand=True)
        self._refresh_papers()

    def _refresh_papers(self):
        for i in self.tv_papers.get_children():
            self.tv_papers.delete(i)
        self._papers = store.list_papers()
        for p in self._papers:
            self.tv_papers.insert("", "end", iid=p["pid"],
                                  values=(p["title"], p["subject"], p["year"],
                                          p["ptype"], p["difficulty"],
                                          len(str(p["qids"]).split(","))))

    def _start_real(self):
        sel = self.tv_papers.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一套试卷")
            return
        paper = next(p for p in self._papers if p["pid"] == sel[0])
        qids = str(paper["qids"]).split(",")
        qs = store.get_questions_by_ids(qids)
        if not qs:
            messagebox.showinfo("提示", "该试卷题目不存在")
            return
        self._mount_quiz(self.frm_real_quiz, qs, "真题 · %s" % paper["title"])

    # ---------- 错题本 ----------
    def _build_wrong_tab(self):
        ttk.Button(self.tab_wrong, text="刷新", command=self._refresh_wrong).pack(anchor="w", padx=10, pady=6)
        cols = ("qid", "subject", "kp", "stem", "ts")
        self.tv_wrong = ttk.Treeview(self.tab_wrong, columns=cols, show="headings", height=10)
        heads = ["题号", "科目", "知识点", "题干（摘要）", "答错时间"]
        widths = [70, 60, 180, 380, 140]
        for c, h, w in zip(cols, heads, widths):
            self.tv_wrong.heading(c, text=h)
            self.tv_wrong.column(c, width=w, anchor="w")
        self.tv_wrong.pack(fill="both", expand=True, padx=10)
        ttk.Button(self.tab_wrong, text="重做全部错题", command=self._redo_wrong).pack(pady=6)
        self.frm_wrong_quiz = ttk.Frame(self.tab_wrong)
        self.frm_wrong_quiz.pack(fill="both", expand=True)
        self._refresh_wrong()

    def _refresh_wrong(self):
        for i in self.tv_wrong.get_children():
            self.tv_wrong.delete(i)
        self._wrong_qs = []
        seen = set()
        for r in store.load_records(username=self.username, correct=False):
            if r["qid"] in seen:
                continue
            seen.add(r["qid"])
            q = store.get_question(r["qid"])
            if not q:
                continue
            self._wrong_qs.append(q)
            self.tv_wrong.insert("", "end", iid=q["qid"],
                                 values=(q["qid"], q["subject"], q["knowledge_point"],
                                         str(q["stem"])[:40], r["ts"]))

    def _redo_wrong(self):
        if not self._wrong_qs:
            messagebox.showinfo("提示", "暂无错题，继续保持！")
            return
        self._mount_quiz(self.frm_wrong_quiz, list(self._wrong_qs),
                         "错题重做（共%d题）" % len(self._wrong_qs))

    # ---------- 学情看板 ----------
    def _build_dash_tab(self):
        self.frm_overall = ttk.LabelFrame(self.tab_dash, text="学习总览")
        self.frm_overall.pack(fill="x", padx=10, pady=8)

        mid = ttk.Frame(self.tab_dash)
        mid.pack(fill="both", expand=True, padx=10)
        left = ttk.LabelFrame(mid, text="知识点掌握度画像（按掌握度升序，越靠前越薄弱）")
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        cols = ("subject", "kp", "mastery", "status", "total", "acc")
        self.tv_kp = ttk.Treeview(left, columns=cols, show="headings")
        heads = ["科目", "知识点", "掌握度", "状态", "练题数", "正确率%"]
        widths = [60, 200, 70, 70, 60, 70]
        for c, h, w in zip(cols, heads, widths):
            self.tv_kp.heading(c, text=h)
            self.tv_kp.column(c, width=w, anchor="center")
        self.tv_kp.pack(fill="both", expand=True, pady=4)

        right = ttk.LabelFrame(mid, text="复测提醒（遗忘曲线）")
        right.pack(side="left", fill="both", expand=True)
        self.txt_due = tk.Text(right, height=10, wrap="word", font=FONT, bg="#FAFAFA")
        self.txt_due.pack(fill="both", expand=True, pady=4)
        ttk.Button(right, text="开始复测", command=self._start_retest).pack(pady=4)

        self.frm_due_quiz = ttk.Frame(self.tab_dash)
        self.frm_due_quiz.pack(fill="both", expand=True)

    def _refresh_dash(self):
        if not hasattr(self, "tv_kp"):
            return
        for w in self.frm_overall.winfo_children():
            w.destroy()
        ov = analytics.overall(self.username)
        txt = ("累计作答 %d 题    总正确率 %.1f%%    已练知识点 %d 个    "
               "已过关 %d 个    待复测 %d 个"
               % (ov["total_answers"], ov["accuracy"], ov["kp_touched"],
                  ov["kp_passed"], ov["retest_due"]))
        ttk.Label(self.frm_overall, text=txt, font=FONT_TITLE,
                  foreground="#1F3B66").pack(padx=10, pady=8)

        for i in self.tv_kp.get_children():
            self.tv_kp.delete(i)
        for row in analytics.kp_summary(self.username):
            self.tv_kp.insert("", "end", values=(
                row["subject"], row["kp"], "%.0f" % row["mastery"],
                row["status"], row["total"], row["accuracy"]))

        self.txt_due.config(state="normal")
        self.txt_due.delete("1.0", "end")
        due = engine.retest_due(self.username)
        if due:
            self.txt_due.insert("1.0", "以下已过关知识点到了复测时间：\n\n")
            for r in due:
                self.txt_due.insert("end", "· %s（掌握度 %.0f，已过 %s 天）\n"
                                           % (r["knowledge_point"],
                                              float(r["mastery"] or 0),
                                              r["last_review"]))
        else:
            self.txt_due.insert("1.0", "暂无到期复测任务。\n已过关知识点会按遗忘曲线（1/2/4/7/15/30天）提醒你复测。")
        self.txt_due.config(state="disabled")

    def _start_retest(self):
        pairs = study_loop.retest_questions(self.username)
        if not pairs:
            messagebox.showinfo("提示", "暂无到期复测任务")
            return
        qs = [q for _, q in pairs]
        self._mount_quiz(self.frm_due_quiz, qs, "复测（共%d个知识点）" % len(qs))
        # 复测判分走 on_answer 已会更新掌握度；此处简化处理，过关复测次数在下次过关时累计

    # ---------- 公共 ----------
    def _mount_quiz(self, frame, questions, title):
        for w in frame.winfo_children():
            w.destroy()
        QuizFrame(frame, self.username, questions, title,
                  on_close=lambda: None).pack(fill="both", expand=True)
