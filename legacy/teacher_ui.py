# -*- coding: utf-8 -*-
"""
teacher_ui.py — 教师端界面
试卷识别录入（上传扫描件/Word/PDF → 大模型识别 → 校对 → 入库）
题库管理 · 知识点维护 · 账号管理 · API设置
"""
import os
import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import excel_store as store
import ai_recognize
import analytics

FONT = ("Microsoft YaHei UI", 10)
FONT_B = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 12, "bold")


# ================= 题目编辑对话框（录入校对/题库管理共用） =================
class QuestionDialog(tk.Toplevel):
    def __init__(self, parent, data, on_save, title="编辑题目"):
        super().__init__(parent)
        self.title(title)
        self.geometry("640x620")
        self.transient(parent)
        self.on_save = on_save
        self.data = dict(data)
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=8)

        row = 0
        ttk.Label(body, text="科目", font=FONT).grid(row=row, column=0, sticky="w", pady=3)
        self.cb_subject = ttk.Combobox(body, values=store.SUBJECTS, font=FONT, width=10)
        self.cb_subject.set(self.data.get("subject", "数学"))
        self.cb_subject.grid(row=row, column=1, sticky="w")
        ttk.Label(body, text="章节", font=FONT).grid(row=row, column=2, sticky="w")
        self.ent_chapter = ttk.Entry(body, font=FONT, width=14)
        self.ent_chapter.insert(0, self.data.get("chapter", ""))
        self.ent_chapter.grid(row=row, column=3, sticky="w")
        ttk.Label(body, text="题型", font=FONT).grid(row=row, column=4, sticky="w")
        self.cb_qtype = ttk.Combobox(body, values=store.QTYPES, state="readonly", font=FONT, width=8)
        self.cb_qtype.set(self.data.get("qtype", "单选"))
        self.cb_qtype.grid(row=row, column=5, sticky="w")
        ttk.Label(body, text="难度(1-5)", font=FONT).grid(row=row, column=6, sticky="w")
        self.sp_diff = ttk.Spinbox(body, from_=1, to=5, width=4, font=FONT)
        self.sp_diff.set(int(self.data.get("difficulty", 3) or 3))
        self.sp_diff.grid(row=row, column=7, sticky="w")

        row += 1
        ttk.Label(body, text="知识点(多个用英文逗号分隔)", font=FONT).grid(row=row, column=0, sticky="w", pady=3)
        self.ent_kp = ttk.Entry(body, font=FONT, width=60)
        self.ent_kp.insert(0, self.data.get("knowledge_point", ""))
        self.ent_kp.grid(row=row, column=1, columnspan=7, sticky="we")

        row += 1
        ttk.Label(body, text="题干", font=FONT).grid(row=row, column=0, sticky="nw", pady=3)
        self.txt_stem = tk.Text(body, height=6, wrap="word", font=FONT)
        self.txt_stem.grid(row=row, column=1, columnspan=7, sticky="we")
        self.txt_stem.insert("1.0", self.data.get("stem", ""))

        row += 1
        ttk.Label(body, text="选项(每行一个,\n不含ABCD标号)", font=FONT).grid(row=row, column=0, sticky="nw", pady=3)
        self.txt_opts = tk.Text(body, height=5, wrap="word", font=FONT)
        self.txt_opts.grid(row=row, column=1, columnspan=7, sticky="we")
        opts = self.data.get("options", [])
        if isinstance(opts, str):
            try:
                opts = json.loads(opts)
            except Exception:
                opts = [opts] if opts else []
        self.txt_opts.insert("1.0", "\n".join(opts))

        row += 1
        ttk.Label(body, text="答案", font=FONT).grid(row=row, column=0, sticky="w", pady=3)
        self.ent_answer = ttk.Entry(body, font=FONT, width=60)
        self.ent_answer.insert(0, str(self.data.get("answer", "")))
        self.ent_answer.grid(row=row, column=1, columnspan=7, sticky="we")

        row += 1
        ttk.Label(body, text="详细解析", font=FONT).grid(row=row, column=0, sticky="nw", pady=3)
        self.txt_analysis = tk.Text(body, height=5, wrap="word", font=FONT)
        self.txt_analysis.grid(row=row, column=1, columnspan=7, sticky="we")
        self.txt_analysis.insert("1.0", self.data.get("analysis", ""))

        # 图形方案标记 + 图片缩略图 + HTML预览
        row += 1
        media_paths = store.media_paths_of(self.data)
        media_type = self.data.get("media_type", "无")
        if media_paths or (media_type and media_type != "无"):
            mf = ttk.LabelFrame(body, text="图形图表（方案: %s）" % (media_type or "未标记"))
            mf.grid(row=row, column=0, columnspan=8, sticky="we", pady=4)
            self._thumbs = []
            from PIL import Image, ImageTk
            for fname in media_paths[:4]:
                p = fname if os.path.isabs(fname) else os.path.join(store.media_dir(), fname)
                if not os.path.exists(p):
                    continue
                try:
                    im = Image.open(p)
                    im.thumbnail((220, 160))
                    photo = ImageTk.PhotoImage(im)
                    self._thumbs.append(photo)
                    tk.Label(mf, image=photo, bd=1, relief="solid").pack(side="left", padx=4, pady=4)
                except Exception:
                    pass
            ttk.Button(mf, text="HTML 预览（KaTeX 渲染公式）",
                       command=self._preview_html).pack(side="right", padx=6)
            row += 1

        row += 1
        ttk.Label(body, text="标签", font=FONT).grid(row=row, column=0, sticky="w", pady=3)
        self.cb_tag = ttk.Combobox(body,
                                    values=("训练题", "高考真题", "高考模拟", "校考", "联考", "竞赛",
                                            "期中", "期末", "一模", "二模", "其他"),
                                    font=FONT, width=12)
        self.cb_tag.insert(0, str(self.data.get("tag", "训练题") or "训练题"))
        self.cb_tag.grid(row=row, column=1, sticky="w")
        ttk.Label(body, text="(可手动改：年份/地区/类型，例如 2023/北京/高考真题)",
                  font=("Microsoft YaHei UI", 9), foreground="#718096").grid(row=row, column=2, columnspan=6, sticky="w", padx=6)

        row += 1
        ttk.Button(body, text="保存", command=self._save).grid(row=row, column=1, sticky="w", pady=10)
        ttk.Button(body, text="取消", command=self.destroy).grid(row=row, column=2, sticky="w", pady=10)

        for c in range(8):
            body.columnconfigure(c, weight=1)

    def _preview_html(self):
        d = dict(self.data)
        d["stem"] = self.txt_stem.get("1.0", "end").strip()
        d["analysis"] = self.txt_analysis.get("1.0", "end").strip()
        d["options"] = [ln.strip() for ln in self.txt_opts.get("1.0", "end").splitlines() if ln.strip()]
        d["answer"] = self.ent_answer.get().strip()
        import html_preview
        html_preview.preview_questions([d], title="题目预览")

    def _save(self):
        d = {
            "subject": self.cb_subject.get().strip() or "数学",
            "chapter": self.ent_chapter.get().strip() or "未分章",
            "knowledge_point": self.ent_kp.get().strip() or "待标注",
            "qtype": self.cb_qtype.get(),
            "difficulty": int(self.sp_diff.get() or 3),
            "stem": self.txt_stem.get("1.0", "end").strip(),
            "options": [ln.strip() for ln in self.txt_opts.get("1.0", "end").splitlines() if ln.strip()],
            "answer": self.ent_answer.get().strip(),
            "analysis": self.txt_analysis.get("1.0", "end").strip(),
            "tag": self.cb_tag.get().strip() or "训练题",
        }
        if not d["stem"]:
            messagebox.showinfo("提示", "题干不能为空", parent=self)
            return
        self.data.update(d)
        self.on_save(self.data)
        self.destroy()


# ================= 教师端主窗口 =================
class TeacherWindow(tk.Toplevel):
    def __init__(self, master, user):
        super().__init__()
        self.master = master
        self.user = user
        self.title("高考学习训练系统（简化版）— 教师端 · %s" % user["name"])
        self.geometry("1080x740")
        self.protocol("WM_DELETE_WINDOW", self._quit)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)

        self.tab_import = ttk.Frame(nb); nb.add(self.tab_import, text="试卷识别录入")
        self.tab_lib = ttk.Frame(nb); nb.add(self.tab_lib, text="题库管理")
        self.tab_dash = ttk.Frame(nb); nb.add(self.tab_dash, text="看板")
        self.tab_kp = ttk.Frame(nb); nb.add(self.tab_kp, text="知识点维护")
        self.tab_user = ttk.Frame(nb); nb.add(self.tab_user, text="账号管理")
        self.tab_api = ttk.Frame(nb); nb.add(self.tab_api, text="API设置")

        self._build_import_tab()
        self._build_lib_tab()
        self._build_dash_tab()
        self._build_kp_tab()
        self._build_user_tab()
        self._build_api_tab()

    def _quit(self):
        self.master.destroy()

    # ---------------- ① 试卷识别录入 ----------------
    def _build_import_tab(self):
        top = ttk.LabelFrame(self.tab_import, text="第1步：上传试卷文件（扫描件 / Word / PDF），第2步：识别，第3步：校对，第4步：入库")
        top.pack(fill="x", padx=10, pady=8)

        self.var_file = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_file, font=FONT, width=52).grid(row=0, column=0, padx=6, pady=8)
        ttk.Button(top, text="浏览…", command=self._pick_file).grid(row=0, column=1, padx=4)
        ttk.Label(top, text="科目提示(可选)", font=FONT).grid(row=0, column=2, padx=4)
        self.cb_hint = ttk.Combobox(top, values=[""] + store.SUBJECTS, state="readonly", font=FONT, width=8)
        self.cb_hint.grid(row=0, column=3, padx=4)
        self.btn_recog = ttk.Button(top, text="开始识别", command=self._start_recognize)
        self.btn_recog.grid(row=0, column=4, padx=8)

        self.lbl_status = ttk.Label(self.tab_import, text="就绪。未配置API时为演示模式（按题号切分文本）。",
                                    font=FONT, foreground="#2B6CB0")
        self.lbl_status.pack(anchor="w", padx=10)

        cols = ("subject", "chapter", "kp", "qtype", "diff", "stem", "answer", "conf", "tag")
        self.tv_imp = ttk.Treeview(self.tab_import, columns=cols, show="headings", height=16)
        heads = ["科目", "章节", "知识点", "题型", "难度", "题干（摘要）", "答案", "置信度", "标签"]
        widths = [60, 90, 150, 50, 50, 320, 90, 60, 160]
        for c, h, w in zip(cols, heads, widths):
            self.tv_imp.heading(c, text=h)
            self.tv_imp.column(c, width=w, anchor="w" if c == "stem" else "center")
        self.tv_imp.pack(fill="both", expand=True, padx=10, pady=4)
        self.tv_imp.bind("<Double-1>", self._edit_import_row)

        bar = ttk.Frame(self.tab_import)
        bar.pack(fill="x", padx=10, pady=6)
        ttk.Button(bar, text="编辑选中(双击行亦可)", command=self._edit_import_row).pack(side="left", padx=4)
        ttk.Button(bar, text="删除选中", command=self._del_import_row).pack(side="left", padx=4)
        self.btn_commit = ttk.Button(bar, text="全部校对完成，写入题库", command=self._commit_import)
        self.btn_commit.pack(side="left", padx=12)

        self._import_items = []

    def _pick_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="选择试卷文件",
            filetypes=[("试卷文件", "*.docx *.pdf *.png *.jpg *.jpeg *.bmp *.txt"),
                       ("Word", "*.docx"), ("PDF", "*.pdf"),
                       ("图片/扫描件", "*.png *.jpg *.jpeg *.bmp"), ("文本", "*.txt")])
        if path:
            self.var_file.set(path)

    def _start_recognize(self):
        path = self.var_file.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "请先选择存在的试卷文件")
            return
        self.btn_recog.config(state="disabled")
        self.btn_commit.config(state="disabled")
        self._import_items = []
        self._reload_import_tv()
        self.lbl_status.config(text="正在启动识别…")

        # 累计统计
        stats = {"new": 0, "skip": 0, "chunks_done": 0}

        def save_chunk(items_chunk, idx, total):
            """每识别完一段立即回调：去重入库 + 把题目追加到表格 + 实时刷新进度。"""
            ids, sk = store.add_questions(items_chunk)
            stats["new"] += len(ids)
            stats["skip"] += sk
            stats["chunks_done"] = idx
            # UI 必须在主线程：通过 after 切回主线程
            def update_ui():
                self._import_items.extend(items_chunk)
                self._reload_import_tv()
                self.lbl_status.config(
                    text=("正在识别中…第 %d/%d 段完成  ｜  本次已自动入库 %d 道题  "
                          "（跳重 %d 道）" % (idx, total, len(ids), sk)))
            self.after(0, update_ui)

        def work():
            err = None
            try:
                items, msg = ai_recognize.recognize(
                    path, self.cb_hint.get().strip(), stream_cb=save_chunk)
            except Exception as e:
                items, msg, err = [], "", e
            self.after(0, lambda: self._recognize_done(items, msg, err, stats))

        threading.Thread(target=work, daemon=True).start()

    def _recognize_done(self, items, msg, err, stats=None):
        self.btn_recog.config(state="normal")
        self.btn_commit.config(state="normal")
        if err:
            self.lbl_status.config(text="识别失败：" + str(err))
            messagebox.showerror("识别失败", str(err))
            return
        # 已经在 stream_cb 中追加过 items，这里只补一遍去重后的最终列表（_recognize 内部按 stem 去重了）
        # 如果 items 是空（已全部回调完毕）则不重复追加
        if items and not self._import_items:
            self._import_items = list(items)
            self._reload_import_tv()
        # 最终状态汇总
        s = stats or {"new": 0, "skip": 0, "chunks_done": 0}
        tail = "。识别过程中已逐段自动入库（防丢失），可继续校对表格中的题目或点“写入题库”提交修改。" \
               if s.get("new") else "。请逐题校对（双击编辑），确认后点击“写入题库”。"
        self.lbl_status.config(
            text=(msg + tail + "  ｜  累计：自动入库 %d 道，跳重 %d 道" % (s["new"], s["skip"])))

    def _reload_import_tv(self):
        for i in self.tv_imp.get_children():
            self.tv_imp.delete(i)
        for idx, it in enumerate(self._import_items):
            self.tv_imp.insert("", "end", iid=str(idx), values=(
                it["subject"], it["chapter"], it["knowledge_point"], it["qtype"],
                it["difficulty"], str(it["stem"])[:45].replace("\n", " "),
                str(it["answer"])[:20], "%.2f" % it.get("confidence", 0.8),
                it.get("tag", "")))

    def _edit_import_row(self, event=None):
        sel = self.tv_imp.selection()
        if not sel:
            return
        idx = int(sel[0])
        QuestionDialog(self, self._import_items[idx],
                       on_save=lambda d: (self._import_items.__setitem__(idx, d),
                                          self._reload_import_tv()),
                       title="校对题目")

    def _del_import_row(self):
        sel = self.tv_imp.selection()
        if not sel:
            return
        for iid in sorted(sel, key=int, reverse=True):
            self._import_items.pop(int(iid))
        self._reload_import_tv()

    def _commit_import(self):
        if not self._import_items:
            messagebox.showinfo("提示", "没有可入库的题目")
            return
        missing = [i + 1 for i, it in enumerate(self._import_items)
                   if not str(it.get("answer", "")).strip()]
        if missing:
            if not messagebox.askyesno("提示", "第 %s 题缺少答案，确定仍要入库吗？" % ",".join(map(str, missing[:5]))):
                return
        ids, skipped = store.add_questions(self._import_items)
        msg = "已成功入库 %d 道题" % len(ids)
        if ids:
            msg += "（题号 %s … %s）" % (ids[0], ids[-1])
        if skipped:
            msg += "\n\n跳过 %d 道重复题（题干与题库已有题目相同），未保存。" % skipped
        messagebox.showinfo("完成", msg)
        self._import_items = []
        self._reload_import_tv()
        self._refresh_lib()
        self._refresh_kp_tree()

    # ---------------- ② 题库管理 ----------------
    def _build_lib_tab(self):
        bar = ttk.LabelFrame(self.tab_lib, text="筛选")
        bar.pack(fill="x", padx=10, pady=8)
        self.cb_lib_subject = ttk.Combobox(bar, values=[""] + store.SUBJECTS, state="readonly", font=FONT, width=8)
        self.cb_lib_type = ttk.Combobox(bar, values=[""] + store.QTYPES, state="readonly", font=FONT, width=8)
        self.ent_lib_kw = ttk.Entry(bar, font=FONT, width=18)
        ttk.Label(bar, text="科目", font=FONT).grid(row=0, column=0, padx=4, pady=6)
        self.cb_lib_subject.grid(row=0, column=1)
        ttk.Label(bar, text="题型", font=FONT).grid(row=0, column=2, padx=4)
        self.cb_lib_type.grid(row=0, column=3)
        ttk.Label(bar, text="题干关键词", font=FONT).grid(row=0, column=4, padx=4)
        self.ent_lib_kw.grid(row=0, column=5)
        ttk.Button(bar, text="查询", command=self._refresh_lib).grid(row=0, column=6, padx=8)

        cols = ("qid", "subject", "chapter", "kp", "qtype", "diff", "stem", "answer", "tag", "media")
        self.tv_lib = ttk.Treeview(self.tab_lib, columns=cols, show="headings", height=17)
        heads = ["题号", "科目", "章节", "知识点", "题型", "难度", "题干（摘要）", "答案", "标签", "图形"]
        widths = [60, 55, 85, 150, 50, 45, 280, 80, 140, 90]
        for c, h, w in zip(cols, heads, widths):
            self.tv_lib.heading(c, text=h)
            self.tv_lib.column(c, width=w, anchor="w" if c == "stem" else "center")
        self.tv_lib.pack(fill="both", expand=True, padx=10, pady=4)
        self.tv_lib.bind("<Double-1>", self._edit_lib_row)

        bar2 = ttk.Frame(self.tab_lib)
        bar2.pack(fill="x", padx=10, pady=6)
        ttk.Button(bar2, text="编辑选中", command=self._edit_lib_row).pack(side="left", padx=4)
        ttk.Button(bar2, text="删除选中", command=self._del_lib_row).pack(side="left", padx=4)
        ttk.Button(bar2, text="🖼 HTML预览选中题", command=self._preview_lib_row).pack(side="left", padx=4)
        ttk.Label(bar2, text="  标签过滤", font=FONT).pack(side="left", padx=(12, 4))
        self.cb_lib_tag = ttk.Combobox(bar2, values=("全部", "训练题", "高考真题", "高考模拟", "校考", "联考", "竞赛"),
                                        state="readonly", font=FONT, width=10)
        self.cb_lib_tag.set("全部")
        self.cb_lib_tag.pack(side="left", padx=4)
        self.cb_lib_tag.bind("<<ComboboxSelected>>", lambda e: self._refresh_lib())
        self.lbl_lib_count = ttk.Label(bar2, text="", font=FONT)
        self.lbl_lib_count.pack(side="right", padx=8)
        self._refresh_lib()

    def _refresh_lib(self):
        tag_filter = ""
        if hasattr(self, "cb_lib_tag"):
            v = self.cb_lib_tag.get()
            if v == "训练题":
                tag_filter = "训练题"
            elif v != "全部":
                tag_filter = v
        self._lib_rows = store.load_questions(
            subject=self.cb_lib_subject.get() or None,
            qtype=self.cb_lib_type.get() or None,
            tag=tag_filter or None,
            keyword=self.ent_lib_kw.get().strip() or None)
        for i in self.tv_lib.get_children():
            self.tv_lib.delete(i)
        for q in self._lib_rows[:500]:
            self.tv_lib.insert("", "end", iid=q["qid"], values=(
                q["qid"], q["subject"], q["chapter"], q["knowledge_point"],
                q["qtype"], q["difficulty"], str(q["stem"])[:45].replace("\n", " "),
                str(q["answer"])[:18], q.get("tag", ""),
                q.get("media_type", "") if q.get("media_type", "无") != "无" else ""))
        self.lbl_lib_count.config(text="共 %d 题%s" % (len(self._lib_rows),
                                                     "（仅显示前500）" if len(self._lib_rows) > 500 else ""))

    def _preview_lib_row(self):
        """用 KaTeX/mhchem 渲染选中题（含图片与公式）并在浏览器打开。"""
        sel = self.tv_lib.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中一道题")
            return
        q = store.get_question(sel[0])
        if not q:
            return
        import html_preview
        html_preview.preview_questions([q], title="题目预览 %s" % q.get("qid", ""))

    def _edit_lib_row(self, event=None):
        sel = self.tv_lib.selection()
        if not sel:
            return
        q = store.get_question(sel[0])
        if not q:            return
        q = dict(q)
        q["options"] = store.parse_options(q)

        def on_save(d):
            store.update_question(q["qid"], d)
            self._refresh_lib()
        QuestionDialog(self, q, on_save=on_save, title="编辑题目 %s" % q["qid"])

    def _del_lib_row(self):
        sel = self.tv_lib.selection()
        if not sel:
            return
        if messagebox.askyesno("确认", "确定删除选中的 %d 道题？" % len(sel)):
            for qid in sel:
                store.delete_question(qid)
            self._refresh_lib()

    # ---------------- ② 看板：题库统计（各科 × 题型分类） ----------------
    def _build_dash_tab(self):
        # 顶部：总览 + 刷新按钮
        top = ttk.Frame(self.tab_dash)
        top.pack(fill="x", padx=12, pady=10)
        self.lbl_dash_total = ttk.Label(top, text="", font=FONT_TITLE, foreground="#1F3B66")
        self.lbl_dash_total.pack(side="left")
        ttk.Button(top, text="🔄 刷新", command=self._refresh_dash_tab).pack(side="right")

        # 整体分类汇总
        self.frm_cat_summary = ttk.LabelFrame(self.tab_dash, text="题库题型分类汇总（不分科目）")
        self.frm_cat_summary.pack(fill="x", padx=12, pady=4)
        self.frm_cat_summary_inner = ttk.Frame(self.frm_cat_summary)
        self.frm_cat_summary_inner.pack(fill="x", padx=10, pady=8)

        # 看板卡片容器（可滚动）
        wrap = ttk.LabelFrame(self.tab_dash, text="各科目分布")
        wrap.pack(fill="both", expand=True, padx=12, pady=8)
        canvas = tk.Canvas(wrap, highlightthickness=0, bg="#F7FAFC")
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        self.frm_dash_cards = ttk.Frame(canvas)
        self.frm_dash_cards.bind("<Configure>",
                                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=self.frm_dash_cards, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        # 鼠标滚轮支持
        def _on_wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        self._refresh_dash_tab()

    def _refresh_dash_tab(self):
        for w in self.frm_cat_summary_inner.winfo_children():
            w.destroy()
        for w in self.frm_dash_cards.winfo_children():
            w.destroy()
        data = analytics.category_stats()
        self.lbl_dash_total.config(text="题库总题数：%d 道" % data["total"])
        # 分类汇总：4 个等宽 stat 卡片
        for i, (cat, _) in enumerate(analytics.CATEGORIES):
            n = data["by_category"].get(cat, 0)
            ratio = (n / data["total"] * 100) if data["total"] else 0
            card = tk.Frame(self.frm_cat_summary_inner, bg=analytics.CAT_COLORS.get(cat, "#4A5568"),
                             padx=14, pady=10, cursor="hand2")
            card.grid(row=0, column=i, sticky="ew", padx=6, pady=4)
            self.frm_cat_summary_inner.columnconfigure(i, weight=1)
            tk.Label(card, text=cat, bg=card["bg"], fg="white",
                     font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
            tk.Label(card, text="%d" % n, bg=card["bg"], fg="white",
                     font=("Microsoft YaHei UI", 20, "bold")).pack(anchor="w")
            tk.Label(card, text="占题库 %.1f%%" % ratio, bg=card["bg"], fg="#E2E8F0",
                     font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        # 各科卡片
        SUBJECTS_ORDER = ["语文", "英语", "数学", "物理", "化学", "生物"]
        MAX = max((s["total"] for s in data["by_subject"]), default=1) or 1
        # 3 列网格
        cols = 3
        for idx, s in enumerate(data["by_subject"]):
            r, c = divmod(idx, cols)
            self._make_subject_card(self.frm_dash_cards, s, max_total=MAX, row=r, col=c)

    def _make_subject_card(self, parent, s, max_total, row, col):
        color = analytics.SUBJECT_COLORS.get(s["subject"], "#4A5568")
        total = s["total"]
        card = tk.Frame(parent, bg="white", highlightbackground="#E2E8F0",
                         highlightthickness=1, padx=0, pady=0, height=300)
        card.grid(row=row, column=col, sticky="nsew", padx=6, pady=6)
        card.pack_propagate(False)    # 固定高度，让同行的卡片等高对齐
        parent.columnconfigure(col, weight=1)
        parent.rowconfigure(row, weight=1)
        # 顶部色条
        bar = tk.Frame(card, bg=color, height=6)
        bar.pack(fill="x")
        body = tk.Frame(card, bg="white", padx=12, pady=10)
        body.pack(fill="both", expand=True)
        # 标题：科目名 + 总数
        head = tk.Frame(body, bg="white")
        head.pack(fill="x")
        dot = tk.Label(head, text="●", bg="white", fg=color,
                       font=("Microsoft YaHei UI", 14))
        dot.pack(side="left")
        tk.Label(head, text="  %s" % s["subject"], bg="white", fg="#1A202C",
                 font=("Microsoft YaHei UI", 13, "bold")).pack(side="left")
        tk.Label(head, text="  共 %d 道" % total, bg="white", fg="#4A5568",
                 font=("Microsoft YaHei UI", 10)).pack(side="right")
        # 整体进度（该科占题库总题数的占比）
        ratio_total = (total / max_total * 100) if max_total else 0
        self._add_progress_bar(body, "题量占比", ratio_total, "#CBD5E0", text_after="%.0f%%" % ratio_total)
        # 4 个分类条
        for cat, _ in analytics.CATEGORIES:
            n = s["categories"].get(cat, 0)
            ratio = (n / total * 100) if total else 0
            c_color = analytics.CAT_COLORS.get(cat, "#4A5568")
            self._add_progress_bar(body, "%-4s %d 道" % (cat, n), ratio, c_color,
                                    text_after="%.0f%%" % ratio)
        # 标签
        tags = sorted([t for t in s["tags"] if t and t != "训练题"])
        if tags:
            tag_row = tk.Frame(body, bg="white")
            tag_row.pack(fill="x", pady=(8, 0))
            tk.Label(tag_row, text="标签：", bg="white", fg="#4A5568",
                     font=("Microsoft YaHei UI", 9)).pack(side="left")
            for t in tags[:4]:    # 最多展示 4 个标签
                tk.Label(tag_row, text="  " + t, bg="#EDF2F7", fg="#2D3748",
                         font=("Microsoft YaHei UI", 9), padx=6, pady=1,
                         relief="flat").pack(side="left", padx=2)
            if len(tags) > 4:
                tk.Label(tag_row, text=f"  +{len(tags) - 4}", bg="#F7FAFC",
                         fg="#718096", font=("Microsoft YaHei UI", 9),
                         padx=4).pack(side="left", padx=2)
        elif total > 0:
            tk.Label(body, text="标签：训练题（无地区/年份信息）", bg="white", fg="#A0AEC0",
                     font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(8, 0))

    def _add_progress_bar(self, parent, label, percent, color, text_after=""):
        row = tk.Frame(parent, bg="white")
        row.pack(fill="x", pady=2)
        tk.Label(row, text=label, bg="white", fg="#2D3748",
                 font=("Microsoft YaHei UI", 9), width=14, anchor="w").pack(side="left")
        # 进度条：用 Canvas 绘制（不依赖 ttk 主题色，跨平台好看）
        bar_w, bar_h = 130, 8
        cnv = tk.Canvas(row, width=bar_w, height=bar_h, bg="#EDF2F7",
                        highlightthickness=0, bd=0)
        cnv.pack(side="left", padx=2)
        if percent > 0:
            cnv.create_rectangle(0, 0, max(1, int(bar_w * min(percent, 100) / 100)), bar_h,
                                 fill=color, outline="")
        tk.Label(row, text=text_after, bg="white", fg="#4A5568",
                 font=("Microsoft YaHei UI", 9), width=6, anchor="w").pack(side="left", padx=4)

    # ---------------- ③ 知识点维护 ----------------
    def _build_kp_tab(self):
        bar = ttk.LabelFrame(self.tab_kp, text="新增知识点")
        bar.pack(fill="x", padx=10, pady=8)
        self.cb_kp_subject = ttk.Combobox(bar, values=store.SUBJECTS, state="readonly", font=FONT, width=8)
        self.ent_kp_chapter = ttk.Entry(bar, font=FONT, width=16)
        self.ent_kp_name = ttk.Entry(bar, font=FONT, width=26)
        ttk.Label(bar, text="科目", font=FONT).grid(row=0, column=0, padx=4, pady=6)
        self.cb_kp_subject.grid(row=0, column=1)
        ttk.Label(bar, text="章节", font=FONT).grid(row=0, column=2, padx=4)
        self.ent_kp_chapter.grid(row=0, column=3)
        ttk.Label(bar, text="知识点", font=FONT).grid(row=0, column=4, padx=4)
        self.ent_kp_name.grid(row=0, column=5)
        ttk.Button(bar, text="添加", command=self._add_kp).grid(row=0, column=6, padx=8)

        self.tv_kp_tree = ttk.Treeview(self.tab_kp, show="tree", height=18)
        self.tv_kp_tree.pack(fill="both", expand=True, padx=10, pady=4)
        self._refresh_kp_tree()

    def _refresh_kp_tree(self):
        tree = store.get_kp_tree()
        self.tv_kp_tree.delete(*self.tv_kp_tree.get_children())
        for s, chapters in tree.items():
            sid = self.tv_kp_tree.insert("", "end", text="📘 " + s, open=True)
            for c, kps in chapters.items():
                cid = self.tv_kp_tree.insert(sid, "end", text="📖 " + c, open=False)
                for k in kps:
                    self.tv_kp_tree.insert(cid, "end", text="· " + k)

    def _add_kp(self):
        s = self.cb_kp_subject.get().strip()
        c = self.ent_kp_chapter.get().strip()
        k = self.ent_kp_name.get().strip()
        if not (s and c and k):
            messagebox.showinfo("提示", "科目/章节/知识点都要填写")
            return
        store.add_knowledge_point(s, c, k)
        self.ent_kp_name.delete(0, "end")
        self._refresh_kp_tree()

    # ---------------- ④ 账号管理 ----------------
    def _build_user_tab(self):
        cols = ("username", "name", "role", "status")
        self.tv_users = ttk.Treeview(self.tab_user, columns=cols, show="headings", height=12)
        for c, h, w in zip(cols, ["用户名", "姓名", "角色", "状态"], [140, 140, 100, 100]):
            self.tv_users.heading(c, text=h)
            self.tv_users.column(c, width=w, anchor="center")
        self.tv_users.pack(fill="both", expand=True, padx=10, pady=8)

        bar = ttk.Frame(self.tab_user)
        bar.pack(fill="x", padx=10, pady=6)
        ttk.Button(bar, text="新增账号", command=lambda: self._user_dialog(None)).pack(side="left", padx=4)
        ttk.Button(bar, text="编辑/重置密码", command=self._edit_user).pack(side="left", padx=4)
        ttk.Button(bar, text="启用/停用", command=self._toggle_user).pack(side="left", padx=4)
        ttk.Button(bar, text="刷新", command=self._refresh_users).pack(side="left", padx=12)
        self._refresh_users()

    def _refresh_users(self):
        for i in self.tv_users.get_children():
            self.tv_users.delete(i)
        for u in store.list_users():
            self.tv_users.insert("", "end", iid=u["username"],
                                 values=(u["username"], u["name"], u["role"], u["status"]))

    def _user_dialog(self, u):
        dlg = tk.Toplevel(self)
        dlg.title("账号信息")
        dlg.geometry("360x260")
        dlg.transient(self)
        ents = {}
        labels = ["用户名", "姓名", "密码(留空则不变/默认123456)"]
        keys = ["username", "name", "password"]
        for i, (lab, key) in enumerate(zip(labels, keys)):
            ttk.Label(dlg, text=lab, font=FONT).grid(row=i, column=0, padx=8, pady=8, sticky="w")
            e = ttk.Entry(dlg, font=FONT, width=22, show="*" if key == "password" else "")
            if u and key != "password":
                e.insert(0, u[key])
            if u and key == "username":
                e.config(state="disabled")
            e.grid(row=i, column=1)
            ents[key] = e
        ttk.Label(dlg, text="角色", font=FONT).grid(row=3, column=0, padx=8, sticky="w")
        cb_role = ttk.Combobox(dlg, values=["student", "teacher"], state="readonly", font=FONT, width=18)
        cb_role.set(u["role"] if u else "student")
        cb_role.grid(row=3, column=1)

        def save():
            un = ents["username"].get().strip()
            if not un:
                messagebox.showinfo("提示", "用户名不能为空", parent=dlg)
                return
            store.save_user(un, ents["password"].get(), cb_role.get(),
                            ents["name"].get().strip() or un)
            dlg.destroy()
            self._refresh_users()
        ttk.Button(dlg, text="保存", command=save).grid(row=4, column=1, sticky="w", pady=10)

    def _edit_user(self):
        sel = self.tv_users.selection()
        if not sel:
            return
        u = next(x for x in store.list_users() if x["username"] == sel[0])
        self._user_dialog(u)

    def _toggle_user(self):
        sel = self.tv_users.selection()
        if not sel:
            return
        for un in sel:
            u = next(x for x in store.list_users() if x["username"] == un)
            store.save_user(un, "", u["role"], u["name"],
                            "停用" if u["status"] == "启用" else "启用")
        self._refresh_users()

    # ---------------- ⑤ API设置 ----------------
    def _build_api_tab(self):
        cfg = ai_recognize.load_settings()
        self._api_text_vars = self._build_api_group(
            self.tab_api, "① 文本 API（用于 Word / 文本型 PDF / 文本文件 题目识别）",
            cfg["text"], kind="text", state_label_attr="_lbl_text_state")
        self._api_vision_vars = self._build_api_group(
            self.tab_api, "② 视觉 API（用于扫描件 / 图片型 PDF / 阅卷 多模态识别）",
            cfg["vision"], kind="vision", state_label_attr="_lbl_vision_state")
        # 共用保存按钮
        ttk.Button(self.tab_api, text="保存两套 API 设置", command=self._save_api).pack(
            anchor="w", padx=14, pady=10)

        tip = tk.Text(self.tab_api, height=10, wrap="word", font=FONT, bg="#FAFAFA")
        tip.pack(fill="x", padx=10, pady=6)
        tip.insert("1.0",
                   "说明：\n"
                   "1. 两套 API 完全独立，可分别使用不同厂商/不同 Key（OpenAI 兼容协议，如 DeepSeek、Qwen、OpenAI、智谱等）。\n"
                   "2. 文本 API 用于 Word / 文本型 PDF / 文本文件；视觉 API 用于扫描件、图片型 PDF，以及未来阅卷。\n"
                   "3. 视觉 API 必须支持图片输入（OpenAI 聊天补全格式）。推荐：qwen-vl-max、gpt-4o、deepseek-vl、智谱 glm-4v-plus 等。\n"
                   "4. 关于智谱 GLM-OCR：它使用专属 /api/paas/v4/layout_parsing 端点 + 不同的请求格式（不是聊天补全），本系统暂不直接支持。\n"
                   "   想用智谱模型做 OCR，请改用 GLM-4V（base_url: https://open.bigmodel.cn/api/paas/v4，path: /chat/completions，model: glm-4v-plus），本系统直接兼容。\n"
                   "5. 接口路径默认 /v1/chat/completions；如你的 API 用别的端点路径，修改此字段（必须以 / 开头）。\n"
                   "6. 不配置时，文本识别走演示模式（仅按题号切分），扫描件需配置视觉 API。\n"
                   "7. API Key 仅保存在本机 data/settings.json 中（Base64 简单编码），不会上传他处。")
        tip.config(state="disabled")

    def _build_api_group(self, parent, title, cfg, kind="text", state_label_attr=None):
        frm = ttk.LabelFrame(parent, text=title)
        frm.pack(fill="x", padx=10, pady=8)
        v_base = tk.StringVar(value=cfg.get("api_base", ""))
        v_path = tk.StringVar(value=cfg.get("path", "/v1/chat/completions"))
        v_key = tk.StringVar(value=cfg.get("api_key", ""))
        v_model = tk.StringVar(value=cfg.get("model", ""))
        rows = [("API 地址 (base_url)", v_base, False, 46),
                ("接口路径（默认 /v1/chat/completions）", v_path, False, 30),
                ("API Key",              v_key,  True, 46),
                ("模型名",               v_model, False, 30)]
        for i, (lab, var, hide, w) in enumerate(rows):
            ttk.Label(frm, text=lab, font=FONT).grid(row=i, column=0, padx=8, pady=6, sticky="w")
            ttk.Entry(frm, textvariable=var, font=FONT, width=w,
                      show="*" if hide else "").grid(row=i, column=1, padx=8, pady=6, sticky="w")
        # 状态标签 + 测试连接按钮
        lbl = ttk.Label(frm, text="", font=FONT, foreground="#2F855A")
        lbl.grid(row=len(rows), column=1, sticky="w", padx=8, pady=4)
        ttk.Button(frm, text="测试连接", width=10,
                   command=lambda k=kind, vb=v_base, vp=v_path, vk=v_key, vm=v_model, l=lbl:
                       self._test_api(k, vb.get(), vp.get(), vk.get(), vm.get(), l)
                   ).grid(row=len(rows), column=1, sticky="e", padx=8)
        if state_label_attr:
            setattr(self, state_label_attr, lbl)
        return (v_base, v_path, v_key, v_model)

    def _test_api(self, kind, base, path, key, model, status_label):
        """点“测试连接”立刻向 API 发一个最简 ping 请求，验证 Key/base/path/model。"""
        cur = ai_recognize.load_settings()
        if kind == "text":
            cur["text"] = {"api_base": base, "path": path, "api_key": key, "model": model}
        else:
            cur["vision"] = {"api_base": base, "path": path, "api_key": key, "model": model}
        ai_recognize.save_settings(cur["text"], cur["vision"])
        status_label.config(text="正在测试…", foreground="#2B6CB0")
        self.update_idletasks()

        def work():
            ok, msg = ai_recognize.test_api(kind)

            def update():
                status_label.config(text=("已配置 ✔ — " if ok else "未通过 — ") + msg,
                                     foreground="#2F855A" if ok else "#C53030")
                if not ok:
                    messagebox.showerror("测试失败", msg)
            self.after(0, update)

        threading.Thread(target=work, daemon=True).start()

    def _save_api(self):
        tb, tp, tk_, tm = self._api_text_vars
        vb, vp, vk, vm = self._api_vision_vars
        ai_recognize.save_settings(
            text={"api_base": tb.get(),  "path": tp.get(), "api_key": tk_.get(), "model": tm.get()},
            vision={"api_base": vb.get(), "path": vp.get(), "api_key": vk.get(), "model": vm.get()})
        # 状态刷新
        ok_text = ai_recognize.api_configured("text")
        ok_vis = ai_recognize.api_configured("vision")
        # 重新构建页面以同步状态（简单起见，提示用户重启或重开此页）
        messagebox.showinfo("完成",
            "API 设置已保存。\n"
            "  文本 API: %s\n"
            "  视觉 API: %s\n"
            "  切换到其他标签页再切回可刷新状态显示。" %
            ("已配置 ✔" if ok_text else "未配置（演示模式）",
             "已配置 ✔" if ok_vis else "未配置（扫描件/阅卷不可用）"))
