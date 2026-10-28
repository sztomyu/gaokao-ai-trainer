# -*- coding: utf-8 -*-
"""
main.py — 高考学习训练系统（简化版）程序入口
单机桌面应用：tkinter GUI + Excel 数据文件（无需 Web Server / 数据库）
首次运行自动初始化 data/ 下的 5 个 Excel 文件与示例数据。
默认账号：teacher/123456（教师）、student/123456（学生）
"""
import tkinter as tk
from tkinter import ttk, messagebox

import excel_store as store

FONT = ("Microsoft YaHei UI", 11)
FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")


class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("高考学习训练系统（简化版）— 登录")
        self.resizable(False, False)
        w, h = 420, 300
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, (sh - h) // 2 - 40))

        ttk.Label(self, text="高考学习训练系统", font=FONT_TITLE,
                  foreground="#1F3B66").pack(pady=(28, 2))
        ttk.Label(self, text="简化版 · 单机桌面应用", font=FONT,
                  foreground="#718096").pack()

        frm = ttk.Frame(self)
        frm.pack(pady=18)
        ttk.Label(frm, text="用户名", font=FONT).grid(row=0, column=0, padx=6, pady=8, sticky="e")
        ttk.Label(frm, text="密  码", font=FONT).grid(row=1, column=0, padx=6, pady=8, sticky="e")
        self.var_user = tk.StringVar()
        self.var_pwd = tk.StringVar()
        e1 = ttk.Entry(frm, textvariable=self.var_user, font=FONT, width=22)
        e2 = ttk.Entry(frm, textvariable=self.var_pwd, font=FONT, width=22, show="*")
        e1.grid(row=0, column=1)
        e2.grid(row=1, column=1)
        e2.bind("<Return>", lambda e: self._login())
        e1.focus_set()

        ttk.Button(self, text="登  录", command=self._login).pack(pady=4)
        ttk.Label(self, text="内置账号：teacher / 123456（教师）    student / 123456（学生）",
                  font=("Microsoft YaHei UI", 9), foreground="#A0AEC0").pack(side="bottom", pady=10)

    def _login(self):
        u = self.var_user.get().strip()
        p = self.var_pwd.get()
        if not u or not p:
            messagebox.showinfo("提示", "请输入用户名和密码")
            return
        user = store.verify_login(u, p)
        if not user:
            messagebox.showerror("登录失败", "用户名或密码错误，或账号已停用")
            return
        self.withdraw()
        if user["role"] == "teacher":
            from teacher_ui import TeacherWindow
            TeacherWindow(self, user)
        else:
            from student_ui import StudentWindow
            StudentWindow(self, user)


def main():
    store.init_all()
    app = LoginWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
