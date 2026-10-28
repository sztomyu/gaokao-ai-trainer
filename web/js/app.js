/* app.js — SPA 入口：登录页 / 主框架 / 标签路由 / 答题全屏切换 */
"use strict";

const TEACHER_TABS = [
  ["import", "试卷导入"],
  ["bank", "题库管理"],
  ["dash", "数据看板"],
  ["kp", "知识点"],
  ["account", "账号与API"]
];
const STUDENT_TABS = [
  ["kp", "知识点训练"],
  ["mock", "模拟卷"],
  ["papers", "真题卷"],
  ["wrong", "错题本"],
  ["dash", "学情看板"]
];

/* ---------------- 屏幕切换 ---------------- */
function clearApp() {
  document.getElementById("app").innerHTML = "";
}

function showLogin() {
  clearApp();
  const wrap = h("div", "login-wrap");
  const card = h("div", "login-card");
  card.appendChild(h("h1", null, "高考学习训练系统"));
  card.appendChild(h("div", "sub", "Web 版 · 公式 / 化学方程式 / 图形全渲染"));
  const un = h("input");
  un.type = "text"; un.placeholder = "用户名"; un.autocomplete = "username";
  const pw = h("input");
  pw.type = "password"; pw.placeholder = "密码"; pw.autocomplete = "current-password";
  const errBox = h("div", "login-err");
  const btn = h("button", "btn", "登 录");
  card.appendChild(un); card.appendChild(pw); card.appendChild(errBox);
  card.appendChild(btn);
  card.appendChild(h("div", "login-hint",
    "内置账号：teacher / 123456（教师）　student / 123456（学生）"));
  wrap.appendChild(card);
  document.getElementById("app").appendChild(wrap);

  const doLogin = async () => {
    errBox.textContent = "";
    if (!un.value.trim() || !pw.value) { errBox.textContent = "请输入用户名和密码"; return; }
    btn.disabled = true; btn.textContent = "登录中…";
    try {
      const r = await api("/api/login", { method: "POST",
        body: { username: un.value.trim(), password: pw.value } });
      Auth.save(r.token, r.user);
      await Meta.load(true);
      showApp();
      toast("欢迎，" + (r.user.name || r.user.username), "ok");
    } catch (e) {
      errBox.textContent = e.message;
    } finally {
      btn.disabled = false; btn.textContent = "登 录";
    }
  };
  btn.onclick = doLogin;
  pw.addEventListener("keydown", e => { if (e.key === "Enter") doLogin(); });
  un.addEventListener("keydown", e => { if (e.key === "Enter") pw.focus(); });
  un.focus();
}

function logout() {
  Auth.clear();
  showLogin();
}

function showApp() {
  clearApp();
  const user = Auth.user;
  if (!user) { showLogin(); return; }
  const isTeacher = user.role === "teacher";
  const tabs = isTeacher ? TEACHER_TABS : STUDENT_TABS;
  const views = isTeacher ? TeacherViews : StudentViews;

  // 顶栏
  const top = h("div", "topbar");
  top.appendChild(h("span", "brand", "高考学习训练系统"));
  const navBox = h("div", "nav-tabs");
  top.appendChild(navBox);
  top.appendChild(h("span", "user-chip",
    (user.name || user.username) + "（" + (isTeacher ? "教师" : "学生") + "）"));
  const outBtn = h("button", "btn btn-ghost btn-sm", "退出登录");
  outBtn.onclick = logout;
  top.appendChild(outBtn);

  const page = h("div", "page");
  page.id = "page";
  document.getElementById("app").appendChild(top);
  document.getElementById("app").appendChild(page);

  const tabBtns = {};
  const activate = (key) => {
    Object.values(tabBtns).forEach(b => b.classList.remove("active"));
    tabBtns[key].classList.add("active");
    const v = views[key];
    if (!isTeacher) StudentHome.current = v;
    v.render(page);
  };
  tabs.forEach(([key, label]) => {
    const b = h("button", "tab", label);
    b.onclick = () => activate(key);
    tabBtns[key] = b;
    navBox.appendChild(b);
  });
  activate(tabs[0][0]);
}

/* ---------------- 答题全屏 ---------------- */
function showAnswerScreen() {
  document.getElementById("app").innerHTML = '<div id="answer-root"></div>';
}
function showMainScreen() {
  showApp();
}

/* ---------------- 启动 ---------------- */
document.addEventListener("DOMContentLoaded", async () => {
  try { await Meta.load(); } catch (e) { /* 服务未就绪时登录后再加载 */ }
  if (Auth.token && Auth.user) {
    showApp();
  } else {
    showLogin();
  }
});
