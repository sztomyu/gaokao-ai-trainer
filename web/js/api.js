/* api.js — 后端 API 封装与登录态管理 */
"use strict";

const Auth = {
  token: localStorage.getItem("gk_token") || "",
  user: null,
  init() {
    try { this.user = JSON.parse(localStorage.getItem("gk_user") || "null"); }
    catch (e) { this.user = null; }
  },
  save(token, user) {
    this.token = token; this.user = user;
    localStorage.setItem("gk_token", token);
    localStorage.setItem("gk_user", JSON.stringify(user));
  },
  clear() {
    this.token = ""; this.user = null;
    localStorage.removeItem("gk_token");
    localStorage.removeItem("gk_user");
  }
};
Auth.init();

/**
 * 统一请求入口。成功返回 data 字段；失败抛出带中文信息的 Error。
 * 401 时自动清登录态并跳回登录页。
 */
async function api(path, opts = {}) {
  const headers = Object.assign({}, opts.headers);
  if (Auth.token) headers["X-Auth-Token"] = Auth.token;
  let body = opts.body;
  if (body !== undefined && !(body instanceof FormData)) {
    headers["Content-Type"] = "application/json";
    body = JSON.stringify(body);
  }
  let resp;
  try {
    resp = await fetch(path, { method: opts.method || "GET", headers, body });
  } catch (e) {
    throw new Error("网络错误：无法连接服务器，请确认服务已启动");
  }
  let data = null;
  try { data = await resp.json(); } catch (e) { /* 非 JSON 响应 */ }
  if (resp.status === 401) {
    Auth.clear();
    if (typeof showLogin === "function") showLogin();
    throw new Error((data && data.error) || "登录已失效，请重新登录");
  }
  if (!data || data.ok !== true) {
    throw new Error((data && data.error) || ("请求失败（HTTP " + resp.status + "）"));
  }
  return data.data;
}

/* 全局元数据缓存（科目/题型/知识点树） */
const Meta = {
  data: null,
  async load(force) {
    if (this.data && !force) return this.data;
    this.data = await api("/api/meta");
    return this.data;
  },
  subjects() { return (this.data && this.data.subjects) || []; },
  qtypes() { return (this.data && this.data.qtypes) || []; },
  kpTree() { return (this.data && this.data.kp_tree) || {}; }
};
