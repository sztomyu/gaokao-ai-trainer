#!/usr/bin/env node
/* dev.mjs — 开发启动器：解析 --port/--host 并透传给 server/app.py
   用法：npm run dev -- --port 7100 --host 127.0.0.1 */
import { spawn, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// ---- 解析命令行参数（支持 --port 7100 与 --port=7100 两种形式）----
const args = process.argv.slice(2);
const forward = [];
let port = process.env.PORT || "7100";
let host = process.env.HOST || "127.0.0.1";
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === "--port" || a === "-p") { port = args[++i] || port; }
  else if (a.startsWith("--port=")) { port = a.slice(7); }
  else if (a === "--host") { host = args[++i] || host; }
  else if (a.startsWith("--host=")) { host = a.slice(7); }
  else forward.push(a);
}
forward.push("--port", String(port), "--host", String(host));

// ---- 按顺序寻找 Python：PYTHON 环境变量 → python → python3 ----
const CANDIDATES = ["python", "python3"];

function findPython() {
  if (process.env.PYTHON) return process.env.PYTHON;
  for (const cmd of CANDIDATES) {
    const probe = spawnSync(cmd, ["--version"], { stdio: "ignore", shell: true });
    if (!probe.error && probe.status === 0) return cmd;
  }
  return "python";
}

const py = findPython();
console.log("=".repeat(56));
console.log(`  启动高考学习训练系统 Web 版`);
console.log(`  Python : ${py}`);
console.log(`  地址   : http://${host}:${port}`);
console.log("=".repeat(56));

const child = spawn(py, ["server/app.py", ...forward], {
  cwd: ROOT,
  stdio: "inherit",
});

// 父进程退出时带走子进程，避免残留后台服务
const killChild = () => { try { child.kill(); } catch (e) { /* 已退出 */ } };
process.on("SIGINT", killChild);
process.on("SIGTERM", killChild);
process.on("exit", killChild);

child.on("error", (e) => {
  console.error("启动失败：", e.message);
  process.exit(1);
});
child.on("exit", (code, signal) => {
  if (signal) process.exit(0);
  process.exit(code == null ? 0 : code);
});
