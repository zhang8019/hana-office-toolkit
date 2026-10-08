// index.js — Office 工具箱 App 入口。
// 一个 App 承担三件事：
//   1) 桥接本机 officecli（AI 友好的 Office CLI）为 Hana 工具：office_cli_run / office_cli_status
//   2) 旧格式文档批量转换：office_convert（.doc/.xls/.ppt → .docx/.xlsx/.pptx，走 MS Office COM）
//   3) 管理面板：状态诊断 + 转换操作，route=/panel.ts
//
// 依赖能力：app/process.spawn（拉起 officecli / PowerShell 子进程）、app/resources.read|write、
//           app/tools.expose-to-model、app/ui.open-external、app/ui.clipboard-write。

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile, spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { defineApp } from "./sdk/app-contract/server-client.js";
import {
  officecliCandidates,
  clip,
  parseVersion,
  hasShellMetachars,
  combine,
} from "./lib/office-core.js";

export const name = "office-toolkit";

const VERSION = "0.7.0";
const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const CONVERT_SCRIPT = path.join(SCRIPT_DIR, "scripts", "office_convert.ps1");
const ENV_PROBE = path.join(SCRIPT_DIR, "scripts", "env_probe.ps1");
const INSTALL_OFFICECLI = path.join(SCRIPT_DIR, "scripts", "install_officecli.ps1");

const TIMEOUT = { probe: 15_000, run: 60_000, runMax: 600_000, convert: 60 * 60_000, stalledFile: 15 * 60_000 };
const JOB_RETENTION_MS = 24 * 60 * 60 * 1000;
const JOB_MAX_OUTPUT = 2 * 1024 * 1024;
const JOB_MAX_STDERR = 80_000;

export default defineApp(async (sdk) => {
  await sdk.logger.info(`office-toolkit ${VERSION} loading`);

  // 转换任务的落盘目录：宿主授权给 App 的可写根，不碰系统临时目录。
  if (!sdk.dataDir) throw new Error("App dataDir is unavailable; cannot safely persist conversion jobs.");
  const workDir = path.join(sdk.dataDir, "jobs");
  const ensureWorkDir = () => fs.mkdirSync(workDir, { recursive: true });

  // ---------------------------------------------------------------- 进程执行

  function run(exe, argv, { timeoutMs = TIMEOUT.run, cwd } = {}) {
    return new Promise((resolve) => {
      const child = execFile(
        exe,
        argv,
        { timeout: timeoutMs, cwd, windowsHide: true, maxBuffer: 16 * 1024 * 1024 },
        (error, stdout, stderr) => {
          const base = { stdout: String(stdout ?? ""), stderr: String(stderr ?? "") };
          if (error) {
            resolve({
              ok: false,
              killed: !!error.killed,
              code: error.code ?? error.signal,
              enoent: error.code === "ENOENT",
              message: error.message ? String(error.message) : "",
              ...base,
            });
          } else {
            resolve({ ok: true, code: 0, killed: false, ...base });
          }
        },
      );
      child.on("error", () => {});
    });
  }

  // officecli 路径缓存：并发调用共享同一次解析。
  let cliPath = null;
  let cliPending = null;
  async function resolveCli() {
    if (cliPath) return cliPath;
    if (cliPending) return cliPending;
    const candidates = officecliCandidates(process.env, process.platform);
    cliPending = (async () => {
      try {
        const info = await sdk.process.resolveExecutable({ candidates });
        return info?.path || candidates.at(-1);
      } catch {
        return candidates.at(-1);
      }
    })()
      .then((p) => { cliPath = p; return p; })
      .finally(() => { cliPending = null; });
    return cliPending;
  }

  const describe = (r) => (r.killed ? "超时被终止" : `退出码 ${r.code ?? "未知"}`);

  async function cli(argv, options = {}) {
    const exe = await resolveCli();
    const r = await run(exe, argv, options);
    if (r.enoent) cliPath = null;
    return r;
  }

  // ---------------------------------------------------------------- 状态

  async function readStatus() {
    const exe = await resolveCli();
    const r = await cli(["--version"], { timeoutMs: TIMEOUT.probe });
    return {
      path: exe,
      installed: r.ok,
      version: r.ok ? parseVersion(r.stdout) || (r.stdout || "").split("\n")[0].trim() : null,
      error: r.ok ? null : combine(r.stderr, r.message) || describe(r),
    };
  }

  // ---------------------------------------------------------------- 转换

  // 转换用 shell：优先 PowerShell 7（pwsh），它按 UTF-8 读脚本；5.1 会把无 BOM 的 UTF-8 当 ANSI。
  // 这里先试已知绝对路径（WindowsApps 别名目录不一定能被 resolveExecutable 命中）。
  async function resolvePwsh() {
    const env = process.env;
    const prefs = [];
    if (env.ProgramFiles) prefs.push(`${env.ProgramFiles}\\PowerShell\\7\\pwsh.exe`);
    if (env.LOCALAPPDATA) prefs.push(`${env.LOCALAPPDATA}\\Microsoft\\WindowsApps\\pwsh.exe`);
    prefs.push("pwsh", "powershell.exe");
    for (const c of prefs) {
      try {
        const info = await sdk.process.resolveExecutable({ candidates: [c] });
        if (info?.path) return info.path;
      } catch { /* 继续下一个 */ }
    }
    return "powershell.exe";
  }

  // 转换任务状态只落在 App 自己的 dataDir；用户目录仍只由 PowerShell 子进程访问。
  const jobs = new Map();
  const activeChildByJob = new Map();
  const terminal = new Set(["completed", "failed", "cancelled", "interrupted"]);
  const statePath = (id) => path.join(workDir, `state-${id}.json`);
  const inputPath = (id) => path.join(workDir, `input-${id}.json`);
  const nowIso = () => new Date().toISOString();

  function publicJob(job) {
    const { child, timer, stallTimer, stdout, stderr, pid, cancelReason, ...visible } = job;
    return { ...visible, updatedAt: job.updatedAt };
  }
  function saveJob(job) {
    job.updatedAt = nowIso();
    try { fs.writeFileSync(statePath(job.jobId), JSON.stringify(publicJob(job)), "utf8"); }
    catch (error) { sdk.logger.warn(`Could not persist conversion job ${job.jobId}: ${error?.message ?? error}`); }
  }
  function jobSummary(job) {
    const { results, ...summary } = publicJob(job);
    return summary;
  }
  function listJobs() {
    return [...jobs.values()].sort((a, b) => b.createdAt.localeCompare(a.createdAt)).map(jobSummary);
  }
  function cleanupOldJobs() {
    const cutoff = Date.now() - JOB_RETENTION_MS;
    for (const [id, job] of jobs) {
      if (terminal.has(job.status) && Date.parse(job.updatedAt || job.createdAt) < cutoff) jobs.delete(id);
    }
    try {
      for (const name of fs.readdirSync(workDir)) {
        if (/^(state|input)-[0-9a-f-]+\.json$/i.test(name)) {
          const full = path.join(workDir, name);
          if (fs.statSync(full).mtimeMs < cutoff) fs.unlinkSync(full);
        }
      }
    } catch { /* best effort */ }
  }
  function recoverJobs() {
    ensureWorkDir();
    try {
      for (const name of fs.readdirSync(workDir)) {
        if (!/^state-[0-9a-f-]+\.json$/i.test(name)) continue;
        const full = path.join(workDir, name);
        const job = JSON.parse(fs.readFileSync(full, "utf8"));
        if (!job?.jobId || !job.status) continue;
        if (job.status === "running" || job.status === "queued" || job.status === "cancelling") {
          job.status = "interrupted";
          job.phase = "App 重启时失去任务跟踪";
          job.error = "App 重启/服务中断后，任务结果不确定。为防止 PID 重用误杀其它程序，不自动重跑或按旧 PID 强杀；请检查产物目录与 WINWORD 进程。";
          job.finishedAt = nowIso();
          job.updatedAt = job.finishedAt;
          fs.writeFileSync(full, JSON.stringify(publicJob(job)), "utf8");
          try { fs.unlinkSync(inputPath(job.jobId)); } catch {}
        }
        jobs.set(job.jobId, job);
      }
    } catch (error) { sdk.logger.warn(`Could not recover conversion jobs: ${error?.message ?? error}`); }
    cleanupOldJobs();
  }
  recoverJobs();

  function killProcessTree(child) {
    if (!child?.pid) return;
    if (process.platform === "win32") {
      const killer = spawn("taskkill.exe", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
      killer.on("error", () => { try { child.kill("SIGKILL"); } catch {} });
      killer.on("close", (code) => { if (code !== 0) { try { child.kill("SIGKILL"); } catch {} } });
    } else {
      try { child.kill("SIGKILL"); } catch {}
    }
  }

  // COM starts Word/Excel/PowerPoint outside the PowerShell process tree, so a taskkill /T
  // on PowerShell leaves the GUI process behind. Kill only the exact PID this job caused to
  // start; never scan-and-kill by image name, to avoid touching the user's own Office.
  function killPid(pid) {
    const id = Number(pid);
    if (!Number.isInteger(id) || id <= 0) return;
    if (process.platform === "win32") {
      const killer = spawn("taskkill.exe", ["/PID", String(id), "/F"], { windowsHide: true, stdio: "ignore" });
      killer.on("error", () => {});
    } else {
      try { process.kill(id, "SIGKILL"); } catch {}
    }
  }

  function resetStallTimer(job) {
    clearTimeout(job.stallTimer);
    if (job.status !== "running") return;
    job.stallTimer = setTimeout(() => {
      if (job.status !== "running") return;
      job.cancelReason = "stalled"; job.status = "cancelling";
      job.phase = "当前文件超过 15 分钟没有进度，正在终止子进程树";
      saveJob(job); killProcessTree(job.child); killPid(job.officePid);
    }, TIMEOUT.stalledFile);
  }

  function parseProgressLine(job, line) {
    const prefix = "OFFICE_PROGRESS:";
    if (!line.startsWith(prefix)) return;
    try {
      const p = JSON.parse(line.slice(prefix.length));
      if (p.phase === "env") {
        try { if (p.fileB64) job.envInfo = Buffer.from(String(p.fileB64), "base64").toString("utf8"); } catch {}
        return;
      }
      if (p.phase === "engine") {
        try { if (p.fileB64) job.engineInfo = Buffer.from(String(p.fileB64), "base64").toString("utf8"); } catch {}
        return;
      }
      const phaseNames = { env: "环境自检", enumerated: "已枚举文件", "word-open": "Word 打开文件中", "word-opened": "Word 已打开", "word-saving": "Word 保存中", "word-saved": "Word 已保存", "word-closed": "Word 已关闭", word: "Word 转换中", excel: "Excel 转换中", powerpoint: "PowerPoint 转换中" };
      job.phase = phaseNames[p.phase] || String(p.phase || job.phase);
      resetStallTimer(job);
      if (Number.isFinite(p.current) && Number.isFinite(p.total) && p.total > 0) {
        job.current = p.current;
        job.total = p.total;
        job.percent = Math.round((p.current / p.total) * 100);
      }
      if (p.fileB64) {
        try { job.currentFile = Buffer.from(String(p.fileB64), "base64").toString("utf8"); } catch {}
      }
      if (Number.isFinite(p.officePid) && p.officePid > 0) job.officePid = p.officePid;
      saveJob(job);
    } catch { /* 忽略格式异常，不影响转换 */ }
  }

  async function startConversion({ input, output, mode, recursive, confirmReplace, to, engine }) {
    const src = String(input || "").trim();
    if (!src) return { error: "缺少 input（待转换的文件或目录路径）。" };
    if (mode === "replace" && confirmReplace !== true) return { error: "replace 会删除源文件，必须先明确确认。确认后再传 confirmReplace=true。" };
    const alreadyRunning = [...jobs.values()].find((j) => ["running", "queued", "cancelling"].includes(j.status));
    if (alreadyRunning) return { error: `已有转换任务在运行（${alreadyRunning.jobId}），先等待或取消它。` };
    const m = ["keep", "backup", "replace"].includes(mode) ? mode : "keep";
    ensureWorkDir();
    cleanupOldJobs();
    const jobId = randomUUID();
    const inFile = inputPath(jobId);
    const job = {
      jobId, status: "queued", phase: "准备启动 PowerShell", percent: 0,
      current: 0, total: null, currentFile: null, error: null, results: null,
      okCount: 0, failCount: 0, mode: m, outDir: output ? String(output) : "",
      backupDir: "", input: src, createdAt: nowIso(), updatedAt: nowIso(), finishedAt: null,
    };
    jobs.set(jobId, job);
    fs.writeFileSync(inFile, JSON.stringify({ input: src, output: job.outDir, mode: m, recursive: recursive !== false, keepTimestamps: true, to: to ? String(to).replace(/^\./, "").toLowerCase() : "", engine: engine ? String(engine) : "auto" }), "utf8");
    saveJob(job);

    let pwsh;
    try { pwsh = await resolvePwsh(); }
    catch (error) {
      job.status = "failed"; job.phase = "无法定位 PowerShell"; job.error = String(error?.message ?? error); job.finishedAt = nowIso(); saveJob(job);
      try { fs.unlinkSync(inFile); } catch {}
      return { error: job.error };
    }
    if (terminal.has(job.status)) return { jobId, status: job.status, message: "任务在启动阶段已被取消。" };
    let child;
    try {
      child = spawn(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", CONVERT_SCRIPT, "-JobFile", inFile], { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    } catch (error) {
      job.status = "failed"; job.phase = "PowerShell 启动失败"; job.error = String(error?.message ?? error); job.finishedAt = nowIso(); saveJob(job);
      try { fs.unlinkSync(inFile); } catch {}
      return { error: job.error };
    }
    job.status = "running"; job.phase = "PowerShell 已启动，等待枚举文件"; job.child = child; job.pid = child.pid; job.stdout = ""; job.stderr = "";
    activeChildByJob.set(jobId, child);
    saveJob(job);

    let stderrBuffer = "";
    child.stdout.on("data", (chunk) => { job.stdout = (job.stdout + String(chunk)).slice(-JOB_MAX_OUTPUT); });
    child.stderr.on("data", (chunk) => {
      const textChunk = String(chunk);
      job.stderr = (job.stderr + textChunk).slice(-JOB_MAX_STDERR);
      stderrBuffer += textChunk;
      const lines = stderrBuffer.split(/\r?\n/);
      stderrBuffer = lines.pop() || "";
      for (const line of lines) parseProgressLine(job, line.trim());
    });
    child.on("error", (error) => {
      clearTimeout(job.timer);
      clearTimeout(job.stallTimer);
      job.child = null; job.pid = null; job.timer = null; job.stallTimer = null;
      job.status = "failed"; job.phase = "PowerShell 启动失败"; job.error = String(error?.message ?? error);
      job.finishedAt = nowIso(); saveJob(job);
      activeChildByJob.delete(jobId);
      try { fs.unlinkSync(inFile); } catch {}
    });
    resetStallTimer(job);
    job.timer = setTimeout(() => {
      if (terminal.has(job.status)) return;
      job.cancelReason = "timeout"; job.status = "cancelling"; job.phase = "转换超过 60 分钟，正在终止子进程树";
      saveJob(job); killProcessTree(child); killPid(job.officePid);
    }, TIMEOUT.convert);
    child.on("close", (code, signal) => {
      clearTimeout(job.timer);
      clearTimeout(job.stallTimer);
      activeChildByJob.delete(jobId);
      if (terminal.has(job.status)) {
        try { fs.unlinkSync(inFile); } catch {}
        return;
      }
      if (stderrBuffer) parseProgressLine(job, stderrBuffer.trim());
      job.child = null; job.pid = null; job.timer = null;
      let parsed = null;
      try { parsed = JSON.parse(String(job.stdout || "").replace(/^\uFEFF/, "").trim().split(/\r?\n/).at(-1)); } catch {}
      if (job.cancelReason === "user") {
        job.status = "cancelled"; job.phase = "已取消"; job.error = null;
      } else if (job.cancelReason === "timeout") {
        job.status = "failed"; job.phase = "超时结束"; job.error = "转换超过 60 分钟，已终止 PowerShell 子进程树。检查输出目录确认已完成的文件。";
      } else if (job.cancelReason === "stalled") {
        job.status = "failed"; job.phase = "单文件转换无进度"; job.error = "当前文件超过 15 分钟没有进度，PowerShell 子进程树与本次拉起的 Office 进程已被终止。请检查当前文件是否损坏、受保护或被 Office 锁定。";
      } else if (code !== 0) {
        job.status = "failed"; job.phase = "转换进程失败";
        job.error = clip(job.stderr || `PowerShell exited with code ${code ?? signal ?? "unknown"}`, 4000);
      } else if (!parsed) {
        job.status = "failed"; job.phase = "结果解析失败"; job.error = `PowerShell 没有返回有效 JSON。${clip(job.stderr, 2000)}`;
      } else if (parsed.error) {
        job.status = "failed"; job.phase = "转换失败"; job.error = String(parsed.error);
      } else {
        job.results = parsed.results || [];
        job.okCount = job.results.filter((x) => x.ok).length;
        job.failCount = job.results.length - job.okCount;
        job.total = parsed.total ?? job.results.length;
        job.percent = 100; job.outDir = parsed.output || job.outDir;
        job.backupDir = parsed.backupDir || "";
        job.status = "completed"; job.phase = "转换完成";
      }
      job.finishedAt = nowIso(); saveJob(job);
      if (job.status !== "completed") killPid(job.officePid);
      try { fs.unlinkSync(inFile); } catch {}
    });
    return { jobId, status: job.status, message: "转换任务已启动。使用 office_convert_status 查询进度，或 office_convert_cancel 取消。" };
  }

  async function cancelConversion(jobId) {
    const job = jobs.get(String(jobId || ""));
    if (!job) return { error: "未找到该转换任务。" };
    if (terminal.has(job.status)) return { error: `任务已结束（${job.status}），无法取消。`, job: publicJob(job) };
    const child = activeChildByJob.get(job.jobId);
    if (!child && job.status === "queued") {
      job.cancelReason = "user"; job.status = "cancelled"; job.phase = "排队阶段已取消"; job.finishedAt = nowIso(); saveJob(job);
      try { fs.unlinkSync(inputPath(job.jobId)); } catch {}
      return { jobId: job.jobId, status: "cancelled", message: "启动前已取消。" };
    }
    if (!child) return { error: "任务进程尚未就绪或已退出，暂时无法取消。", job: publicJob(job) };
    job.cancelReason = "user"; job.status = "cancelling"; job.phase = "正在取消并清理子进程树"; saveJob(job);
    killProcessTree(child);
    killPid(job.officePid);
    return { jobId: job.jobId, status: "cancelling", message: "已请求取消；稍后查询状态确认进程已退出。" };
  }


  // ---------------------------------------------------------------- 依赖自检

  // Other people install this App from the marketplace on machines that may have none of the
  // pieces we lean on. Report exactly what is missing and what to run, instead of failing late.
  const DEP_GUIDE = {
    libreoffice: {
      label: "LibreOffice",
      why: "转换引擎：旧格式现代化、转 PDF/HTML/CSV 等都靠它",
      required: "文档转换相关工具",
      how: "winget",
      winget: "TheDocumentFoundation.LibreOffice",
      url: "https://www.libreoffice.org/download/download-libreoffice/",
    },
    uv: {
      label: "uv / uvx",
      why: "运行随包托管的 timeverse-office-doc MCP（74 项文档工具）",
      required: "office_doc MCP 工具集",
      how: "winget",
      winget: "astral-sh.uv",
      url: "https://docs.astral.sh/uv/getting-started/installation/",
    },
    officecli: {
      label: "officecli",
      why: "docx/xlsx/pptx 的读写桥接",
      required: "office_cli_run / office_cli_status",
      how: "download",
      winget: "",
      url: "https://github.com/iOfficeAI/OfficeCLI",
    },
    officeCom: {
      label: "Office COM（WPS 或 MS Office）",
      why: "没有 LibreOffice 时的转换兜底（仅旧格式→现代格式）",
      required: "可选兜底",
      winget: "",
      url: "https://www.libreoffice.org/download/download-libreoffice/",
    },
  };

  async function probeDeps() {
    const pwsh = await resolvePwsh();
    const r = await run(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ENV_PROBE], { timeoutMs: 90_000 });
    if (!r.ok) return { error: `依赖自检失败（${describe(r)}）：${clip(combine(r.stdout, r.stderr), 1500)}` };
    try { return JSON.parse(String(r.stdout).replace(/^\uFEFF/, "").trim().split(/\r?\n/).at(-1)); }
    catch { return { error: `自检输出无法解析：${clip(r.stdout, 1000)}` }; }
  }

  function renderDeps(deps) {
    const lines = [];
    const missing = [];
    const add = (key, ok, detail) => {
      const g = DEP_GUIDE[key];
      lines.push(`${ok ? "✓" : "✗"} ${g.label}${detail ? ` — ${detail}` : ""}`);
      if (!ok) missing.push(key);
    };
    add("libreoffice", !!deps.libreoffice?.found, deps.libreoffice?.found ? `${deps.libreoffice.version || "版本未知"}${deps.libreoffice.path ? ` · ${deps.libreoffice.path}` : ""}` : "未找到");
    add("officecli", !!deps.officecli?.found, deps.officecli?.found ? `${deps.officecli.version || "版本未知"}${deps.officecli.path ? ` · ${deps.officecli.path}` : ""}` : "未找到");
    const hasUv = !!(deps.uv?.found || deps.uvx?.found);
    add("uv", hasUv, hasUv ? (deps.uvx?.path || deps.uv?.path) : "未找到");
    const com = deps.officeCom || {};
    lines.push(`${(com.wps || com.ms) ? "✓" : "✗"} ${DEP_GUIDE.officeCom.label} — ${com.wps ? "WPS " : ""}${com.ms ? "MS Office" : ""}${(com.wps || com.ms) ? "" : "未检测到"}`);

    if (missing.length) {
      lines.push("", "缺的东西与装法：");
      for (const key of missing) {
        const g = DEP_GUIDE[key];
        lines.push(`• ${g.label}（${g.required}）：${g.why}`);
        if (g.how) lines.push(`  自动：office_deps_install(target="${key}", confirm=true)`);
        lines.push(`  手动：${g.url}`);
      }
    } else {
      lines.push("", "依赖齐了，所有功能可用。");
    }
    return { text: lines.join("\n"), missing: missing.map((k) => DEP_GUIDE[k].label) };
  }

  // ---------------------------------------------------------------- 工具

  await sdk.tools.register({
    name: "office_deps_status",
    description:
      "检查本机文档工具链依赖：LibreOffice（转换引擎）、officecli（docx/xlsx/pptx 桥接）、uv/uvx（运行随包托管的 timeverse-office-doc MCP）、Office COM（WPS/MS Office 兜底）。返回每项是否就绪、路径/版本，以及缺失项的安装方式。首次使用或功能报「未找到」时先跑它。无需参数。",
    parameters: { type: "object", properties: {} },
    execute: async () => {
      const deps = await probeDeps();
      if (deps.error) return fail(deps.error);
      return text(`Office 工具箱 v${VERSION} 依赖自检\n\n${renderDeps(deps).text}`);
    },
  });

  await sdk.tools.register({
    name: "office_deps_install",
    description:
      "安装缺失的文档工具链依赖。target：libreoffice（转换引擎，走 winget）、uv（运行文档 MCP，走 winget）、officecli（从官方 GitHub Release 下载并校验 sha256 后安装到 %LOCALAPPDATA%\\OfficeCli）。可能耗时数分钟；winget 可能弹系统权限确认。执行前应向用户说明将要安装什么。",
    parameters: {
      type: "object",
      properties: {
        target: { type: "string", enum: ["libreoffice", "uv", "officecli"], description: "要安装的依赖。" },
        confirm: { type: "boolean", description: "必须显式传 true 才会执行安装。" },
      },
      required: ["target"],
    },
    execute: async ({ target, confirm }) => {
      const g = DEP_GUIDE[target];
      if (!g || !g.how) return fail("target 只能是 libreoffice、uv 或 officecli。");

      if (target === "officecli") {
        if (confirm !== true) return fail("将从 iOfficeAI/OfficeCLI 官方 Release 下载 officecli 并校验 sha256，安装到 %LOCALAPPDATA%\\OfficeCli。同意后传 confirm=true 重试。");
        const pwsh = await resolvePwsh();
        const r = await run(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", INSTALL_OFFICECLI], { timeoutMs: 20 * 60_000 });
        let parsed = null;
        try { parsed = JSON.parse(String(r.stdout).replace(/^\uFEFF/, "").trim().split(/\r?\n/).at(-1)); } catch {}
        if (!r.ok || !parsed) return fail(`安装 officecli 失败（${describe(r)}）：\n${clip(combine(r.stdout, r.stderr) || r.message, 2000)}`, `也可手动下载：${g.url}`);
        if (!parsed.ok) return fail(`安装 officecli 失败：${parsed.error}`, `也可手动下载：${g.url}`);
        const after = await probeDeps();
        return text(`officecli ${parsed.version} 已安装到 ${parsed.path}（sha256 校验通过）。\n\n重新自检：\n\n${after.error ? after.error : renderDeps(after).text}`);
      }

      if (confirm !== true) {
        return fail(`将执行：winget install --id ${g.winget} --silent。已安装请先用 office_deps_status 确认。取得用户同意后传 confirm=true 重试。`);
      }
      const args = ["install", "--id", g.winget, "--accept-package-agreements", "--accept-source-agreements", "--silent"];
      let r = await run("winget", args, { timeoutMs: 30 * 60_000 });
      if (r.enoent) r = await run("winget.exe", args, { timeoutMs: 30 * 60_000 });
      if (r.enoent) return fail("找不到 winget，请按链接手动安装：" + g.url);
      const body = clip(combine(r.stdout, r.stderr) || r.message, 3000);
      if (!r.ok) return fail(`安装 ${g.label} 失败（${describe(r)}）：\n${body}`, `也可手动安装：${g.url}`);
      const after = await probeDeps();
      return text(`winget 已执行完毕。重新自检：\n\n${after.error ? after.error : renderDeps(after).text}\n\n若仍显示未就绪，重开一个会话/重载应用后再试。`);
    },
  });

  await sdk.tools.register({
    name: "office_cli_status",
    description:
      "查看本机 officecli（AI 友好的 Office CLI）状态：可执行路径、版本、是否可用。用于确认环境，或在「office 命令报错」时先诊断。无需参数。",
    parameters: { type: "object", properties: {} },
    execute: async () => {
      const st = await readStatus();
      const lines = [`Office 工具箱 App v${VERSION}`];
      if (!st.installed) {
        lines.push(`officecli：不可用（探测路径：${st.path}）`);
        if (st.error) lines.push(`原因：${st.error}`);
        lines.push("下一步：确认已安装 officecli，或用 office_cli_run 指定其它路径。");
        return text(lines.join("\n"));
      }
      lines.push(`officecli：${st.version || "（版本未知）"}`, `路径：${st.path}`);
      lines.push("可用动作：office_cli_run 执行任意 officecli 子命令；office_convert 转换旧格式文档。");
      return text(lines.join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_cli_run",
    description:
      "执行一条 officecli 命令（参数以数组传入，不经 shell，不做通配/管道展开）。读操作可直接跑，例如 [\"get\",\"a.docx\",\"/body/p[1]\"]、[\"view\",\"a.docx\",\"outline\"]、[\"query\",\"a.xlsx\",\"cell\"]。写操作（set/add/remove/move/swap/raw-set/create/merge/import 等）请先向用户展示完整参数并获确认后再调用。涉及 .docx 的 refresh 需 Windows + Word。可用 cwd 指定工作目录。",
    parameters: {
      type: "object",
      properties: {
        args: { type: "array", items: { type: "string" }, description: "officecli 之后的参数数组，不含 officecli 本身，不含 shell 元字符。", minItems: 1 },
        cwd: { type: "string", description: "可选。运行目录（绝对路径）。" },
        timeoutMs: { type: "number", description: `可选超时（毫秒），默认 ${TIMEOUT.run}，上限 ${TIMEOUT.runMax}。` },
      },
      required: ["args"],
    },
    execute: async ({ args, cwd, timeoutMs }) => {
      if (!Array.isArray(args) || args.length === 0) return fail("args 必须是非空字符串数组。");
      const argv = args.map(String);
      if (argv[0]?.toLowerCase() === "officecli") argv.shift();
      if (hasShellMetachars(argv)) {
        return fail("参数含 shell 元字符，已拒绝。officecli 本身支持所需能力（如 --json、路径选择器），无需管道。");
      }
      const timeout = Math.min(Math.max(Number(timeoutMs) || TIMEOUT.run, 5_000), TIMEOUT.runMax);
      const r = await cli(argv, { timeoutMs: timeout, cwd });
      const label = `officecli ${argv.join(" ")}`;
      if (r.enoent) return fail(`未找到 officecli。先用 office_cli_status 诊断，或设置环境变量 OFFICECLI_PATH。`);
      if (r.ok) {
        return text(`${label} 执行成功：\n\n${combine(r.stdout, r.stderr && `[stderr]\n${r.stderr}`) || "（无输出）"}`);
      }
      return fail(
        `${label} 失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 8000)}`,
        "officecli 以退出码 1 表示失败、2 表示有警告但可能部分成功；报文里会带 error/code 字段。",
      );
    },
  });

  await sdk.tools.register({
    name: "office_convert",
    description:
      "启动文档转换并立即返回 jobId。缺省把旧格式转现代格式（.doc/.rtf→.docx、.xls→.xlsx、.ppt→.pptx）；传 to 可转成指定格式（pdf、docx、doc、xlsx、xls、pptx、ppt、html、txt、csv、rtf、odt、ods、odp、epub、md），此时会转换输入内所有可识别文档。转 PDF 等能力由本机 LibreOffice 提供（officecli 做不到）。使用 office_convert_status 查询进度/结果，office_convert_cancel 取消。mode=keep|backup|replace（replace 删除源文件，须先向用户确认）。",
    parameters: {
      type: "object",
      properties: {
        input: { type: "string", description: "待转换的文件或目录绝对路径。" },
        output: { type: "string", description: "可选输出目录，默认源目录。" },
        mode: { type: "string", enum: ["keep", "backup", "replace"], description: "原件处理：默认 keep。replace 删除源文件。" },
        recursive: { type: "boolean", description: "目录输入时是否递归，默认 true。" },
        confirmReplace: { type: "boolean", description: "仅 mode=replace 时必填 true；先向用户说明源文件会删除并取得明确同意。" },
        to: { type: "string", description: "可选目标格式。缺省时按旧格式默认映射（doc/rtf→docx、xls→xlsx、ppt→pptx）。指定后转换输入目录内所有可识别文档，支持：pdf、docx、doc、xlsx、xls、pptx、ppt、html、txt、csv、rtf、odt、ods、odp、epub、md。转 PDF 需本机 LibreOffice。" },
        engine: { type: "string", enum: ["auto", "libreoffice", "wps", "office"], description: "可选引擎，默认 auto（LibreOffice 优先）。" },
      },
      required: ["input"],
    },
    execute: async (params) => {
      const res = await startConversion(params || {});
      return res.error ? fail(res.error) : text(`${res.message}\n任务 ID：${res.jobId}`);
    },
  });

  await sdk.tools.register({
    name: "office_convert_status",
    description: "查询 Office 旧文档转换任务状态。传 jobId 查询单个任务；不传则列出最近 10 个任务。可看到阶段、百分比、成功/失败数和错误信息。",
    parameters: { type: "object", properties: { jobId: { type: "string", description: "可选任务 ID；省略时列出最近任务。" } }, required: [] },
    execute: async ({ jobId }) => {
      const selected = jobId ? jobs.get(String(jobId)) : null;
      if (jobId && !selected) return fail("未找到该任务 ID。");
      const rows = selected ? [publicJob(selected)] : listJobs().slice(0, 10);
      return text(JSON.stringify(rows, null, 2));
    },
  });

  await sdk.tools.register({
    name: "office_convert_cancel",
    description: "取消正在运行的 Office 转换任务，并尝试结束其 PowerShell 子进程树。取消后用 office_convert_status 确认终态；已经写出的文件不会自动删除。",
    parameters: { type: "object", properties: { jobId: { type: "string", description: "office_convert 返回的任务 ID。" } }, required: ["jobId"] },
    execute: async ({ jobId }) => {
      const res = await cancelConversion(jobId);
      return res.error ? fail(res.error) : text(JSON.stringify(res));
    },
  });

  // ---------------------------------------------------------------- 面板路由

  await sdk.routes.register((app) => {
    app.get("/status", async (c) => {
      const st = await readStatus();
      return c.json({ ok: true, app: { version: VERSION }, officecli: st });
    });

    app.post("/run", async (c) => {
      const body = await c.req.json().catch(() => ({}));
      const argv = Array.isArray(body?.args) ? body.args.map(String) : [];
      if (argv.length === 0) return c.json({ ok: false, error: "args 为空" }, 400);
      if (hasShellMetachars(argv)) return c.json({ ok: false, error: "参数含 shell 元字符" }, 400);
      const r = await cli(argv, { timeoutMs: TIMEOUT.run, cwd: body?.cwd ? String(body.cwd) : undefined });
      return c.json({ ok: r.ok, code: r.code ?? null, stdout: clip(r.stdout, 20000), stderr: clip(r.stderr, 20000) });
    });

    app.get("/deps", async (c) => {
      const deps = await probeDeps();
      if (deps.error) return c.json({ ok: false, error: deps.error }, 500);
      const rendered = renderDeps(deps);
      return c.json({ ok: true, deps, missing: rendered.missing });
    });

    app.get("/jobs", (c) => c.json({ ok: true, jobs: listJobs().slice(0, 20) }));

    app.get("/jobs/:jobId", (c) => {
      const job = jobs.get(c.req.param("jobId"));
      return job ? c.json({ ok: true, job: publicJob(job) }) : c.json({ ok: false, error: "未找到任务。" }, 404);
    });

    app.post("/jobs/:jobId/cancel", async (c) => {
      const res = await cancelConversion(c.req.param("jobId"));
      return res.error ? c.json({ ok: false, error: res.error, job: res.job }, 409) : c.json({ ok: true, ...res });
    });

    app.post("/convert", async (c) => {
      const body = await c.req.json().catch(() => ({}));
      const res = await startConversion({ input: body?.input, output: body?.output, mode: body?.mode, recursive: body?.recursive, confirmReplace: body?.confirmReplace, to: body?.to, engine: body?.engine });
      return res.error ? c.json({ ok: false, error: res.error }, 400) : c.json({ ok: true, ...res });
    });
  });

  // Surface missing dependencies at load time, so a broken install shows up in the log instead
  // of only when a tool is called. The probe must never block or fail the load.
  try {
    const deps = await probeDeps();
    if (deps.error) {
      await sdk.logger.warn(`office-toolkit dependency probe failed: ${deps.error}`);
    } else {
      const r = renderDeps(deps);
      if (r.missing.length) await sdk.logger.warn(`office-toolkit missing dependencies: ${r.missing.join(", ")} (run office_deps_status)`);
      else await sdk.logger.info("office-toolkit dependencies: all present");
    }
  } catch (error) {
    await sdk.logger.warn(`office-toolkit dependency probe error: ${error?.message ?? error}`);
  }

  await sdk.logger.info(`office-toolkit ${VERSION} ready: tools=office_deps_status/office_deps_install/office_cli_status/office_cli_run/office_convert/office_convert_status/office_convert_cancel, routes=status/deps/run/convert/jobs`);
});

// ---------------------------------------------------------------- 工具返回封装

function text(s) {
  return { content: [{ type: "text", text: String(s) }] };
}
function fail(message, hint) {
  return { content: [{ type: "text", text: hint ? `${message}\n\n提示：${hint}` : String(message) }], isError: true };
}

