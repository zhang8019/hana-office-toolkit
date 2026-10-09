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

const VERSION = "0.9.0";
const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const CONVERT_SCRIPT = path.join(SCRIPT_DIR, "scripts", "office_convert.ps1");
const COMPOSE_CAPTION = path.join(SCRIPT_DIR, "scripts", "compose_caption.ps1");
const INSERT_ROWS = path.join(SCRIPT_DIR, "scripts", "docx_insert_rows.ps1");
// Accepted values for the `to` / `engine` options (keep in sync with the tool schema).
const ALLOWED_TARGETS = ["pdf", "docx", "doc", "xlsx", "xls", "pptx", "ppt", "html", "txt", "csv", "rtf", "odt", "ods", "odp", "epub", "md"];
const ALLOWED_ENGINES = ["auto", "office", "wps", "libreoffice"];
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
    for (const base of [env.ProgramFiles, env.ProgramW6432, "C:\\Program Files"]) {
      if (base) prefs.push(`${base}\\PowerShell\\7\\pwsh.exe`);
    }
    for (const base of [env.LOCALAPPDATA, env.USERPROFILE ? `${env.USERPROFILE}\\AppData\\Local` : ""]) {
      if (base) prefs.push(`${base}\\Microsoft\\WindowsApps\\pwsh.exe`);
    }
    prefs.push("pwsh");
    for (const c of prefs) {
      try {
        const info = await sdk.process.resolveExecutable({ candidates: [c] });
        if (info?.path) return info.path;
      } catch { /* 继续下一个 */ }
    }
    // resolveExecutable can reject the WindowsApps app-exec alias, which is the usual way pwsh is
    // installed on Windows. Verify the bare name by actually running it before giving up: falling
    // through to 5.1 silently changes quoting and encoding behaviour for the whole pipeline.
    const probe = await run("pwsh", ["-NoProfile", "-NonInteractive", "-Command", "$PSVersionTable.PSVersion.Major"], { timeoutMs: 10_000 });
    if (probe.ok && String(probe.stdout || "").trim().startsWith("7")) {
      await sdk.logger.info("office-toolkit: using pwsh (resolved by probe)");
      return "pwsh";
    }
    await sdk.logger.warn("office-toolkit: pwsh not found, falling back to Windows PowerShell 5.1");
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
    const { child, timer, stallTimer, stdout, stderr, pid, cancelReason, procsBefore, ...visible } = job;
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
      saveJob(job); killProcessTree(job.child); killPid(job.officePid); killPid(job.workerPid);
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
      if (Number.isFinite(p.workerPid) && p.workerPid > 0) job.workerPid = p.workerPid;
      saveJob(job);
    } catch { /* 忽略格式异常，不影响转换 */ }
  }

  // A wedged job must not block the queue forever: if the running job has had no progress for a
  // while, cancel it and let the new one through (previously the user had to cancel by hand).
  const STALE_JOB_MS = 5 * 60_000;
  async function clearStalledJob() {
    const active = [...jobs.values()].filter((j) => ["running", "queued", "cancelling"].includes(j.status));
    if (active.length === 0) return null;
    const running = active.find((j) => j.status === "running");
    if (!running) return active[0];
    const idleMs = Date.now() - Date.parse(running.updatedAt || running.createdAt);
    if (idleMs <= STALE_JOB_MS) return running;
    sdk.logger.warn(`office-toolkit: job ${running.jobId} idle ${Math.round(idleMs / 1000)}s, auto-cancelling to unblock the queue`);
    await cancelConversion(running.jobId);
    for (let i = 0; i < 20; i++) {
      await new Promise((r) => setTimeout(r, 500));
      if (!["running", "queued", "cancelling"].includes(running.status)) break;
    }
    return ["running", "queued", "cancelling"].includes(running.status) ? running : null;
  }

  // Settings declared in manifest contributes.settings act as defaults when a call omits the value.
  async function readConfigDefaults() {
    const pick = async (key) => { try { const v = await sdk.config.get(key); return v; } catch { return undefined; } };
    return { engine: await pick("engine"), to: await pick("targetFormat"), mode: await pick("mode") };
  }

  async function startConversion({ input, output, mode, recursive, confirmReplace, to, engine }) {
    const src = String(input || "").trim();
    if (!src) return { error: "缺少 input（待转换的文件或目录路径）。" };
    const cfg = await readConfigDefaults();
    const effMode = (mode && String(mode)) || (cfg.mode && String(cfg.mode)) || "keep";
    const effTo = (to && String(to)) || (cfg.to && String(cfg.to)) || "";
    const effEngine = (engine && String(engine)) || (cfg.engine && String(cfg.engine)) || "auto";
    // Reject nonsense at submit time. A bad `to` used to sail through and stall a job mid-flight,
    // and since only one job runs at a time that wedged every later request.
    const normTo = effTo ? String(effTo).replace(/^\./, "").toLowerCase() : "";
    const normEngine = String(effEngine).toLowerCase();
    if (normTo && !ALLOWED_TARGETS.includes(normTo)) {
      return { error: `不支持的目标格式“${effTo}”。可选：${ALLOWED_TARGETS.join("、")}。` };
    }
    if (!ALLOWED_ENGINES.includes(normEngine)) {
      return { error: `不支持的引擎“${effEngine}”。可选：${ALLOWED_ENGINES.join("、")}。` };
    }
    if (effMode === "replace" && confirmReplace !== true) return { error: "replace 会删除源文件，必须先明确确认。确认后再传 confirmReplace=true。" };
    const blocker = await clearStalledJob();
    if (blocker) return { error: `已有转换任务在运行（${blocker.jobId}），先等待或取消它。` };
    const m = ["keep", "backup", "replace"].includes(effMode) ? effMode : "keep";
    ensureWorkDir();
    cleanupOldJobs();
    const jobId = randomUUID();
    const inFile = inputPath(jobId);
    const job = {
      jobId, status: "queued", phase: "准备启动 PowerShell", percent: 0,
      current: 0, total: null, currentFile: null, error: null, results: null,
      okCount: 0, skippedCount: 0, failCount: 0, mode: m, outDir: output ? String(output) : "",
      backupDir: "", input: src, createdAt: nowIso(), updatedAt: nowIso(), finishedAt: null,
    };
    jobs.set(jobId, job);
    fs.writeFileSync(inFile, JSON.stringify({ input: src, output: job.outDir, mode: m, recursive: recursive !== false, keepTimestamps: true, to: normTo, engine: normEngine }), "utf8");
    saveJob(job);

    let pwsh;
    try { pwsh = await resolvePwsh(); }
    catch (error) {
      job.status = "failed"; job.phase = "无法定位 PowerShell"; job.error = String(error?.message ?? error); job.finishedAt = nowIso(); saveJob(job);
      try { fs.unlinkSync(inFile); } catch {}
      return { error: job.error };
    }
    if (terminal.has(job.status)) return { jobId, status: job.status, message: "任务在启动阶段已被取消。" };
    // Take a snapshot ONLY. Do not kill anything here: see listHeadlessOfficeProcs above.
    job.procsBefore = (await listOfficeProcs()).map((p) => p.pid);
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
      saveJob(job); killProcessTree(child); killPid(job.officePid); killPid(job.workerPid);
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
        job.skippedCount = job.results.filter((x) => x.skipped).length;
        job.failCount = job.results.length - job.okCount - job.skippedCount;
        job.total = parsed.total ?? job.results.length;
        job.percent = 100; job.outDir = parsed.output || job.outDir;
        job.backupDir = parsed.backupDir || "";
        job.status = "completed"; job.phase = "转换完成";
      }
      job.finishedAt = nowIso(); saveJob(job);
      if (job.status !== "completed") { killPid(job.officePid); killPid(job.workerPid); }
      try { fs.unlinkSync(inFile); } catch {}
      void sweepTemp();
      // Kill only the processes THIS job spawned (snapshot difference). Nothing is killed by guesswork.
      void killJobSpawnedProcs(job)
        .then((n) => { if (n > 0) sdk.logger.info(`office-toolkit: post-run cleanup removed ${n} Office process(es) spawned by job ${job.jobId}`); })
        .catch(() => {});
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
    killPid(job.workerPid);
    return { jobId: job.jobId, status: "cancelling", message: "已请求取消；稍后查询状态确认进程已退出。" };
  }


  // ---------------------------------------------------------------- 依赖自检

  // Other people install this App from the marketplace on machines that may have none of the
  // pieces we lean on. Report exactly what is missing and what to run, instead of failing late.
  const DEP_GUIDE = {
    officeMs: {
      label: "MS Office（Word / Excel / PowerPoint）",
      why: "首选转换引擎：旧格式→现代格式、导出 PDF，能保住原件版式",
      required: "转换（首选）",
      how: "",
      winget: "",
      url: "https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products",
    },
    officeWps: {
      label: "WPS Office",
      why: "第二顺位转换引擎，能力同 MS Office",
      required: "转换（次选）",
      how: "",
      winget: "",
      url: "https://www.wps.com/download/",
    },
    libreoffice: {
      label: "LibreOffice",
      why: "回落引擎：html / csv / txt / rtf / odt / ods / odp / epub / md 只有它能输出",
      required: "额外目标格式",
      how: "winget",
      winget: "TheDocumentFoundation.LibreOffice",
      url: "https://www.libreoffice.org/download/download-libreoffice/",
    },
    officecli: {
      label: "officecli",
      why: "docx/xlsx/pptx 的读写桥接（只认新版 DrawingML，看不见老式 VML 浮动图）",
      required: "office_cli_run / office_cli_status",
      how: "download",
      winget: "",
      url: "https://github.com/iOfficeAI/OfficeCLI",
    },
    uv: {
      label: "uv / uvx",
      why: "运行随包托管的 timeverse-office-doc MCP（74 项文档工具）",
      required: "office_doc MCP 工具集",
      how: "winget",
      winget: "astral-sh.uv",
      url: "https://docs.astral.sh/uv/getting-started/installation/",
    },
  };

  // Housekeeping: the conversion script cleans the transient %TEMP% dirs it creates. Called on
  // load and after every finished job so repeated runs do not accumulate junk.
  async function sweepTemp() {
    try {
      const pwsh = await resolvePwsh();
      const r = await run(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", CONVERT_SCRIPT, "-SweepTemp"], { timeoutMs: 60_000 });
      return r.ok;
    } catch { return false; }
  }

  // Zombie-process control. Office/WPS automation leaks instances (Microsoft documents this for
  // server-side automation: once instances leak, later calls hang). We snapshot the Office/soffice
  // processes before a job and kill only the ones that appeared during it, so a user's own open
  // Word/Excel is never touched.
  async function listOfficeProcs() {
    try {
      const pwsh = await resolvePwsh();
      const r = await run(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", CONVERT_SCRIPT, "-ListOfficePids"], { timeoutMs: 60_000 });
      if (!r.ok) return [];
      const txt = String(r.stdout).replace(/^\uFEFF/, "").trim();
      if (!txt) return [];
      const parsed = JSON.parse(txt.split(/\r?\n/).at(-1));
      // Defensive: PowerShell can hand back a nested array for a single result; flatten and drop junk.
      const flat = Array.isArray(parsed) ? parsed.flat(Infinity) : [parsed];
      return flat.filter((p) => p && typeof p === "object" && Number(p.pid) > 0);
    } catch { return []; }
  }

  async function killJobSpawnedProcs(job) {
    const before = new Set(job.procsBefore || []);
    const after = await listOfficeProcs();
    let n = 0;
    for (const p of after) {
      if (before.has(p.pid)) continue;
      killPid(p.pid);
      n++;
    }
    return n;
  }

  // Read-only companion to killJobSpawnedProcs. Earlier versions of this app KILLED every Office/
  // WPS process that had no main window title, on the theory that "no title = automation leftover".
  // That theory is wrong and destructive: WPS runs several wps.exe helpers, and an editor window
  // shows no title while a large document is still loading. Killing those takes down a real editing
  // session (observed 2026-10-09: a user's WPS session died and produced a recovery prompt right
  // after a pre-run sweep). We now only REPORT candidates; nothing is killed on a guess.
  async function listHeadlessOfficeProcs() {
    const procs = await listOfficeProcs();
    return procs.filter((p) => !String(p.title || "").trim());
  }

  // One cleanup routine behind both the tool and the panel button: temp dirs + recorded child
  // processes + headless Office/soffice zombies.
  async function runCleanup() {
    const swept = await sweepTemp();
    const cutoff = Date.now() - 60 * 60_000;
    let killed = 0;
    for (const job of jobs.values()) {
      if (!terminal.has(job.status)) continue;
      if (Date.parse(job.finishedAt || job.updatedAt || job.createdAt) < cutoff) continue;
      // Recorded PIDs from finished jobs; most have already exited, so this counts attempts rather
      // than confirmed kills. The meaningful figure is `zombies` (headless leftovers actually found).
      killed += [job.officePid, job.workerPid].filter((p) => Number.isInteger(Number(p)) && Number(p) > 0).length;
      killPid(job.officePid);
      killPid(job.workerPid);
    }
    const candidates = await listHeadlessOfficeProcs();
    return { swept, killed, candidates };
  }

  async function probeDeps() {
    const pwsh = await resolvePwsh();
    const r = await run(pwsh, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ENV_PROBE], { timeoutMs: 90_000 });
    if (!r.ok) return { error: `依赖自检失败（${describe(r)}）：${clip(combine(r.stdout, r.stderr), 1500)}` };
    try { return JSON.parse(String(r.stdout).replace(/^\uFEFF/, "").trim().split(/\r?\n/).at(-1)); }
    catch { return { error: `自检输出无法解析：${clip(r.stdout, 1000)}` }; }
  }

  function renderDeps(deps) {
    const lines = [];
    const installable = [];
    const detail = (o) => (o?.found ? `${o.version || "版本未知"}${o.path ? ` · ${o.path}` : ""}` : "未检测到");
    const row = (key, found, text) => {
      const g = DEP_GUIDE[key];
      lines.push(`${found ? "✓" : "✗"} ${g.label} — ${text}`);
      if (!found && g.how) installable.push(key);
    };
    const ms = deps.officeMs || {};
    const wps = deps.officeWps || {};
    const lo = deps.libreoffice || {};
    row("officeMs", !!ms.found, detail(ms));
    row("officeWps", !!wps.found, detail(wps));
    row("libreoffice", !!lo.found, detail(lo));
    row("officecli", !!deps.officecli?.found, detail(deps.officecli));
    const hasUv = !!(deps.uv?.found || deps.uvx?.found);
    row("uv", hasUv, hasUv ? (deps.uvx?.path || deps.uv?.path) : "未检测到");

    lines.push("");
    if (ms.found || wps.found) {
      lines.push("转换：可用（旧格式→现代格式、导出 PDF；引擎顺序 MS Office → WPS → LibreOffice）。");
      if (lo.found) lines.push("额外目标格式（html / csv / txt / rtf / odt / ods / odp / epub / md）由 LibreOffice 提供。");
    } else if (lo.found) {
      lines.push("转换：仅 LibreOffice。能转，但某些老文档版式可能失真（如图片溢出页外）。建议装 MS Office 或 WPS 以提升保真度。");
    } else {
      lines.push("转换：不可用。需要 MS Office / WPS 之一，或安装 LibreOffice。");
    }

    const manual = [];
    if (!ms.found) manual.push("officeMs");
    if (!wps.found) manual.push("officeWps");
    if (manual.length) {
      lines.push("", "以下为商业软件，本 App 不代装：");
      for (const key of manual) lines.push(`• ${DEP_GUIDE[key].label}：${DEP_GUIDE[key].url}`);
    }
    if (installable.length) {
      lines.push("", "可自动安装的缺失项：");
      for (const key of installable) {
        const g = DEP_GUIDE[key];
        lines.push(`• ${g.label}（${g.required}）：${g.why}`);
        lines.push(`  自动：office_deps_install(target="${key}", confirm=true)`);
        lines.push(`  手动：${g.url}`);
      }
    }
    return { text: lines.join("\n"), missing: installable.map((k) => DEP_GUIDE[k].label) };
  }

  // ---------------------------------------------------------------- 工具

  await sdk.tools.register({
    name: "office_cleanup",
    description:
      "清理本 App 留下的临时文件与子进程。① 清扫 %TEMP% 下过期的运行目录（仅删一小时前的）；② 按任务记录结束本 App 自己起过的、已结束任务遗留的子进程。**不再按“无窗口标题”猜测并杀进程**：那条规则会误杀正在加载大文件的 WPS/Office 编辑会话（已出现真实事故），现在只报告可疑项、由你决定。confirm 必须显式传 true。",
    parameters: { type: "object", properties: { confirm: { type: "boolean", description: "必须显式传 true 才会执行。" } } },
    execute: async ({ confirm }) => {
      if (confirm !== true) return fail("将清扫临时目录，并按记录结束已结束任务遗留的子进程（仅限提前一小时内结束的任务，降低 PID 重用误杀风险）。同意后传 confirm=true 重试。");
      const r = await runCleanup();
      const cand = r.candidates || [];
      const candText = cand.length
        ? `另发现 ${cand.length} 个无窗口标题的 Office/WPS 进程（${cand.map((p) => `${p.name}#${p.pid}`).join(", ")}），**未动它们**——无标题不等于残留，可能是正在加载文档的编辑会话；确认后请手动处理。`
        : "未发现无窗口标题的 Office/WPS 进程。";
      return text(`清理完成：临时目录清扫${r.swept ? "已执行" : "跳过（脚本不可用）"}；按记录复核了 ${r.killed} 个历史子进程 PID（多数已自行退出）。${candText}`);
    },
  });

  await sdk.tools.register({
    name: "office_deps_status",
    description:
      "检查本机文档工具链依赖（按引擎优先级排序）：MS Office（首选转换引擎）、WPS Office（次选）、LibreOffice（回落引擎，负责 html/csv/txt 等额外目标格式）、officecli（docx/xlsx/pptx 读写桥接）、uv/uvx（运行随包托管的 timeverse-office-doc MCP）。返回每项是否就绪、路径/版本，以及转换当前是否可用、缺失项怎么装。首次使用或功能报错时先跑它。无需参数。",
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
      "执行一条 officecli 命令（参数以数组传入，不经 shell，不做通配/管道展开）。读操作可直接跑，例如 [\"get\",\"a.docx\",\"/body/p[1]\"]、[\"view\",\"a.docx\",\"outline\"]、[\"query\",\"a.xlsx\",\"cell\"]。写操作（set/add/remove/move/swap/raw-set/create/merge/import 等）请先向用户展示完整参数并获确认后再调用。涉及 .docx 的 refresh 需 Windows + Word。可用 cwd 指定工作目录。\n\n已知边界：officecli 只认新版 DrawingML，对老式 VML 浮动图（w:pict / v:shape）不可见——view outline 会报 0 images、query picture 返回空。遇到这种情况不要断言「文档里没图」，改用 raw / raw-set 读取，或交给 office_convert 让 Word/WPS 处理。",
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
      "启动文档转换并立即返回 jobId。引擎自动顺序为 MS Office → WPS → LibreOffice（Word/WPS 能保住原件版式，LibreOffice 在某些老文档上会重排失真；COM 引擎由 explorer 派生的非提升子进程执行，失败自动回落 LibreOffice）。缺省把旧格式转现代格式（.doc/.rtf→.docx、.xls→.xlsx、.ppt→.pptx）；传 to 可转成指定格式（pdf、docx、doc、xlsx、xls、pptx、ppt、html、txt、csv、rtf、odt、ods、odp、epub、md）。目标是 pdf 时走 Word/WPS 导出，html/csv/txt 等仅 LibreOffice 能做。结果里会报出实际使用的引擎。使用 office_convert_status 查询进度/结果，office_convert_cancel 取消。mode=keep|backup|replace（replace 删除源文件，须先向用户确认）。",
    parameters: {
      type: "object",
      properties: {
        input: { type: "string", description: "待转换的文件或目录绝对路径。" },
        output: { type: "string", description: "可选输出目录，默认源目录。" },
        mode: { type: "string", enum: ["keep", "backup", "replace"], description: "原件处理：默认 keep。replace 删除源文件。" },
        recursive: { type: "boolean", description: "目录输入时是否递归，默认 true。" },
        confirmReplace: { type: "boolean", description: "仅 mode=replace 时必填 true；先向用户说明源文件会删除并取得明确同意。" },
        to: { type: "string", description: "可选目标格式。缺省时按旧格式默认映射（doc/rtf→docx、xls→xlsx、ppt→pptx）。指定后转换输入目录内所有可识别文档，支持：pdf、docx、doc、xlsx、xls、pptx、ppt、html、txt、csv、rtf、odt、ods、odp、epub、md。转 PDF 需本机 LibreOffice。" },
        engine: { type: "string", enum: ["auto", "office", "wps", "libreoffice"], description: "可选引擎，默认 auto（MS Office → WPS → LibreOffice）。" },
      },
      required: ["input"],
    },
    execute: async (params) => {
      const res = await startConversion(params || {});
      return res.error ? fail(res.error) : text(`${res.message}\n任务 ID：${res.jobId}`);
    },
  });

  await sdk.tools.register({
    name: "office_images",
    description:
      "列出 .docx 里的所有图片，包含 officecli 自己看不见的老式 VML 浮动图。officecli 的 query picture / view outline 只认新版 DrawingML，w:pict / v:shape 这类老式浮动图会被它漏掉，据此判断“文档里没图”是错的。本工具直接解析文档 XML（经 officecli raw），分别统计老式 VML 与新式 DrawingML，并给出每张图的尺寸、位置与描边。只读，不修改文件。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: ".docx 文件路径。" },
        verbose: { type: "boolean", description: "true 时逐张列出（id/尺寸/位置/描边/关系 id）。默认只给汇总。" },
        maxItems: { type: "number", description: "verbose 时最多列多少张，默认 40。" },
        contains: { type: "string", description: "verbose 时只列包含这个字串的图形（匹配 id / style / 标题），便于在大文档里直奔目标。" },
      },
      required: ["path"],
    },
    execute: async ({ path: p, verbose, maxItems, contains }) => {
      const file = String(p || "").trim();
      if (!file) return fail("缺少 path（.docx 文件路径）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx；老式 .doc 请先用 office_convert 转成 .docx。");
      const r = await cli(["raw", file, "/document"], { timeoutMs: TIMEOUT.run });
      if (r.enoent) return fail("未找到 officecli。先用 office_cli_status 诊断。");
      if (!r.ok) return fail(`读取文档 XML 失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 2000)}`);
      const xml = String(r.stdout || "");
      if (!xml.trim()) return fail("officecli raw 返回为空，无法解析。");

      const hits = (re) => (xml.match(re) || []).length;
      const tags = (re) => xml.match(re) || [];
      const attr = (tag, name) => {
        const m = tag.match(new RegExp(name.replace(/[:.]/g, "\\$&") + '="([^"]*)"', "i"));
        return m ? m[1] : "";
      };
      const emu2cm = (v) => {
        const n = Number(v);
        return Number.isFinite(n) && n > 0 ? `${(n / 360000).toFixed(2)}cm` : "";
      };

      // Legacy VML. \b after "shape" keeps v:shapetype (18 of 20 in a real sample) out of the count.
      const vShapes = tags(/<v:shape\b[^>]*>/gi);
      const vImages = tags(/<v:imagedata\b[^>]*>/gi);
      const pictCount = hits(/<w:pict\b/gi);
      const groupCount = hits(/<v:group\b/gi);
      // Modern DrawingML.
      const drawings = hits(/<w:drawing\b/gi);
      const anchors = tags(/<wp:anchor\b[^>]*>/gi);
      const inlines = hits(/<wp:inline\b/gi);
      const extents = tags(/<wp:extent\b[^>]*>/gi);
      const blips = tags(/<a:blip\b[^>]*>/gi);

      const vmlVis = vShapes.filter((t) => {
        const st = attr(t, "style").toLowerCase();
        return st.includes("position:absolute") || st.includes("margin-left") || st.includes("margin-top");
      }).length;

      const lines = [];
      lines.push(`${file}`);
      lines.push(`老式 VML：v:shape ${vShapes.length} 个（其中浮动定位 ${vmlVis} 个），w:pict ${pictCount} 个，v:imagedata ${vImages.length} 个，v:group ${groupCount} 个`);
      lines.push(`新式 DrawingML：w:drawing ${drawings} 个（浮动 wp:anchor ${anchors.length}、内联 wp:inline ${inlines}），a:blip ${blips.length} 个`);
      lines.push("");
      if (vShapes.length > 0) {
        lines.push("注意：officecli 自己的 query picture 只能看到新式的 " + drawings + " 个，上面 " + vShapes.length + " 个 VML 图形它看不见。");
        lines.push("不要据此判断文档里没有图；要改这类图请用 raw / raw-set，或交给 office_convert 让 Word/WPS 处理。");
      } else if (drawings === 0) {
        lines.push("两种图片都没有计数到，可以认为文档内确实无图。");
      }

      if (verbose) {
        const limit = Number.isFinite(Number(maxItems)) && Number(maxItems) > 0 ? Number(maxItems) : 40;
        const needle = String(contains || "").trim().toLowerCase();
        const keep = (t) => !needle || String(t).toLowerCase().includes(needle);
        const visShapes = needle ? vShapes.filter(keep) : vShapes;
        const visImages = needle ? vImages.filter(keep) : vImages;
        const visBlips = needle ? blips.filter(keep) : blips;
        lines.push("");
        lines.push("--- VML 图形（officecli 不可见）---");
        if (visShapes.length === 0) lines.push(needle ? "（无匹配）" : "（无）");
        visShapes.slice(0, limit).forEach((t, i) => {
          const id = attr(t, "id") || `#${i + 1}`;
          const style = attr(t, "style");
          const stroke = [attr(t, "strokecolor"), attr(t, "strokeweight"), attr(t, "stroked")].filter(Boolean).join(" / ");
          lines.push(`${id}  style='${clip(style, 160)}'${stroke ? `  stroke={${stroke}}` : ""}`);
        });
        if (visShapes.length > limit) lines.push(`… 还有 ${visShapes.length - limit} 个未列出`);
        lines.push("");
        lines.push("--- VML 内嵌图片关系 ---");
        if (visImages.length === 0) lines.push(needle ? "（无匹配）" : "（无）");
        visImages.slice(0, limit).forEach((t) => {
          lines.push(`r:id=${attr(t, "r:id") || "?"}${attr(t, "o:title") ? `  title='${attr(t, "o:title")}'` : ""}`);
        });
        lines.push("");
        lines.push("--- 新式 DrawingML ---");
        if (anchors.length + inlines === 0) lines.push("（无）");
        anchors.slice(0, limit).forEach((t, i) => {
          const ex = extents[i] || "";
          lines.push(`anchor ${i + 1}  ${emu2cm(attr(ex, "cx"))} x ${emu2cm(attr(ex, "cy"))}  behindDoc=${attr(t, "behindDoc") || "0"}  relativeHeight=${attr(t, "relativeHeight") || ""}`);
        });
        blips.slice(0, limit).filter(keep).forEach((t, i) => {
          lines.push(`blip ${i + 1}  r:embed=${attr(t, "r:embed") || "?"}`);
        });
      }
      return text(lines.join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_add_image",
    description:
      "向 .docx 插入图片（默认浮动锤定）。officecli 的 add picture 本身就能做浮动插入（anchor=true），但它默认的参考系是段落/行，给负偏移或大偏移时图会跑到页眉/页脚上去——所以本工具把 hRelative/vRelative 默认钉成 page，并把长度参数的坑堵上：裸数字在 officecli 里会被当成 EMU（914400 每英寸），本工具一律按厘米处理。可选 anchorText 用文本定位（插到包含该文字的段落之后）。写操作，会修改文档，调用前先向用户确认参数。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: ".docx 文件路径。" },
        image: { type: "string", description: "图片文件路径（png/jpg 等）。" },
        anchorText: { type: "string", description: "插到包含这段文字的段落之后（可用 --after find: 定位）；省略则追加到文档末尾。" },
        parent: { type: "string", description: "父节点路径，默认 /body。" },
        inline: { type: "boolean", description: "true = 内联图（随文字走）；默认 false = 浮动锤定图。" },
        width: { type: "number", description: "宽（厘米）。省略则用图片原始尺寸。" },
        height: { type: "number", description: "高（厘米）。建议与 width 同给，避免被拉伸。" },
        hPosition: { type: "number", description: "水平偏移（厘米），相对 hRelative。" },
        vPosition: { type: "number", description: "垂直偏移（厘米），相对 vRelative。" },
        hRelative: { type: "string", description: "水平参考系：page/margin/column/character，默认 page。" },
        vRelative: { type: "string", description: "垂直参考系：page/margin/paragraph/line，默认 page。" },
        hAlign: { type: "string", description: "水平对齐（left/center/right），给了它就忽略 hPosition。" },
        vAlign: { type: "string", description: "垂直对齐（top/bottom/center），给了它就忽略 vPosition。" },
        wrap: { type: "string", description: "环绕方式：none/square/tight/topandbottom/through。" },
        behindText: { type: "boolean", description: "true = 图衬在文字下方（配合 wrap=none）。" },
        alt: { type: "string", description: "替代文字（无障碍/可检索性）。" },
        caption: { type: "string", description: "把图注烧进图片（可选路线）。适合只用于打印、不需要检索的场景。注意：图注会变成像素，不可搜索/编辑/复制，且字号会跟着图片缩放。烧入时必须同时给 width，否则字号算不准。" },
        captionSizePt: { type: "number", description: "图注在文档里应显示的字号（磅），默认 9。" },
        captionFont: { type: "string", description: "图注字体，默认 SimSun（宋体）。" },
        captionColor: { type: "string", description: "图注颜色，默认 #000000。" },
        captionGap: { type: "number", description: "图与图注之间的间距（像素），默认 12。" },
        captionAlign: { type: "string", description: "图注对齐：center（默认）/ left。" },
        border: { type: "string", description: "描边：如 1pt:#FFC000、2pt #FF0000、#FFC000（只给颜色则默认 1pt）、none（不加）。单位支持 pt/cm/px。默认不加。officecli 本身无边框属性，这里是拿到图后按 paraId 精确定位 spPr 注入 <a:ln>。" },
        healCheck: { type: "boolean", description: "true 时在改完边框后跑一次 validate 校验文档结构，默认 true。" },
        dryRun: { type: "boolean", description: "true 时只回显将执行的 officecli 参数，不真写。" },
      },
      required: ["path", "image"],
    },
    execute: async ({ path: p, image, anchorText, parent, inline, width, height, hPosition, vPosition, hRelative, vRelative, hAlign, vAlign, wrap, behindText, alt, caption, captionSizePt, captionFont, captionColor, captionGap, captionAlign, border, healCheck, dryRun }) => {
      const file = String(p || "").trim();
      const img = String(image || "").trim();
      if (!file) return fail("缺少 path（.docx 文件路径）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx；老式 .doc 请先用 office_convert 转成 .docx。");
      if (!img) return fail("缺少 image（图片路径）。");

      // Optional route: burn the caption into the picture so photo+caption go in as a single element.
      let effectiveImage = img;
      const capNotes = [];
      if (caption) {
        if (width === undefined || width === null || width === "") {
          return fail("烧入图注时请同时给 width（厘米）。图注字号是按「图片在文档里的实际宽度」换算的，不给就没法算准。");
        }
        ensureWorkDir();
        const capFile = path.join(workDir, `caption_${randomUUID()}.txt`);
        const composed = path.join(workDir, `composed_${randomUUID()}.png`);
        try {
          fs.writeFileSync(capFile, String(caption), "utf8");
          const pwsh = await resolvePwsh();
          const cr = await run(pwsh, [
            "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", COMPOSE_CAPTION,
            "-Image", img, "-Out", composed, "-CaptionFile", capFile,
            "-InsertWidthCm", String(width), "-SizePt", String(captionSizePt === undefined ? 9 : captionSizePt),
            "-FontName", String(captionFont || "SimSun"), "-Color", String(captionColor || "#000000"),
            "-GapPx", String(captionGap === undefined ? 12 : captionGap), "-Align", String(captionAlign || "center"),
          ], { timeoutMs: TIMEOUT.run });
          if (!cr.ok) return fail(`图注合成失败（${describe(cr)}）。\n${clip(combine(cr.stdout, cr.stderr) || cr.message, 2000)}`);
          let info = null;
          try { info = JSON.parse(String(cr.stdout).trim().split(/\r?\n/).at(-1)); } catch { /* 下面统一报错 */ }
          if (!info || !info.ok) return fail(`图注合成未返回有效结果：${clip(cr.stdout, 500)}`);
          effectiveImage = String(info.out);
          capNotes.push(`图注已烧进图片：${info.width}x${info.height}，字号 ${info.fontPt}pt${info.mappedPt ? `（按插入宽度 ${width}cm 换算）` : "（未按宽度换算）"}。`);
          capNotes.push("提醒：图注现已是像素，不可搜索/编辑/复制，且会随图片缩放。需要可检索的文字图注请用 office_insert_caption。");
        } finally {
          try { fs.unlinkSync(capFile); } catch {}
        }
      }

      // Bare numbers are EMU to officecli, which silently makes an invisible image. Always send cm.
      const len = (v) => (v === undefined || v === null || v === "" ? "" : `${String(v).trim()}cm`);
      const props = [`src=${effectiveImage}`];
      const notes = [];
      if (!inline) {
        props.push("anchor=true");
        // Defaulting the reference frames to the page is the whole point: officecli's own default is
        // paragraph/line, so a large or negative offset slams the image into the header.
        props.push(`hRelative=${String(hRelative || "page")}`);
        props.push(`vRelative=${String(vRelative || "page")}`);
        notes.push(`参考系：水平=${String(hRelative || "page")}、垂直=${String(vRelative || "page")}`);
      }
      if (width !== undefined && width !== null && width !== "") props.push(`width=${len(width)}`);
      if (height !== undefined && height !== null && height !== "") props.push(`height=${len(height)}`);
      if (!inline && hAlign) props.push(`hAlign=${String(hAlign)}`);
      else if (!inline && hPosition !== undefined && hPosition !== null && hPosition !== "") props.push(`hPosition=${len(hPosition)}`);
      if (!inline && vAlign) props.push(`vAlign=${String(vAlign)}`);
      else if (!inline && vPosition !== undefined && vPosition !== null && vPosition !== "") props.push(`vPosition=${len(vPosition)}`);
      if (!inline && wrap) props.push(`wrap=${String(wrap)}`);
      if (!inline && behindText === true) props.push("behindText=true");
      if (alt) props.push(`alt=${String(alt)}`);

      const argv = ["add", file, String(parent || "/body"), "--type", "picture"];
      for (const kv of props) argv.push("--prop", kv);
      if (anchorText) argv.push("--after", `find:${String(anchorText)}`);

      if (dryRun === true) {
        return text(`将执行（dryRun，未写入）：\nofficecli ${argv.map((a) => (a.includes(" ") ? `"${a}"` : a)).join(" ")}\n\n${notes.join("\n")}`);
      }
      const r = await cli(argv, { timeoutMs: TIMEOUT.run });
      if (r.enoent) return fail("未找到 officecli。先用 office_cli_status 诊断。");
      if (!r.ok) {
        return fail(
          `插入失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 4000)}`,
          "常见原因：图片路径不存在、anchorText 没匹配到段落、或文档正被 Word/WPS 占用。",
        );
      }
      const out = clip(combine(r.stdout, r.stderr) || "(无输出)", 2000);
      const extra = [];

      // Border. officecli's picture has no line/border property at all, so the stroke is injected as
      // <a:ln> into the picture's own pic:spPr, scoped by the paragraph id that add just reported.
      // Appending is schema-legal: <a:ln> follows xfrm/geom in CT_ShapeProperties.
      const bw = (() => {
        const s = String(border || "").trim();
        if (!s || s.toLowerCase() === "none") return null;
        const colorM = s.match(/#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})/);
        const widthM = s.match(/([\d.]+)\s*(pt|cm|px)?/i);
        const factor = { pt: 12700, cm: 360000, px: 9525 };
        const unit = (widthM && widthM[2] ? widthM[2] : "pt").toLowerCase();
        const num = widthM && widthM[1] ? Number(widthM[1]) : 1;
        const emu = Math.max(1, Math.round((Number.isFinite(num) && num > 0 ? num : 1) * (factor[unit] || 12700)));
        return { emu, color: colorM ? colorM[1].slice(0, 6).toUpperCase() : "000000" };
      })();

      if (bw) {
        const pid = (String(r.stdout || "").match(/paraId=([0-9A-Fa-f]+)/) || [])[1];
        if (!pid) {
          extra.push(`边框未加：没能从 add 的返回里取到 paraId（原文：${clip(r.stdout, 200)}）。请用 office_cli_run 走 raw-set 自己加。`);
        } else {
          const frag = `<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="${bw.emu}"><a:solidFill><a:srgbClr val="${bw.color}"/></a:solidFill></a:ln>`;
          const xp = `//w:p[@w14:paraId='${pid}']//pic:spPr`;
          const br = await cli(["raw-set", file, "/document", "--xpath", xp, "--action", "append", "--xml", frag], { timeoutMs: TIMEOUT.run });
          if (br.ok) {
            extra.push(`已加描边 ${(bw.emu / 12700).toFixed(2)}pt #${bw.color}（定位 ${xp}）`);
            if (healCheck !== false) {
              const vr = await cli(["validate", file], { timeoutMs: TIMEOUT.run });
              extra.push(`结构校验：${clip(combine(vr.stdout, vr.stderr) || "(无输出)", 200)}`);
            }
          } else {
            extra.push(`边框注入失败（${describe(br)}）：${clip(combine(br.stdout, br.stderr) || br.message, 400)}`);
          }
        }
      } else {
        extra.push("未加描边（默认无框；需要时传 border=\"1pt:#FFC000\"）。");
      }
      return text([out, capNotes.join("\n"), notes.join("\n"), extra.filter(Boolean).join("\n")].filter(Boolean).join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_render_preview",
    description:
      "把 .docx 渲染成页面预览图（PNG），用于校版式。走 officecli 的 view screenshot，在 Windows 上可用原生 Word 渲染，与打印效果一致。单页用 page，整本缩略图用 grid。生成的 PNG 会直接写盘并返回路径，可用 read 当图看。注意：officecli 没有 docx→pdf 导出器（view pdf 会报 No exporter plugin），要 PDF 请用 office_convert(to=\"pdf\")，那条走 Word COM、同样无声。本工具只读文档，不修改它。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: ".docx 文件路径。" },
        page: { type: "string", description: "页码或区间，如 3 或 1-3 或 1,3,5。省略时默认第 1 页。" },
        grid: { type: "number", description: "整本缩略图拼图：传列数（如 3），或传 0 自动选列。给了它就不看 page。" },
        out: { type: "string", description: "输出 PNG 路径。省略则放在源文件旁的 _preview 子目录。" },
        width: { type: "number", description: "视口宽（像素），默认 1600。" },
        height: { type: "number", description: "视口高（像素），默认 1200。" },
        range: { type: "string", description: "只裁某个区域的包围盒，便于放大看细节。传元素路径（如 /body/table[1]、/body/p[12]）或表格单元格区间（Sheet1!A1:C3）。" },
        render: { type: "string", description: "渲染通道：auto（默认，Windows 上有 Word/PowerPoint 就走原生）、native（强制原生，没有则报错）、html（强制 HTML 通道）。想跟打印一致就用 native。" },
      },
      required: ["path"],
    },
    execute: async ({ path: p, page, grid, out, width, height, range, render }) => {
      const file = String(p || "").trim();
      if (!file) return fail("缺少 path（.docx 文件路径）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx；老式 .doc 请先用 office_convert 转成 .docx。");

      const dir = file.replace(/[\\\/][^\\\/]*$/, "");
      const base = (file.match(/([^\\\/]+)\.docx$/i) || [, "doc"])[1];
      const useGrid = grid !== undefined && grid !== null && grid !== "";
      const tag = useGrid ? `grid${Number(grid) > 0 ? Number(grid) : "auto"}` : `p${String(page || "1").replace(/[,\-]/g, "_")}`;
      const outPath = String(out || "").trim() || `${dir}\\_preview\\${base}_${tag}.png`;

      // officecli does not create missing parent directories (it fails with "Could not find a part
      // of the path"), so materialise the output folder first. The App's PowerShell child can write
      // outside its own dataDir, which is why this goes through pwsh instead of node fs.
      const outDir = outPath.replace(/[\\\/][^\\\/]*$/, "");
      if (outDir) {
        try {
          const pwsh = await resolvePwsh();
          await run(pwsh, ["-NoProfile", "-NonInteractive", "-Command", `New-Item -ItemType Directory -Force -Path '${outDir.replace(/'/g, "''")}' | Out-Null`], { timeoutMs: 20_000 });
        } catch { /* 建不出来就让 officecli 自己报错，错误信息更具体 */ }
      }

      const argv = ["view", file, "screenshot"];
      if (useGrid) argv.push("--grid", String(Number(grid) > 0 ? Number(grid) : "auto"));
      else argv.push("--page", String(page || "1"));
      if (width) argv.push("--screenshot-width", String(width));
      if (height) argv.push("--screenshot-height", String(height));
      if (range) argv.push("--range", String(range));
      if (render) argv.push("--render", String(render));
      argv.push("-o", outPath);

      const r = await cli(argv, { timeoutMs: TIMEOUT.runMax });
      if (r.enoent) return fail("未找到 officecli。先用 office_cli_status 诊断。");
      if (!r.ok) {
        return fail(
          `渲染失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 3000)}`,
          "如果报 No exporter plugin，说明该 officecli 没装渲染插件；改用 office_convert(to=\"pdf\") 出 PDF 再看。",
        );
      }
      const note = [
        `渲染完成：${outPath}`,
        "（可直接用 read 读这个 PNG 校版式；要 PDF 用 office_convert 传 to=\"pdf\"。）",
      ].join("\n");
      return text([clip(r.stdout, 800), note].filter(Boolean).join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_insert_caption",
    description:
      "在 .docx 里插入图注段落（文字，与原件风格一致）。可用 template 克隆文档里已有的某个图注段（格式完全继承），或用 style/font/align/indent/spaceBefore 等参数自己描。left+right 时按 separatorSpaces 个空格拼成一行两条（复刻原件一栏两图的写法）。keepNext/keepLines 默认关：它们会在 Word 左边距留可见黑方块，要交出去的文档请改用隐形表格 + cantSplit 防拆散。写操作，会修改文档，调用前先向用户确认参数。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: ".docx 文件路径。" },
        text: { type: "string", description: "图注文字。与 left/right 二选一。" },
        left: { type: "string", description: "左栏图注（与 right 一起用，中间用空格分隔）。" },
        right: { type: "string", description: "右栏图注。" },
        separatorSpaces: { type: "number", description: "left 与 right 之间的空格数，默认 27（原件写法）。" },
        template: { type: "string", description: "克隆源段落路径（如 /body/p[21]），格式完全继承它。推荐从文档自身已有图注段拷。" },
        after: { type: "string", description: "插到包含这段文字的段落之后。" },
        afterPath: { type: "string", description: "插到这个段落路径之后（如 /body/p[12]），与 after 二选一。" },
        before: { type: "string", description: "插到包含这段文字的段落之前。" },
        style: { type: "string", description: "段落样式 id（如 Normal、caption、题注）。" },
        font: { type: "string", description: "字体名（如宋体、Times New Roman）。" },
        align: { type: "string", description: "对齐：left/center/right/justify。" },
        indent: { type: "string", description: "左缩进（带单位，如 0.5cm、12pt）。" },
        firstLineIndent: { type: "string", description: "首行缩进（带单位）。" },
        spaceBefore: { type: "string", description: "段前距（带单位，如 3pt）。" },
        spaceAfter: { type: "string", description: "段后距（带单位）。" },
        lineSpacing: { type: "string", description: "行距（如 1.5、14pt）。" },
        lineRule: { type: "string", description: "行距规则：auto/exact/atLeast。" },
        charSpacing: { type: "number", description: "字符间距（pt），需要把图注拉宽时用。" },
        keepNext: { type: "boolean", description: "与下段同页（默认 false）。注意：Word 开启「显示编辑标记」时，带分页属性的段落会在左边距显出一竖列黑方块；要交给别人的正式文档不要默认开，防拆散请用无边框单格表格 + cantSplit。" },
        keepLines: { type: "boolean", description: "整段不拆行（默认 false）。同上：会在 Word 左边距留可见标记。" },
        dryRun: { type: "boolean", description: "true 时只回显将执行的 officecli 参数，不真写。" },
      },
      required: ["path"],
    },
    execute: async ({ path: p, text, left, right, separatorSpaces, template, after, afterPath, before, style, font, align, indent, firstLineIndent, spaceBefore, spaceAfter, lineSpacing, lineRule, charSpacing, keepNext, keepLines, dryRun }) => {
      const file = String(p || "").trim();
      if (!file) return fail("缺少 path（.docx 文件路径）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx；老式 .doc 请先用 office_convert 转成 .docx。");
      const sep = Number.isFinite(Number(separatorSpaces)) && Number(separatorSpaces) >= 0 ? Number(separatorSpaces) : 27;
      let body = String(text || "");
      if (!body && (left || right)) body = `${String(left || "")}${" ".repeat(sep)}${String(right || "")}`;
      if (!body) return fail("缺少图注内容：传 text，或同时传 left 与 right。");
      if (!after && !afterPath && !before) {
        return fail("需要指定插入位置：传 after（文本定位）、afterPath（段落路径）或 before 之一。");
      }

      const props = [`text=${body}`];
      // keepNext/keepLines default OFF. They do stop a caption splitting away from its photo, but Word
      // paints a visible black square in the left margin for paragraphs carrying page-break properties
      // (seen with formatting marks on; a user rejected a delivery over exactly this). For a document
      // that goes to someone else, the anti-split mechanism should be an invisible single-cell table
      // with cantSplit, which leaves no mark.
      const kn = keepNext === true;
      const kl = keepLines === true;
      if (kn) props.push("keepNext=true");
      if (kl) props.push("keepLines=true");
      if (style) props.push(`style=${String(style)}`);
      if (font) props.push(`font=${String(font)}`);
      if (align) props.push(`align=${String(align)}`);
      if (indent) props.push(`indent=${String(indent)}`);
      if (firstLineIndent) props.push(`firstLineIndent=${String(firstLineIndent)}`);
      if (spaceBefore) props.push(`spaceBefore=${String(spaceBefore)}`);
      if (spaceAfter) props.push(`spaceAfter=${String(spaceAfter)}`);
      if (lineSpacing) props.push(`lineSpacing=${String(lineSpacing)}`);
      if (lineRule) props.push(`lineRule=${String(lineRule)}`);
      if (charSpacing !== undefined && charSpacing !== null && charSpacing !== "") props.push(`charSpacing=${String(charSpacing)}`);

      const anchorArgs = [];
      if (afterPath) anchorArgs.push("--after", String(afterPath));
      else if (after) anchorArgs.push("--after", `find:${String(after)}`);
      else if (before) anchorArgs.push("--before", `find:${String(before)}`);

      // With a template we clone that paragraph (formatting inherited verbatim) and then overwrite
      // only the text; without one we create a fresh paragraph and set every property explicitly.
      const argv = template
        ? ["add", file, "/body", "--from", String(template), ...anchorArgs]
        : ["add", file, "/body", "--type", "paragraph", ...props.flatMap((kv) => ["--prop", kv]), ...anchorArgs];

      if (dryRun === true) {
        const q = (a) => (a.includes(" ") ? `"${a}"` : a);
        const lines = [`将执行（dryRun，未写入）：`, `officecli ${argv.map(q).join(" ")}`];
        if (template) lines.push(`然后 set 文本：officecli set ${q(file)} <新段落> --prop ${q("text=" + body)}`);
        return text(lines.join("\n"));
      }

      const r = await cli(argv, { timeoutMs: TIMEOUT.run });
      if (r.enoent) return fail("未找到 officecli。先用 office_cli_status 诊断。");
      if (!r.ok) {
        return fail(
          `插入失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 4000)}`,
          "常见原因：template 或 afterPath 不存在、after 的文字没匹配到、文档被 Word/WPS 占用。",
        );
      }
      const lines = [clip(combine(r.stdout, r.stderr) || "(无输出)", 1500)];

      if (template) {
        // Two shapes come back depending on the route: "Added paragraph at /body/p[@paraId=X]" and
        // "Copied to /body/p[N]". Accept either, otherwise the clone silently keeps the template text.
        const outStr = String(r.stdout || "");
        const pid = (outStr.match(/paraId=([0-9A-Fa-f]+)/) || [])[1];
        const pIdx = (outStr.match(/\/body\/p\[(\d+)\]/) || [])[1];
        const target = pid ? `/body/p[@paraId=${pid}]` : pIdx ? `/body/p[${pIdx}]` : "";
        if (!target) {
          lines.push(`已克隆 ${template}，但没从返回里识别出目标段落（原文：${clip(outStr, 200)}），文本未替换。请用 office_cli_run 的 set 手动改。`);
        } else {
          const setArgs = ["set", file, target];
          for (const kv of props) setArgs.push("--prop", kv);
          const sr = await cli(setArgs, { timeoutMs: TIMEOUT.run });
          lines.push(sr.ok ? `已克隆模板并写入文本（目标 ${target}）` : `克隆成功但 set 文本失败：${clip(combine(sr.stdout, sr.stderr) || sr.message, 400)}`);
        }
      } else {
        lines.push("已插入图注段。");
      }
      if (kn || kl) {
        lines.push("注意：已按你的要求加了 keepNext/keepLines。Word 开启「显示编辑标记」时会在左边距显出黑方块；正式交付前建议改用隐形表格 + cantSplit。");
      }
      return text(lines.join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_insert_photo_rows",
    description:
      "按文档自身版式批量插入现场照片行（克隆原语）。针对 A3 横版双欄、含老式 VML 浮动图的工程文档：从目标文档自己的图片段/空行/图注段取模板，只换图片与图注文字，每行套隐形表格 + cantSplit 防跳欄拆散。这是报告里最有效的那条原语——跨文档套模板必然错位，所以几何一律从本文档自己算。rows 为数组，每项 {left:{img,cap}, right:{img,cap}}（左右可缺一个）。写操作，会修改文档（默认另存为 _withphotos.docx），调用前先向用户确认。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: "源 .docx 路径。" },
        rows: { type: "array", description: "照片行数组：[{left:{img,cap},right:{img,cap}}, ...]。img 是图片路径，cap 是图注文字。", items: { type: "object" } },
        out: { type: "string", description: "输出路径。省略则写在源文件旁，名字加 _withphotos 后缀。" },
        anchorText: { type: "string", description: "插到这个文字所在段落之前（如“变更内容”）。省略则插到文末 sectPr 之前。" },
        targetWidth: { type: "number", description: "内嵌照片宽（像素），默认 1040，控制成品体积。" },
        quality: { type: "number", description: "JPEG 质量，默认 82。" },
        borderColor: { type: "string", description: "VML 描边颜色（如 #FFC000）或 none（默认，不加框）。" },
        photoGapPt: { type: "number", description: "同排两图间距（pt）。省略时默认 24。" },
        rowGapExtra: { type: "number", description: "每行后额外空行数，默认 0。" },
        captionGapExtra: { type: "number", description: "图注上方额外空行数，默认 0。" },
        dryRun: { type: "boolean", description: "true 时只回显参数与将执行的命令，不写入。" },
      },
      required: ["path", "rows"],
    },
    execute: async ({ path: p, rows, out, anchorText, targetWidth, quality, borderColor, photoGapPt, rowGapExtra, captionGapExtra, dryRun }) => {
      const file = String(p || "").trim();
      if (!file) return fail("缺少 path（源 .docx）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx。");
      const list = Array.isArray(rows) ? rows : [];
      if (list.length === 0) return fail("rows 为空：至少给一行 {left:{img,cap}} 或 {right:{img,cap}}。");
      const bad = list.findIndex((r) => {
        if (!r || typeof r !== "object") return true;
        const okOne = (n) => n && typeof n === "object" && String(n.img || "").trim();
        return !okOne(r.left) && !okOne(r.right);
      });
      if (bad >= 0) return fail(`rows[${bad}] 无效：每行至少要有 left 或 right，且其 img 不能为空。`);

      ensureWorkDir();
      const specFile = path.join(workDir, `rows_${randomUUID()}.json`);
      fs.writeFileSync(specFile, JSON.stringify(list), "utf8");
      const argv = [
        "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", INSERT_ROWS,
        "-Docx", file, "-SpecJson", specFile,
      ];
      if (out) argv.push("-Out", String(out));
      if (anchorText) argv.push("-AnchorText", String(anchorText));
      if (targetWidth !== undefined && targetWidth !== null && targetWidth !== "") argv.push("-TargetW", String(targetWidth));
      if (quality !== undefined && quality !== null && quality !== "") argv.push("-Quality", String(quality));
      if (borderColor) argv.push("-BorderColor", String(borderColor));
      if (photoGapPt !== undefined && photoGapPt !== null && photoGapPt !== "") argv.push("-PhotoGapPt", String(photoGapPt));
      if (rowGapExtra) argv.push("-RowGapExtra", String(rowGapExtra));
      if (captionGapExtra) argv.push("-CaptionGapExtra", String(captionGapExtra));

      if (dryRun === true) {
        try { fs.unlinkSync(specFile); } catch {}
        const q = (a) => (a.includes(" ") ? `"${a}"` : a);
        return text(`将执行（dryRun，未写入）：\n${argv.map(q).join(" ")}\n\n共 ${list.length} 行；锚点=${anchorText || "（文末）"}。`);
      }

      try {
        const pwsh = await resolvePwsh();
        const r = await run(pwsh, argv, { timeoutMs: TIMEOUT.runMax });
        if (!r.ok) {
          return fail(
            `插入失败（${describe(r)}）。\n${clip(combine(r.stdout, r.stderr) || r.message, 4000)}`,
            "常见原因：图片路径不存在、anchorText 没匹配到、目标 docx 被 Office/WPS 占用、或 rows 里 img 写错。",
          );
        }
        let info = null;
        try { info = JSON.parse(String(r.stdout).trim().split(/\r?\n/).at(-1)); } catch { /* 下面统一报错 */ }
        if (!info || !info.ok) return fail(`脚本没返回有效结果：${clip(r.stdout, 800)}`);
        const extra = [
          "",
          `产出：${info.out}`,
          `照片 ${info.photos} 张 / ${info.rows} 行；模板来源 ${info.template}；行高 ${info.rowHeight}pt；栏宽 ${info.colWidth}pt（图框 ${info.boxW}pt）；图注制表位 ${info.tabs}`,
          "防拆散用的是隐形表格 + cantSplit，不会在 Word 左边距留黑方块。建议再用 office_render_preview 渲染校版。",
        ];
        return text([clip(r.stdout, 400), extra.join("\n")].join("\n"));
      } finally {
        try { fs.unlinkSync(specFile); } catch {}
      }
    },
  });

  await sdk.tools.register({
    name: "office_doc_geometry",
    description:
      "读出一个 .docx 的版式几何量：页面尺寸与边距、分欄数与欄间距、算出欄宽，并给出推荐的图框尺寸与左右位置；同时列出文档里已有浮动图（VML）的 margin-top/left/width/height 与图注段落样本。用途：按原件尺寸对齐插图，避免跨文档套模板导致错位。只读，不修改文件。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: ".docx 文件路径。" },
        gapPt: { type: "number", description: "同排两图间距（pt），默认 24。" },
        cap: { type: "number", description: "图框宽上限（pt），默认 224。" },
        ratio: { type: "number", description: "图框高宽比的分母：高 = 宽 × ratio / 4（默认 3 即 4:3）。" },
      },
      required: ["path"],
    },
    execute: async ({ path: p, gapPt, cap, ratio }) => {
      const file = String(p || "").trim();
      if (!file) return fail("缺少 path（.docx 文件路径）。");
      if (!/\.docx$/i.test(file)) return fail("只处理 .docx。");
      const r = await cli(["raw", file, "/document"], { timeoutMs: TIMEOUT.run });
      if (r.enoent) return fail("未找到 officecli。");
      if (!r.ok) return fail(`读取文档 XML 失败（${describe(r)}）。`);
      const xml = String(r.stdout || "");
      const num = (re, s) => {
        const m = (s || xml).match(re);
        return m ? Number(m[1]) : 0;
      };
      const sec = (xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/) || [""])[0];
      const tw = (v) => (v ? (v / 20).toFixed(2) : "?");   // twips -> pt
      const pgW = num(/<w:pgSz[^>]*w:w="(\d+)"/, sec);
      const pgH = num(/<w:pgSz[^>]*w:h="(\d+)"/, sec);
      const mL = num(/<w:pgMar[^>]*w:left="(\d+)"/, sec);
      const mR = num(/<w:pgMar[^>]*w:right="(\d+)"/, sec);
      const mT = num(/<w:pgMar[^>]*w:top="(\d+)"/, sec);
      const mB = num(/<w:pgMar[^>]*w:bottom="(\d+)"/, sec);
      const cols = num(/<w:cols[^>]*w:num="(\d+)"/, sec) || 1;
      const csp = num(/<w:cols[^>]*w:space="(\d+)"/, sec);
      const cw = pgW > 0 ? (pgW - mL - mR - csp * (cols - 1)) / cols / 20 : 0;

      const gap = Number.isFinite(Number(gapPt)) && Number(gapPt) > 0 ? Number(gapPt) : 24;
      const capW = Number.isFinite(Number(cap)) && Number(cap) > 0 ? Number(cap) : 224;
      const rDen = Number.isFinite(Number(ratio)) && Number(ratio) > 0 ? Number(ratio) : 3;
      let boxW = capW;
      if (cw > 60) {
        boxW = Math.min(capW, (cw - gap) / 2 - 2);
        if (boxW < 60) boxW = 60;
      }
      const boxH = (boxW * rDen) / 4;
      const pairW = 2 * boxW + gap;
      const leftMl = cw > 0 ? Math.max(0, (cw - pairW) / 2) : 0;

      // Existing floating VML shapes, so the caller can align with what is already there.
      const shapes = (xml.match(/<v:shape\b[^>]*>/gi) || []).map((t) => {
        const st = (t.match(/style="([^"]*)"/) || [, ""])[1];
        const g = (k) => {
          const m = st.match(new RegExp(k + ":([-0-9.]+)"));
          return m ? m[1] : "";
        };
        return { id: (t.match(/(?<![:\w])id="([^"]*)"/) || [, "?"])[1], ml: g("margin-left"), mt: g("margin-top"), w: g("width"), h: g("height") };
      });
      const vImages = (xml.match(/<v:imagedata\b/gi) || []).length;

      // A caption paragraph sample: the first paragraph whose text carries a stake-number pattern.
      let capSample = "";
      for (const pm of xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []) {
        const t = (pm.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map((s) => (s.match(/>([^<]*)</) || [, ""])[1]).join("");
        if (/\d+\s*\+\s*\d+/.test(t) && t.trim().length <= 80) { capSample = t.trim(); break; }
      }

      const lines = [
        file,
        `页面 ${tw(pgW)} x ${tw(pgH)} pt；边距 左/右/上/下 = ${tw(mL)}/${tw(mR)}/${tw(mT)}/${tw(mB)} pt`,
        `分欄 ${cols} 欄，欄间距 ${tw(csp)} pt → **欄宽 ${cw ? cw.toFixed(2) : "?"} pt**`,
        "",
        `建议图框：${boxW.toFixed(2)} x ${boxH.toFixed(2)} pt（上限 ${capW}，间距 ${gap}）`,
        `一对两图总宽 ${pairW.toFixed(2)} pt → 左图 margin-left ${leftMl.toFixed(2)}、右图 ${(leftMl + boxW + gap).toFixed(2)}`,
        "",
        `文档已有浮动图（含真图 ${vImages} 张，v:shape ${shapes.length} 个）：`,
      ];
      if (shapes.length === 0) lines.push("  （无）");
      shapes.slice(0, 20).forEach((s) => lines.push(`  ${s.id}  ml=${s.ml || "?"} mt=${s.mt || "?"} ${s.w || "?"} x ${s.h || "?"}`));
      if (shapes.length > 20) lines.push(`  … 还有 ${shapes.length - 20} 个`);
      lines.push("");
      lines.push(capSample ? `图注段样本：${clip(capSample, 100)}` : "图注段样本：（没找到带桩号的段落）");
      lines.push("提示：几何一律以本文档为准；跨文档套模板必然错位。批量插图请用 office_insert_photo_rows。");
      return text(lines.join("\n"));
    },
  });

  await sdk.tools.register({
    name: "office_convert_status",
    description: "查询 Office 旧文档转换任务状态。传 jobId 查询单个任务；不传则列出最近 10 个任务。可看到阶段、百分比、成功数、跳过数（目标已存在时不会覆盖，计为跳过而非失败）、失败数与错误信息。",
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

    app.post("/cleanup", async (c) => {
      // Panel button: same routine as the office_cleanup tool. Clicking the button is the explicit
      // confirmation, so no extra confirm flag is needed on this path.
      const r = await runCleanup();
      return c.json({ ok: true, ...r });
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

  // Best-effort sweep on load so a previous session's leftovers do not pile up.
  void sweepTemp();

  await sdk.logger.info(`office-toolkit ${VERSION} ready: tools=office_deps_status/office_deps_install/office_cli_status/office_cli_run/office_convert/office_convert_status/office_convert_cancel, routes=status/deps/run/convert/jobs`);
});

// ---------------------------------------------------------------- 工具返回封装

function text(s) {
  return { content: [{ type: "text", text: String(s) }] };
}
function fail(message, hint) {
  return { content: [{ type: "text", text: hint ? `${message}\n\n提示：${hint}` : String(message) }], isError: true };
}


