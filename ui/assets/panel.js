// Office 工具箱面板：状态 / 异步旧文档转换 / officecli 命令。
import { hana } from "./sdk.js";

const ROUTE = { status: "/status", run: "/run", convert: "/convert", jobs: "/jobs", deps: "/deps", cleanup: "/cleanup" };
const ACTIVE = new Set(["queued", "running", "cancelling"]);
const ui = {
  root: document.getElementById("panel"), startup: document.getElementById("startup"),
  dot: document.getElementById("health-dot"), appVersion: document.getElementById("app-version"),
  cliHint: document.getElementById("cli-hint"), secStatus: document.getElementById("sec-status"),
  secConvert: document.getElementById("sec-convert"), secRun: document.getElementById("sec-run"),
  cvInput: document.getElementById("cv-input"), cvOutput: document.getElementById("cv-output"),
  cvMode: document.getElementById("cv-mode"), cvRecursive: document.getElementById("cv-recursive"),
  cvRun: document.getElementById("cv-run"), cvResult: document.getElementById("cv-result"),
  cvJobs: document.getElementById("cv-jobs"), cvConfirm: document.getElementById("cv-confirm"),
  depsList: document.getElementById("deps-list"), depsRefresh: document.getElementById("deps-refresh"),
  depsCleanup: document.getElementById("deps-cleanup"), depsNote: document.getElementById("deps-note"),
  secDeps: document.getElementById("sec-deps"),
  cvConfirmNo: document.getElementById("cv-confirm-no"), cvConfirmYes: document.getElementById("cv-confirm-yes"),
  runArgs: document.getElementById("run-args"),
  runGo: document.getElementById("run-go"), runResult: document.getElementById("run-result"),
  refresh: document.getElementById("refresh"), updated: document.getElementById("updated"),
};

async function api(path, init) {
  const res = await hana.api.fetch(path, init);
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch { data = { raw }; }
  return { status: res.status, data };
}
const post = (path, body) => api(path, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body || {}) });
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
function setResult(el, value, kind) {
  el.hidden = false; el.textContent = value; el.classList.remove("result--ok", "result--err");
  if (kind) el.classList.add(kind === "ok" ? "result--ok" : "result--err");
}
function stamp() { ui.updated.textContent = `更新于 ${new Date().toLocaleTimeString("zh-CN", { hour12: false })}`; }

async function refreshDeps() {
  try {
    const { status, data } = await api(ROUTE.deps);
    if (status !== 200 || !data?.ok) throw new Error(data?.error || `HTTP ${status}`);
    const d = data.deps || {};
    const rows = [
      ["MS Office / WPS", !!(d.officeCom?.ms || d.officeCom?.wps), [d.officeCom?.ms ? "MS Office" : "", d.officeCom?.wps ? "WPS" : ""].filter(Boolean).join(" + ") || "未检测到（首选引擎：旧格式→现代格式、导出 PDF）"],
      ["LibreOffice", !!d.libreoffice?.found, d.libreoffice?.found ? `${d.libreoffice.version || "版本未知"} · ${d.libreoffice.path}` : "未安装（回落引擎：html/csv/txt/rtf/odt/ods/epub 等只有它能做）"],
      ["officecli", !!d.officecli?.found, d.officecli?.found ? `${d.officecli.version || "版本未知"} · ${d.officecli.path}` : "未安装（docx/xlsx/pptx 读写）"],
      ["uv / uvx", !!(d.uv?.found || d.uvx?.found), (d.uvx?.path || d.uv?.path) || "未安装（文档 MCP）"],
    ];
    ui.depsList.replaceChildren(...rows.map(([name, ok, detail]) => {
      const row = document.createElement("div");
      row.className = "dep" + (ok ? "" : " dep--miss");
      row.innerHTML = `<span class="dep__mark">${ok ? "✓" : "✗"}</span><span class="dep__name">${esc(name)}</span><span class="dep__detail">${esc(detail)}</span>`;
      return row;
    }));
    const miss = (data.missing || []);
    if (miss.length) {
      const p = document.createElement("p");
      p.className = "deps__hint";
      p.textContent = `缺：${miss.join("、")}。可以让我用 office_deps_install 装，或按 README 的链接手动装。`;
      ui.depsList.append(p);
    }
    ui.secDeps.hidden = false;
    stamp();
  } catch (err) {
    ui.secDeps.hidden = false;
    ui.depsList.textContent = `依赖检查失败：${err?.message || err}`;
  }
}

async function refreshStatus() {
  try {
    const { status, data } = await api(ROUTE.status);
    if (status !== 200 || !data?.ok) throw new Error(data?.error || `HTTP ${status}`);
    ui.appVersion.textContent = `v${data.app?.version ?? "—"}`;
    const cli = data.officecli || {};
    ui.dot.classList.toggle("dot--on", !!cli.installed);
    ui.cliHint.textContent = cli.installed ? `${cli.version || "（版本未知）"} · ${cli.path}` : `不可用：${cli.error || cli.path || "未找到"}`;
    ui.secStatus.hidden = ui.secConvert.hidden = ui.secRun.hidden = false;
    ui.startup.hidden = true; stamp(); return true;
  } catch (err) {
    ui.startup.hidden = false; ui.startup.textContent = `读取状态失败：${err?.message || err}`; return false;
  }
}

function renderJobs(jobs) {
  ui.cvJobs.replaceChildren();
  if (!jobs?.length) { const p = document.createElement("p"); p.className = "jobs__empty"; p.textContent = "暂无转换任务"; ui.cvJobs.append(p); return; }
  for (const job of jobs) {
    const card = document.createElement("article"); card.className = "job";
    const pct = Number.isFinite(job.percent) ? Math.max(0, Math.min(100, job.percent)) : 0;
    const statusName = { queued: "排队中", running: "运行中", cancelling: "正在取消", completed: "已完成", failed: "失败", cancelled: "已取消", interrupted: "中断，结果待检查" }[job.status] || job.status;
    const detail = job.error || job.currentFile || job.input || "";
    const summary = job.status === "completed" ? `成功 ${job.okCount || 0} · 失败 ${job.failCount || 0}` : "";
    card.innerHTML = `<div class="job__head"><strong>${esc(statusName)}</strong><span>${esc(job.phase || "")}</span></div>
      <div class="progress" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${pct}"><span style="width:${pct}%"></span></div>
      <div class="job__meta">${job.total ? `${esc(job.current || 0)} / ${esc(job.total)} · ${pct}%` : esc(job.status === "running" ? "处理中，正在识别文件或等待 Office" : "—")} ${esc(summary)}</div>
      <div class="job__detail" title="${esc(detail)}">${esc(detail)}</div>
      <div class="job__foot"><code>${esc(job.jobId)}</code>${ACTIVE.has(job.status) ? `<button class="btn btn--danger btn--compact" data-cancel="${esc(job.jobId)}" ${job.status === "cancelling" ? "disabled" : ""}>取消</button>` : ""}</div>`;
    if (job.status === "failed" || job.status === "interrupted") card.classList.add("job--error");
    ui.cvJobs.append(card);
  }
  ui.cvJobs.querySelectorAll("[data-cancel]").forEach(button => button.addEventListener("click", () => cancelJob(button.dataset.cancel)));
}

async function refreshJobs() {
  try {
    const { status, data } = await api(ROUTE.jobs);
    if (status !== 200 || !data?.ok) throw new Error(data?.error || `HTTP ${status}`);
    renderJobs(data.jobs || []);
    const hasRunning = (data.jobs || []).some(j => ACTIVE.has(j.status));
    ui.cvRun.disabled = hasRunning;
    ui.root.setAttribute("aria-busy", hasRunning ? "true" : "false");
    stamp();
  } catch (err) {
    ui.cvJobs.textContent = `读取任务状态失败：${err?.message || err}`;
  }
}

async function doConvert() {
  if (ui.cvMode.value === "replace") { ui.cvConfirm.hidden = false; return; }
  await submitConversion();
}

async function submitConversion(confirmedReplace = false) {
  const input = ui.cvInput.value.trim();
  if (!input) { setResult(ui.cvResult, "请先填写输入路径。", "err"); return; }
  ui.cvConfirm.hidden = true;
  ui.cvRun.disabled = true;
  setResult(ui.cvResult, "正在提交转换任务…", null);
  try {
    const { status, data } = await post(ROUTE.convert, { input, output: ui.cvOutput.value.trim() || undefined, mode: ui.cvMode.value, recursive: ui.cvRecursive.checked, confirmReplace: confirmedReplace });
    if (status !== 200 || !data?.ok) { setResult(ui.cvResult, `提交失败：${data?.error || `HTTP ${status}`}`, "err"); ui.cvRun.disabled = false; return; }
    setResult(ui.cvResult, `任务已提交：${data.jobId}`, "ok");
    await refreshJobs();
  } catch (err) {
    setResult(ui.cvResult, `提交失败：${err?.message || err}`, "err"); ui.cvRun.disabled = false;
  }
}

async function cancelJob(jobId) {
  try {
    const { status, data } = await post(`${ROUTE.jobs}/${encodeURIComponent(jobId)}/cancel`, {});
    if (status !== 200 || !data?.ok) throw new Error(data?.error || `HTTP ${status}`);
    await refreshJobs();
  } catch (err) { setResult(ui.cvResult, `取消失败：${err?.message || err}`, "err"); }
}

async function doRun() {
  const argv = ui.runArgs.value.trim().split(/\s+/).filter(Boolean);
  if (!argv.length) { setResult(ui.runResult, "请填写命令参数。", "err"); return; }
  ui.runGo.disabled = true; setResult(ui.runResult, "执行中…", null);
  try {
    const { status, data } = await post(ROUTE.run, { args: argv });
    const out = [data?.stdout, data?.stderr && `[stderr]\n${data.stderr}`].filter(Boolean).join("\n") || "（无输出）";
    setResult(ui.runResult, status !== 200 || !data?.ok ? `失败（退出码 ${data?.code ?? "?"}）：\n${data?.error || out}` : out, status !== 200 || !data?.ok ? "err" : "ok"); stamp();
  } catch (err) { setResult(ui.runResult, `失败：${err?.message || err}`, "err"); }
  finally { ui.runGo.disabled = false; }
}

async function doCleanup() {
  ui.depsCleanup.disabled = true; setResult(ui.depsNote, "清理中…", null);
  try {
    const { status, data } = await post(ROUTE.cleanup, {});
    if (status !== 200 || !data?.ok) throw new Error(data?.error || `HTTP ${status}`);
    setResult(ui.depsNote, `已清理：临时目录${data.swept ? "已扫" : "跳过"}；子进程 ${data.killed} 个；无窗口残留 Office/soffice ${data.zombies} 个。`, "ok");
    stamp();
  } catch (err) { setResult(ui.depsNote, `清理失败：${err?.message || err}`, "err"); }
  finally { ui.depsCleanup.disabled = false; }
}

ui.refresh.addEventListener("click", async () => { await refreshStatus(); await refreshDeps(); await refreshJobs(); });
ui.depsRefresh.addEventListener("click", refreshDeps);
ui.depsCleanup.addEventListener("click", doCleanup);
ui.cvRun.addEventListener("click", doConvert);
ui.cvConfirmNo.addEventListener("click", () => { ui.cvConfirm.hidden = true; });
ui.cvConfirmYes.addEventListener("click", () => submitConversion(true));
ui.runGo.addEventListener("click", doRun);
hana.ready();
(async () => { await refreshStatus(); await refreshDeps(); await refreshJobs(); setInterval(refreshJobs, 1500); })();
