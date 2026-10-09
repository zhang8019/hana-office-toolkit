// doc-mcp.js -- minimal MCP stdio client so the App can host the document MCP itself.
//
// Why this exists: an App-provided MCP connector (owner.kind === "app") is registered by the host but
// its tools are never projected into an agent's tool namespace, so all 74 tools of
// timeverse-office-doc are unreachable that way. The App's OWN tools do get exposed, so we speak MCP
// to the server ourselves and hand the results out through our own tools.
//
// Transport: newline-delimited JSON-RPC 2.0 over the child's stdin/stdout (that is what the python
// mcp SDK uses over stdio), NOT Content-Length framing.

import { spawn } from "node:child_process";

const PROTOCOL_VERSION = "2024-11-05";
const DEFAULT_TIMEOUT = 60_000;
const START_TIMEOUT = 90_000;   // uvx may need to resolve/build the package on a cold start

export class DocMcpClient {
  /**
   * @param {{ resolveExecutable?: (o: {candidates: string[]}) => Promise<{path?: string}>, logger?: { info: Function, warn: Function } }} opts
   */
  constructor(opts = {}) {
    this.resolveExecutable = opts.resolveExecutable;
    this.logger = opts.logger;
    this.child = null;
    this.starting = null;
    this.dead = false;
    this.seq = 0;
    this.pending = new Map();
    this.buf = "";
    this.stderr = "";
    this.tools = null;
    this.serverInfo = null;
  }

  async _resolveUvx() {
    if (this.resolveExecutable) {
      for (const c of ["uvx", "uvx.exe"]) {
        try {
          const info = await this.resolveExecutable({ candidates: [c] });
          if (info && info.path) return info.path;
        } catch { /* try next */ }
      }
    }
    return "uvx";
  }

  async _spawn() {
    const exe = await this._resolveUvx();
    // The server sandboxes every file operation to OFFICE_ALLOWED_DIRS (prefix match; a relative path is
    // anchored to OFFICE_BASE_DIR). Without this the first call fails with "不在允许的目录范围内".
    // Default: the user profile plus the non-system drive roots people actually keep documents on;
    // OFFICE_ALLOWED_DIRS in the environment overrides it.
    const allowed =
      String(process.env.OFFICE_ALLOWED_DIRS || "").trim() ||
      [process.env.USERPROFILE, "C:\\", "D:\\", "E:\\", "F:\\", "G:\\", "H:\\"]
        .filter(Boolean)
        .join(",");
    const env = {
      ...process.env,
      OFFICE_ALLOWED_DIRS: allowed,
      OFFICE_BASE_DIR: String(process.env.OFFICE_BASE_DIR || process.env.USERPROFILE || "."),
    };
    const child = spawn(exe, ["timeverse-office-doc-mcp"], {
      windowsHide: true,
      stdio: ["pipe", "pipe", "pipe"],
      env,
    });
    // Never let an EventEmitter 'error' event take the whole app down.
    child.on("error", () => {});
    child.stdin.on("error", () => {});
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");

    child.stdout.on("data", (chunk) => this._onData(chunk));
    child.stderr.on("data", (chunk) => {
      this.stderr = (this.stderr + String(chunk)).slice(-8000);
    });
    child.on("exit", () => {
      this.dead = true;
      this.child = null;
      const err = new Error("document MCP server exited");
      for (const [, p] of this.pending) p.reject(err);
      this.pending.clear();
    });
    this.child = child;
    return child;
  }

  _onData(chunk) {
    this.buf += chunk;
    let nl;
    while ((nl = this.buf.indexOf("\n")) >= 0) {
      const line = this.buf.slice(0, nl).trim();
      this.buf = this.buf.slice(nl + 1);
      if (!line) continue;
      let msg;
      try { msg = JSON.parse(line); } catch { continue; }   // ignore non-JSON noise
      if (msg.id === undefined) continue;                    // notification
      const p = this.pending.get(msg.id);
      if (!p) continue;
      this.pending.delete(msg.id);
      clearTimeout(p.timer);
      if (msg.error) p.reject(new Error(String(msg.error.message || msg.error.code || "MCP error")));
      else p.resolve(msg.result);
    }
  }

  _write(obj) {
    if (!this.child || this.dead) throw new Error("document MCP server is not running");
    this.child.stdin.write(JSON.stringify(obj) + "\n");
  }

  /** Send a request and await its result. */
  _request(method, params, timeoutMs = DEFAULT_TIMEOUT) {
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`MCP request "${method}" timed out after ${timeoutMs}ms`));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      try {
        this._write({ jsonrpc: "2.0", id, method, params });
      } catch (e) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(e);
      }
    });
  }

  /** Start (or reuse) the server and handshake. Safe to call repeatedly. */
  async ensureStarted() {
    if (this.child && !this.dead && this.tools) return true;
    if (this.starting) return this.starting;
    this.starting = (async () => {
      this.dead = false;
      this.buf = "";
      this.stderr = "";
      await this._spawn();
      const init = await this._request(
        "initialize",
        {
          protocolVersion: PROTOCOL_VERSION,
          capabilities: {},
          clientInfo: { name: "office-toolkit", version: "1.0" },
        },
        START_TIMEOUT,
      );
      this.serverInfo = init && init.serverInfo ? init.serverInfo : null;
      // The initialized notification has no id and expects no reply.
      this._write({ jsonrpc: "2.0", method: "notifications/initialized", params: {} });
      const listed = await this._request("tools/list", {}, START_TIMEOUT);
      this.tools = Array.isArray(listed && listed.tools) ? listed.tools : [];
      if (this.logger) this.logger.info(`office-toolkit: document MCP ready (${this.tools.length} tools)`);
      return true;
    })();
    try {
      return await this.starting;
    } catch (e) {
      this.dead = true;
      if (this.child) { try { this.child.kill(); } catch {} this.child = null; }
      throw e;
    } finally {
      this.starting = null;
    }
  }

  async listTools() {
    await this.ensureStarted();
    return this.tools || [];
  }

  /** Call one MCP tool. Returns a normalised { ok, text, raw } shape. */
  async call(name, args, timeoutMs = DEFAULT_TIMEOUT) {
    await this.ensureStarted();
    const res = await this._request("tools/call", { name: String(name), arguments: args && typeof args === "object" ? args : {} }, timeoutMs);
    const content = Array.isArray(res && res.content) ? res.content : [];
    const text = content
      .filter((c) => c && c.type === "text")
      .map((c) => String(c.text || ""))
      .join("\n");
    const other = content.filter((c) => c && c.type !== "text").map((c) => c.type);
    return {
      ok: !(res && res.isError),
      text,
      other,
      structured: res && res.structuredContent ? res.structuredContent : null,
    };
  }

  async stop() {
    const c = this.child;
    this.child = null;
    this.dead = true;
    this.tools = null;
    if (!c) return;
    try { c.stdin.end(); } catch {}
    try { c.kill(); } catch {}
  }

  status() {
    return {
      running: !!(this.child && !this.dead),
      dead: this.dead,
      toolCount: this.tools ? this.tools.length : 0,
      server: this.serverInfo,
      stderrTail: this.stderr ? this.stderr.slice(-500) : "",
    };
  }
}
