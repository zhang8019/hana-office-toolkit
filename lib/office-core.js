// office-core.js — 纯逻辑与平台适配层，只依赖 Node，不碰 App 运行时。
// index.js 负责与宿主 SDK 打交道；这里承担：可执行定位、目录遍历、输出裁剪、版本解析。

import fs from "node:fs";
import path from "node:path";

export const OFFICECLI_ENV = "OFFICECLI_PATH";

/** officecli 可执行候选路径（按优先级）。 */
export function officecliCandidates(env = {}, platform = process.platform) {
  const out = [];
  if (env[OFFICECLI_ENV]) out.push(env[OFFICECLI_ENV]);
  if (platform === "win32") {
    if (env.LOCALAPPDATA) out.push(`${env.LOCALAPPDATA}\\OfficeCli\\officecli.exe`);
    if (env.ProgramFiles) out.push(`${env.ProgramFiles}\\OfficeCli\\officecli.exe`);
  } else {
    if (env.HOME) out.push(`${env.HOME}/.local/bin/officecli`);
    out.push("/usr/local/bin/officecli", "/opt/homebrew/bin/officecli");
  }
  out.push("officecli");
  return [...new Set(out.filter(Boolean))];
}

/** 旧格式 → 新格式 的映射；用作转换判据。 */
export const LEGACY_KINDS = {
  doc: { ext: ["doc", "rtf"], to: "docx" },
  xls: { ext: ["xls"], to: "xlsx" },
  ppt: { ext: ["ppt"], to: "pptx" },
};

/** 由扩展名判断旧格式类别；返回 "doc"|"xls"|"ppt"|null。 */
export function legacyKindOf(file) {
  const ext = path.extname(String(file)).slice(1).toLowerCase();
  for (const [kind, spec] of Object.entries(LEGACY_KINDS)) {
    if (spec.ext.includes(ext)) return kind;
  }
  return null;
}

/** 递归收集目录里的旧格式文件（跳过隐藏目录与已转换产物）。 */
export function collectLegacy(root, { recursive = true } = {}) {
  const found = [];
  const walk = (dir) => {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      if (e.name.startsWith(".")) continue;
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (recursive) walk(full);
      } else if (e.isFile()) {
        const kind = legacyKindOf(full);
        if (kind) found.push({ src: full, type: kind });
      }
    }
  };
  walk(root);
  return found.sort((a, b) => a.src.localeCompare(b.src));
}

/** 输出裁剪：超长时保留头部，附一句可读的截断说明。 */
export function clip(text, max = 60_000) {
  const value = String(text ?? "");
  if (value.length <= max) return value;
  return `${value.slice(0, max)}\n\n[输出已截断：共 ${value.length} 字符，仅保留前 ${max} 字符。用 --json / 收窄路径再取。]`;
}

/** 从 `officecli 1.0.155` 之类里取语义化版本；失败返回 null。 */
export function parseVersion(output) {
  const m = String(output ?? "").match(/(\d+)\.(\d+)\.(\d+)/);
  return m ? `${m[1]}.${m[2]}.${m[3]}` : null;
}

/** 参数里若含 shell 元字符则拒绝（我们不经 shell，防止误用注入式写法）。 */
export function hasShellMetachars(argv) {
  return argv.some((a) => /[<>|;&$`]/.test(String(a)));
}

/** 拼接子进程输出：非空段落之间用换行连接。 */
export function combine(...parts) {
  return parts.filter((p) => p && String(p).trim()).join("\n");
}
