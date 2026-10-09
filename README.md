# Office 工具箱（office-toolkit）

一个 Hana v2 App，把本机 **officecli** 桥接为 Hana 工具，并合入**旧格式文档批量转换**。目标是让 docx/xlsx/pptx 的读写与 .doc/.xls/.ppt 的现代化转换都从一个入口走。

## 能力（manifest.capabilities）

| 能力 | 用途 |
| --- | --- |
| `app/tools.expose-to-model` | 向模型注册工具 |
| `app/process.spawn` | 启动 officecli 与转换用 PowerShell 子进程（不授此权，App 内 `execFile` 会 `ERR_ACCESS_DENIED`） |

## 工具

- **office_cli_status** — officecli 可执行路径、版本、可用性。诊断入口。
- **office_cli_run** — 透传一条 officecli 子命令（参数数组，不经 shell）。覆盖 docx/xlsx/pptx 的 `get/set/add/remove/query/view/raw-set/dump/batch/create/merge/import` 等。
- **office_convert** — 提交转换，立即返回 jobId。两种用法：
  - 不传 `to`：旧格式默认映射（.doc/.rtf→.docx、.xls→.xlsx、.ppt→.pptx）
  - 传 `to`：把输入内**所有可识别文档**转成指定格式。支持 `pdf`、`docx`、`doc`、`xlsx`、`xls`、`pptx`、`ppt`、`html`、`txt`、`csv`、`rtf`、`odt`、`ods`、`odp`、`epub`、`md`
- **office_convert_status** — 查询任务进度、结果、失败原因；不传 jobId 列最近任务。
- **office_convert_cancel** — 请求取消并结束该任务的子进程树；已经写出的文件不会自动删除。

## 转换能力矩阵（实测）

| 转换 | 结果 |
| --- | --- |
| .doc → .docx（默认） | ✓ 30/30，约 1 分钟 |
| .doc → .pdf | ✓ |
| .xls → .xlsx（默认） | ✓ |
| .ppt → .pptx（默认） | ✓ |
| .docx → .pdf | ✓ |
| .csv → .xlsx | ✓ |
| 混合目录（doc/xls/ppt/csv）→ pdf | ✓ 4/4，按家族自动选 writer/calc/impress PDF 过滤器 |

**引擎能力边界**：MS Office/WPS 覆盖「旧格式→现代格式」与「导出 PDF」；其余目标格式（html/csv/txt/rtf/odt/ods/odp/epub/md）只有 LibreOffice 能做。

旧格式现代化与导出 PDF 走 Word/WPS；HTML/CSV/TXT 等由 LibreOffice 提供。**officecli 本身不做转换**（它只管 docx/xlsx/pptx 的增删改查），而且**它看不见老式 VML 浮动图**（只认新版 DrawingML）。

**已知边界**：
- 递归目录里同名文件（如 `demo.ppt` 与 `demo.pptx`）转成同一目标格式会撞名，后者会被跳过。
- 目标文件已存在时跳过，不静默覆盖。
- PDF 作为输入（PDF→Office）只有 LibreOffice Draw 的失真导入，本版不做。
- **officecli 看不见老式 VML 浮动图**（`w:pict`/`v:shape`）：`view outline` 会报 0 images、`query picture` 返回空。这类图要用 `raw`/`raw-set` 手工处理，或交给 Word/WPS 转换；officecli 生成的图是新版 DrawingML，也带不上原件那种边框属性。

## 面板

卡片 route `/panel.html`：显示 officecli 状态、提供旧文档转换表单、异步任务进度/取消以及 officecli 命令执行框。数据面 `hana.api.fetch("/api/apps/office-toolkit/routes/...")`，路由：`/status`、`/convert`、`/run`、`/jobs`、`/jobs/:jobId`、`/jobs/:jobId/cancel`。

## 依赖与环境

- **officecli**：默认探测 `%LOCALAPPDATA%\OfficeCli\officecli.exe`，可用环境变量 `OFFICECLI_PATH` 覆盖。
- **转换引擎（自动顺序：MS Office → WPS → LibreOffice）**：
  1. **MS Office COM**（`Word/Excel/PowerPoint.Application`）：旧格式→现代格式、导出 PDF。**优先**，因为它能保住原件版式。
  2. **WPS COM**（`KWPS/KET/KWPP.Application`）：同上，第二顺位。
  3. **LibreOffice headless**：其余目标格式（html/csv/txt/rtf/odt/ods/odp/epub/md）只有它能做；也是 COM 失败时的回落。
  4. 可用 `engine` 参数强制：`auto`（默认）/`office`/`wps`/`libreoffice`。任务结果里会报出**实际使用的引擎**。

  **为什么 COM 要走特殊通道**：Office/WPS 的自动化会拒绝“已提升”的调用方（报 `0x800702E4`，字面像“需要提升”，实际是相反意思）；WPS 的 per-user 注册还会被限制性执行环境隔离（报 `0x80040154`）。所以 COM 一律由 `explorer.exe` 派生的**非提升、脱离隔离**子进程执行，进度与结果经文件回传；失败自动回落 LibreOffice。**不要去改 DCOM / 组策略 / 注册表**——那是错方向。
- **PowerShell**：优先 `pwsh`，回退 `powershell.exe`。脚本纯 ASCII，兼容 5.1；子进程环境精简时脚本会补齐 `PATHEXT`。

## 工具一览

| 工具 | 作用 | 读写 |
| --- | --- | --- |
| `office_deps_status` | 依赖自检（引擎优先级 + 各组件路径/版本） | 读 |
| `office_deps_install` | 安装缺失依赖（libreoffice / uv / officecli） | 写 |
| `office_cli_status` / `office_cli_run` | officecli 桥接：状态查询与命令执行 | 读/写 |
| `office_convert` + `_status` + `_cancel` | 旧文档批量转换（MS Office → WPS → LibreOffice，逐文件回落） | 写 |
| `office_images` | 列出文档内图片，**包含 officecli 看不见的老式 VML 浮动图** | 读 |
| `office_doc_geometry` | 读出页面/边距/分欄并算出**欄宽与推荐图框尺寸**，列出已有浮动图几何 | 读 |
| `office_add_image` | 插图（默认浮动锚定，参考系默认 page；可选描边，可选把图注烧进图片） | 写 |
| `office_insert_caption` | 插入文字图注（可克隆文档自身图注段；一欄两条按空格分隔） | 写 |
| `office_insert_photo_rows` | **按文档自身版式批量插照片行**（克隆原语；隐形表格 + cantSplit 防拆散） | 写 |
| `office_render_preview` | 渲染页面预览图（走 officecli `view screenshot`，原生 Word 渲染） | 读 |
| `office_doc_tools` | 列出随包托管的文档 MCP 工具清单（74 项：Word 19 / Excel 16 / PPT 15 / PDF 16 / 模板会话 8） | 读 |
| `office_doc_tool` | 调用上述 MCP 的一项工具（传 `name` + `arguments`） | 读/写 |
| `office_cleanup` | 清理临时目录与本 App 自己遗留的子进程（**不猜测、不误杀**） | 写 |

### 文档 MCP 为什么由 App 自己托管

宿主会为一个 App 声明的 MCP 连接器启动进程并授权工具，但**不会把工具投影进 agent 的工具命名空间**（同步安装记录时 `owner.kind === "app"` 被跳过）。App 自己的工具则正常暴露，所以 App 直接与 `uvx timeverse-office-doc-mcp` 讲 MCP 协议，再通过 `office_doc_tool` 转出来。因此本 App 不再声明 `mcp.json`，也不需要 `app/mcp.provide`。

服务端有路径白名单（前缀匹配，`OFFICE_ALLOWED_DIRS`）：App 启动它时默认设为**用户主目录 + C–H 盘根**，可用环境变量 `OFFICE_ALLOWED_DIRS` 覆盖。不设这个，任何调用都会被服务端自己的沙箱挡回。

### 两条插图路线怎么选

- **要可检索、可编辑、与原件风格一致** → 文字图注：`office_insert_caption`（或 `office_insert_photo_rows` 里已内置图注生成）
- **只要打印好看、不需要检索** → 烧进图片：`office_add_image` 的 `caption` 参数（图注变像素，且字号会随图缩放）

插入前先用 `office_doc_geometry` 或 `office_images` 看清目标文档自己的几何与已有图形，**不要跨文档套模板**。

## 子进程与临时文件清理（只清自己起的）

Office/WPS 的自动化实例在异常路径上会泄漏（微软自己的 server-side Automation 文档就把这点列为不支持服务端自动化的理由之一：实例一旦泄漏，之后该应用的调用会集体卡死）。本 App 只做**精确清理**：

- **用后**：任务结束（含取消/超时）时，按“任务启动前快照 → 结束后快照”的**差集**结束本轮新起的进程；临时目录（`ot_com_*` / `unelevated_*` / `lo_profile_*` / `officecli_probe_*`，仅限一小时前的）一并清掉。
- **用前**：只做快照，**不杀任何东西**。
- **手动兜底**：工具 `office_cleanup`（需 `confirm=true`）与面板「清理残留」按钮。

**为什么不再按“无窗口标题”清进程（v0.8.1 行为变更）**：早期版本把“无主窗口标题”当作自动化残留的判据，并自动结束这类进程。这条规则是错的、且有破坏性——WPS 会同时开多个 `wps.exe` 辅助进程，而编辑窗口在大文档还没加载完时标题是空的，把它们结束掉会直接搞挂用户正在编辑的会话。2026-10-09 出现过真实事故：一次任务的“用前清理”杀掉了 7 个进程，用户的 WPS 会话随之崩溃并弹出恢复提示。现在对这类进程**只报告、不处理**，由用户自行判断。

- **实现要点**：App 运行在受限执行环境里，进程视图受限，看不到（也就杀不掉）外部起的实例；因此枚举与 COM 走同一条 `explorer.exe` 派生通道，在隔离之外执行。
- **残留怎么手动处理**：确认某个 `wps.exe` / `WINWORD.EXE` 确实是你不再需要的残留后，再自己结束它。App 不会替你判断。

## 首次使用（在新电脑上装）

这个 App 会用到三样本机组件，**新机器上可能都没有**。装上后先跑 `office_deps_status`（面板里叫「环境依赖」）看缺什么：

| 组件 | 作用 | 缺了会怎样 | 怎么装 |
| --- | --- | --- | --- |
| **LibreOffice** | 转换引擎（旧格式现代化、转 PDF/HTML/CSV…） | 转换工具不可用；若装了 Office/WPS 仍可回落做到「旧格式→现代格式」 | `office_deps_install(target="libreoffice", confirm=true)`，或官网下载 |
| **officecli** | docx/xlsx/pptx 的读写桥接 | `office_cli_run` / `office_cli_status` 报未找到 | `office_deps_install(target="officecli", confirm=true)`：从官方 Release 下载 `officecli-win-x64.exe`、校验 sha256 后装到 `%LOCALAPPDATA%\OfficeCli` |
| **uv / uvx** | 运行随包托管的 `timeverse-office-doc` MCP（74 项工具） | 随包 MCP 起不来，其余功能不受影响 | `office_deps_install(target="uv", confirm=true)`，或 <https://docs.astral.sh/uv/> |

设计原则：**缺哪个就只少哪一块**，不会因为少一样就整个 App 不能用。

- 没有 LibreOffice、但有 WPS/MS Office：**能**做旧格式→现代格式与导出 PDF（走 COM，经 explorer 降权子进程）；其余目标格式仍需 LibreOffice。两者都没有：转换功能完全不可用。
- 没有 officecli：转换功能照常，只是不能读写 docx/xlsx/pptx
- 没有 uv：托管的那套 MCP 工具不出现，App 自己的工具不受影响

## 托管的 MCP

包内 `mcp.json` 声明了一个 stdio 连接器 `timeverse-office-doc`（`uvx timeverse-office-doc-mcp`），需要 `app/mcp.provide` 授权与全局 MCP 开关同时满足才会启动。它提供 Word/Excel/PPT/PDF 共 74 项专用工具。

装了本 App 后，**原来单独配置的同名连接器必须移除（remove），不能只停用**：工具名前缀绑定连接器 ID，被停用的旧连接器仍会占名，导致本 App 托管的那套工具调不到。移除后工具名仍是 `timeverse-office-doc_*`。

（officecli 自己也有 MCP，但它只暴露 1 个「跑一条命令」的工具，与本 App 的 `office_cli_run` 重复，因此不托管。）

## 装法与授权

在「应用管理」里授予：**运行外部程序**（`app/process.spawn`）、**提供 MCP**（`app/mcp.provide`）。`app/process.spawn` 是启动期闸门，授权后需重载 App 才生效。App 只把任务 JSON 与状态写入自己的 `dataDir`，用户路径的枚举和转换由已授权子进程完成。

## 与 office MCP 的关系

本 App 把 `timeverse-office-doc` 这套 74 项工具**随包托管**（`mcp.json` + `app/mcp.provide`），不再需要用户单独配置该连接器。

而 officecli 虽然自己也能起 MCP，但它只暴露 **1 个**「跑一条 officecli 命令」的工具，与本 App 的 `office_cli_run` 重复，因此不托管它。

两件事需要分清：officecli 提供 docx/xlsx/pptx 的通用命令读写；这 74 项工具提供语义化的专用工具，另有 PDF 操作与模板能力，两者互补。

## 许可

本 App 以 **MIT** 许可发布，见 `LICENSE`。

它**不打包、不内联**任何第三方组件：LibreOffice（MPL-2.0）、OfficeCLI（Apache-2.0）、timeverse-office-doc-mcp（MIT）都是运行时检测并按需引导安装的外部程序/包。各自的许可、作用与获取方式见 `THIRD-PARTY-NOTICES.md`。

## 异步任务与恢复

转换提交后立即返回任务 ID；面板每 1.5 秒轮询任务列表，可看阶段、当前文件、进度、终态和错误。状态保存在 App 的 `dataDir/jobs`，服务重启后不会盲目重跑；在途任务标记为“中断、结果待检查”，因为部分文件可能已生成。

超时保护：单文件 15 分钟无进度、或总任务 60 分钟，会请求结束子进程树（包括本次自己拉起的 Office/soffice 进程）。进程终止后已写出的产物保留，需人工检查。

