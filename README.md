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

## 自动清理（僵尸进程与临时文件）

Office/WPS 的自动化实例在异常路径上会泄漏（微软自己的 server-side Automation 文档就把这点列为不支持服务端自动化的理由之一：实例一旦泄漏，之后该应用的调用会集体卡死）。本 App 的做法是**用前清一次、用后清一次**，把残留掐在萌芽：

- **判断依据**：自动化实例没有主窗口标题；你自己打开着的 Word/Excel/WPS 文档有标题，因而不受影响。LibreOffice 是例外（`soffice.bin` 即使在有窗口时也是它），所以只清命令行里带我方 `lo_profile_*` 标记的实例。
- **用前**：任务启动前先做一次无窗口残留清扫，避免上一轮崩溃留下的实例把这一轮拖死。
- **用后**：任务结束（含取消/超时）后，按“启动前快照 → 结束后快照”的差集精确结束本轮新起的进程，再兜一次无窗口清扫；临时目录（`ot_com_*` / `unelevated_*` / `lo_profile_*` / `officecli_probe_*`，仅限一小时前的）一并清掉。
- **手动兜底**：工具 `office_cleanup`（需 `confirm=true`），面板里对应「清理」按钮。
- **实现要点**：App 运行在受限执行环境里，进程视图受限，看不到（也就杀不掉）外部起的实例；因此清扫与 COM 走同一条 `explorer.exe` 派生通道，在隔离之外执行。

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

