# 第三方组件与参考项目说明

本 App 自身以 **MIT** 许可发布（见 `LICENSE`）。

它**不打包、不内联**下面任何一个第三方组件：它们都是外部程序或外部包，由本 App 在运行时**检测**（`office_deps_status` / `scripts/env_probe.ps1`），并在用户明确同意后**引导安装**。因此本仓库不构成对它们的再分发。

## 一、运行时依赖（外部程序 / 外部包，均不随包分发）

| 组件 | 许可 | 作用 | 获取方式 |
| --- | --- | --- | --- |
| [LibreOffice](https://www.libreoffice.org/) | Mozilla Public License 2.0 | 文档转换引擎（旧格式现代化、导出 PDF/HTML/CSV 等） | 用户自行安装，或 `office_deps_install(target="libreoffice")` 走 winget 安装 |
| [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)（iOfficeAI） | Apache License 2.0 | docx/xlsx/pptx 的读写桥接 | 用户自行安装，或 `office_deps_install(target="officecli")` 从官方 Release 下载并校验 sha256 后装到 `%LOCALAPPDATA%\OfficeCli` |
| [timeverse-office-doc-mcp](https://pypi.org/project/timeverse-office-doc-mcp/)（TimeVerse） | MIT | 本包 `mcp.json` 声明托管的 MCP 连接器，提供 Word/Excel/PowerPoint/PDF 共 74 项工具 | 由宿主以 `uvx timeverse-office-doc-mcp` 从 PyPI 拉取运行 |
| Microsoft Office / WPS Office（可选） | 各自商业许可 | 无 LibreOffice 时的兜底转换引擎（COM 自动化） | 用户已安装则使用；本 App 不安装、不改动 |

## 二、参考项目（仅借鉴思路，未复制代码）

| 项目 | 许可 | 借鉴点 |
| --- | --- | --- |
| [BatchOfficeFormatConverter](https://github.com/) （旧格式批量转换工具） | MIT | 早期确定「用本机 Office 做 doc/xls/ppt → 现代格式」这一路线的可行性；本 App 的转换实现为独立编写 |
| [github-cli](https://github.com/hayou2002/hana-github-cli)（Hana App） | 见其仓库 | 作为 Hana v2 App 的结构模板：`sdk.process.resolveExecutable` + 子进程 + 工具注册 + 面板路由的写法 |

## 三、说明

- 上述组件的许可条款以各自官方发布为准；本文件仅作说明，不替代其许可文本。
- 若你所在环境对某个组件的许可有额外要求，请自行核对后再使用。
- 本 App 不会在未经用户同意的情况下安装任何外部程序。
