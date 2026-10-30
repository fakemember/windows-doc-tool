# Windows Doc Tool (wdt)

专为 Windows 平台打造的高性能、高并发办公与临床/科研大文件批处理工具。支持 `.doc`/`.rtf` 极速转 `.docx`/`.pdf`（含灰度归档），并提供多套精美、合规的定制 TOC 目录注入样式。

## 核心特性

- **免装 LibreOffice 的内置裁剪运行时**：随软件分发内置便携版最小运行时（未压缩约 200~250MB，7z 压缩后安装包仅 80~130MB），解压即用。不写 Windows 注册表，无需系统管理员权限。
- **无 COM、无弹窗、多进程稳定并发**：彻底摒弃脆弱的微软 Word COM (`win32com`) 自动化，采用无头（Headless）沙箱进程与分块调度机制。杜绝弹窗中断与僵尸进程，支持 4~8 进程线性并行加速。
- **智能剥离旧目录与字体/版心自适应**：注入新目录前在字节级自动剥离源文件中由 Word 生成的旧目录（`<w:sdt>` 内容控件块及 `TOC \o` 字段段落），自动吸收孤立换页符；自适应当前节版心宽度计算右制表位，并自动探测正文字体（docDefaults / Normal）保持一致视觉体验。
- **实测精准页码（measured 模式）**：通过微型隐藏锚点与临时 PDF 逐页实测，彻底解决流式文档因无硬分页符导致的页码错误及 LibreOffice 无头导出 PDF 目录空白问题，交付物自动彻底清理锚点，干净无痕。

---

## 与 LibreOffice 的关系

本工具依赖 LibreOffice 的无头转换能力，但针对企业与临床科研场景进行了专门的便携化与沙箱改造：

### 1. 内置便携运行时目录结构
发布版软件根目录下内置裁剪后的 LibreOffice 运行时：
```
windows-doc-tool/
├── wdt.exe                         # 主程序
├── runtime/
│   ├── libreoffice/                # 裁剪版 LibreOffice 便携运行时
│   │   └── program/
│   │       ├── soffice.com         # 控制台批处理包装器
│   │       ├── soffice.exe / soffice.bin
│   │       └── ...                 # 核心动态链接库
│   ├── 7za.exe                     # 超大文件原位更新工具
│   └── licenses/                   # 依赖组件许可证
└── licenses/                       # 开源许可文本声明
```

### 2. 探测与回退顺序
调度引擎按照以下优先级自动探测转换引擎：
1. 本地内置运行时目录：`./runtime/libreoffice/program/soffice.com`；
2. 操作系统环境变量 `PATH` 中的 `soffice`；
3. Windows 常见安装路径（如 `C:\Program Files\LibreOffice\program\soffice.exe`）。

### 3. 独立沙箱隔离机制
每次启动转换任务时，系统均通过 `-env:UserInstallation=file:///...` 为工作进程指派独立的临时沙箱配置目录，各进程完全隔离，彻底避免多进程间的配置锁冲突。

### 4. 开源许可合规
内置的 LibreOffice 运行时遵循 Mozilla Public License Version 2.0 (MPL-2.0)。分发时保留原有版权与 NOTICE 声明，完整许可文本见 `licenses/LICENSE-LibreOffice-MPL-2.0.txt`，源码获取渠道详见 `licenses/README.txt`。

---

## 安装说明

### 开发者安装
需要 Python 3.10+ 环境。建议使用 `uv` 进行开发环境管理：

```bash
# 克隆代码仓库
git clone <repo-url> windows-doc-tool
cd windows-doc-tool

# 创建虚拟环境并安装开发依赖
uv venv
uv pip install -e ".[dev]"

# 若需开发 GUI 界面，安装含 Qt 依赖的配置
uv pip install -e ".[gui,dev]"
```

### 终端用户安装
终端用户无需安装 Python 或 LibreOffice，可选用以下任一方式：
- **安装包版（推荐）**：下载官方发布的 `WindowsDocTool_Setup_v*.exe`（基于 Inno Setup 6 打包），支持当前普通用户权限安装，自动生成开始菜单与桌面快捷方式。
- **绿色便携版**：下载 `WindowsDocTool_v*_portable.zip`，解压至任意目录即可双击 `wdt.exe` 启动使用。

> 🛠️ **开发者打包指引**：详见 [docs/PHASE_5_PACKAGING_GUIDE.md](docs/PHASE_5_PACKAGING_GUIDE.md)，包含详细的 LibreOffice 裁剪步骤、PyInstaller one-dir 打包、Inno Setup 编译命令以及 GitHub Actions 云端全自动流水线。

---

## CLI 命令行规格与用法

### 命令总览

```bash
wdt detect                                          # 环境诊断
wdt convert <target> [-r] [-o DIR] [-j N] [--to docx|pdf|both] [--gray] [--pages A-B] [-f]
wdt toc <files...> [--style classic|modern|regulatory|minimal] [--depth 1|2|3]
        [--title TXT] [--position document-start|before-first-heading|replace-placeholder|bookmark:NAME]
        [--page-mode static-restart|static-offset|static-in-place|measured|dynamic] [--page-offset N] [--no-backup]
wdt run <target> [-r] [-o DIR] [-j N] [--style ...同上...] [--page-mode measured|...] [--pdf] [--gray] [-f]
wdt gui                                             # 图形界面（需 .[gui] 额外依赖）
```

### 1. `wdt detect`：运行环境诊断
检查内置/外部 LibreOffice、7za.exe、CPU 核心数与可用内存状态：
```bash
wdt detect
```

### 2. `wdt convert`：批量格式转换
将 `.doc`/`.rtf` 转换为 `.docx` 或 `.pdf`。
- `-r, --recursive`：递归遍历子目录。
- `-o, --outdir`：指定输出目录。
- `-j, --jobs`：设置并发工作者数（默认根据 CPU 与内存自动计算）。
- `--to docx|pdf|both`：目标输出格式（默认 `docx`）。
- `--gray`：导出灰度 PDF。**注意：此选项需要内置或外部 LibreOffice 版本 ≥ 7.4**。
- `--pages A-B`：仅提取并转换指定页码区间（基于 `pypdf`，仅支持单文件）。
- `-f, --force`：覆盖已存在的目标文件。

### 3. `wdt toc`：DOCX 目录注入
直接对 `.docx` 文件注入规范排版的目录与页码。
- `--style classic|modern|regulatory|minimal`：目录视觉样式（默认 `classic`）。
- `--depth 1|2|3`：提取标题层级深度（默认 `2`）。
- `--title TXT`：自定义目录标题文字。
- `--position document-start|before-first-heading|replace-placeholder|bookmark:NAME`：目录插入位置。
- `--page-mode static-restart|static-offset|static-in-place|measured|dynamic`：页码计算模式（默认 `static-restart`；指定 `measured` 实测模式需可用引擎）。
- `--page-offset N`：页码偏移补偿量。
- `--no-backup`：就地修改，不保留 `.bak` 备份文件。

### 4. `wdt run`：端到端一体化流水线
级联执行：`.doc`/`.rtf`/`.docx` → 转换为 `.docx` → 注入 TOC →（可选）导出 `.pdf`。
- `--page-mode`：静态系默认走 `measured`（基于微型隐藏锚点实测页码，保证 PDF 目录页码非空且与实际页一致）。
- `--pdf`：在完成 DOCX 目录注入后级联生成 PDF。
- `--gray`：导出的 PDF 为灰度版本（需要 LO ≥ 7.4）。

**使用示例：**
```bash
# 示例 1: 完整流水线处理单个大文件：转 docx、注入经典目录并基于实测页码最终生成 PDF
wdt run trial_listing.rtf -o ./final --style classic --pdf

# 示例 2: 批量将 raw_docs/ 下全部文档转为含 Listing 规范目录的 docx，并输出至 dist/
wdt run ./raw_docs -r -o ./dist -j 4 --style regulatory --depth 2
```

### 5. `wdt gui`：图形界面
拖拽文件/文件夹即可批处理，适合非命令行用户：
```bash
wdt gui
```
- 支持转换/注入/一体化流水线与实测页码选择。

---

## 4 种 TOC 样式一览表

| 样式编号 | 样式名称 | 视觉与排版规范 | 适用场景 |
| :--- | :--- | :--- | :--- |
| **STYLE_1** | **经典规约 (Classic Dot)** | 宋体/Times New Roman，五号字，右对齐密集点导线，标准层级缩进 | 监管报送、学术论文、正式合同 |
| **STYLE_2** | **现代商务 (Modern Accent)** | 微软雅黑/Calibri，一级标题加粗商务蓝 `#1F4E79`，紧凑行距 | 商业企划、项目汇报 |
| **STYLE_3** | **临床 Listing (Regulatory)** | 严格遵循 SAS / NMPA Listing 规范，10.5pt，精确大纲级别与物理页码 | 临床试验统计报告与巨型附表 |
| **STYLE_4** | **极简优雅 (Minimal Clean)** | 标题与页码纯空白隔离，无导线，页码小字号弱化 | 白皮书、产品手册 |

---

## 已知限制

使用本工具前，请知悉以下技术限制与边界：

1. **远古格式**：Word 6.0/95 之前二进制 `.doc` 转换保真度有限；
2. **内嵌对象降级**：OLE 对象、宏、表单控件在 docx/pdf 输出中被剥离或静态化；
3. **Word 与 LibreOffice 分页渲染差异**：Word 与 LibreOffice 排版引擎在跨页断行、表格跨页等流式算法上可能存在 $\pm 1$ 页差异。DOCX 内部写入的是标准 `PAGEREF` 域，在 Word 中可通过 `Ctrl+A` 然后按 `F9` 强制刷新为 Word 自身渲染的页码；
4. **两遍渲染耗时**：`run --pdf` 采用 `measured` 实测模式时，需在无头环境下进行临时 PDF 测算与最终 PDF 导出两遍渲染，整体耗时较单次渲染约翻倍（基于 LibreOffice UNO 原生 1-Pass 刷新域作为后续版本性能优化演进方向）；
5. **未签名分发**：无代码签名证书时，Windows SmartScreen 与部分杀软（PyInstaller 产物易误报）会告警。

---

## 许可证说明

- **Windows Doc Tool (wdt) 应用程序本体**：自有专有许可。
- **LibreOffice Portable 裁剪运行时**：遵循 [Mozilla Public License Version 2.0 (MPL-2.0)](https://www.mozilla.org/en-US/MPL/2.0/)。
- **PySide6 (Qt for Python)**：遵循 [GNU Lesser General Public License Version 3 (LGPL-3.0)](https://www.gnu.org/licenses/lgpl-3.0.html)。
- **7-Zip (7za.exe)**：遵循 [GNU Lesser General Public License Version 2.1 or later (LGPL-2.1+)](https://www.7-zip.org/license.txt)。
- **pypdf**：遵循 [BSD 3-Clause License](https://raw.githubusercontent.com/py-pdf/pypdf/main/LICENSE)。

所有发布二进制包必须在 `licenses/` 目录下携带上述各开源组件的完整许可文本。详细许可对应关系与源码获取途径参见 `licenses/README.txt`。
