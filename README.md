# TFLOW (Clinical TFL & Document Pipeline)

专为 Windows 平台打造的高性能临床 TFL（Tables, Figures, Listings）与科研/统计大文件批处理流水线工具。基于纯 Windows 原生 Microsoft Word COM 自动化引擎，提供一站式 `.doc`/`.rtf` 极速转 `.docx`/`.pdf`，以及交付级原生动态 TOC 目录编排与页码重算能力。

![TFLOW Logo](src/wdt/ui/assets/tflow_logo.svg)

---

## 核心特性

- **纯 Windows 原生 Microsoft Word COM 自动化**：彻底摆脱第三方开源排版引擎对 Word 复杂表格、横竖版混排、嵌套页眉页脚的排版失真，100% 保持 Word 原生视觉渲染与打印品质。
- **交付级原生动态 TOC（Dynamic TOC）**：插入合规的 Word 原生 `TOC \o "1-3"` 域与分节符，正文首节精准重置页码（`start="1"`）。导出前全自动刷新所有 StoryRanges 与目录对象，用户在 Word 中按 `F9` 亦可无缝刷新。
- **单遍高速渲染直出**：告别旧式临时 PDF 测量回填的双倍耗时流程，结合后台拼写/语法波浪线静默优化，大幅压降多页复杂表格的渲染时间。
- **现代化 GUI 与 CLI 双模交互**：既可通过现代桌面应用（PySide6）拖拽执行，也可通过灵活的命令行（`tflow run` / `tflow convert` / `tflow toc`）集成至批处理脚本与统计分析流水线。

---

## 快速上手

### 命令行使用

安装后可直接使用 `tflow` 命令行工具（亦保留 `wdt` 别名）：

```bash
# 启动现代化桌面图形界面
tflow gui

# 环境自检与 Word COM 引擎诊断
tflow detect

# 一键流水线：RTF 转 DOCX + 注入原生动态目录 + 直接导出 PDF
tflow run ./inputs -o ./outputs --pdf --page-mode dynamic

# 批量转换 RTF/DOC 为 DOCX
tflow convert ./inputs -o ./outputs --to docx

# 单独为已有的 DOCX 文档注入临床经典目录
tflow toc report.docx --style classic --depth 3
```

### 目录样式

TFLOW 内置 4 套面向医药法规与专业报告的目录排版样式：
- `modern`：现代商务（默认，推荐）
- `classic`：经典点线（法规推荐，适合 CSR/TFL 交付）
- `regulatory`：法规清单风格
- `minimal`：极简纯净风格
