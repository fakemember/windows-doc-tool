================================================================================
Windows Doc Tool (wdt) 依赖组件许可证清单与合规指引
================================================================================

本项目发布包（Inno Setup 安装包与绿色解压包）集成了若干开源组件与运行时。
根据各开源项目的许可证要求，在发布二进制分发包时，本目录必须包含所有依赖
组件的完整许可文本文件，并在关于页面及安装说明中明确告知用户。

--------------------------------------------------------------------------------
一、发布包必须包含的许可文件清单
--------------------------------------------------------------------------------

1. LibreOffice 裁剪版便携运行时
   - 对应文件/目录: runtime/libreoffice/
   - 许可证类型: Mozilla Public License Version 2.0 (MPL-2.0)
   - 本地许可文件名: LICENSE-LibreOffice-MPL-2.0.txt
   - 官方许可全文: https://www.mozilla.org/en-US/MPL/2.0/
   - 许可说明地址: https://www.libreoffice.org/about-us/licenses/
   - 源代码获取渠道: https://www.libreoffice.org/about-us/source-code/
   - 合规说明:
     LibreOffice 采用 MPL-2.0 许可证。随产品分发其二进制程序（如 soffice.com、
     soffice.bin 及相关 DLL）属于“Larger Work”组合分发。必须保留原有版权与
     NOTICE 声明，并在安装目录保留 MPL-2.0 完整文本与源码获取指引。

2. PySide6 (Qt for Python)
   - 对应文件/目录: PyInstaller 打包产物中的 Qt6/PySide6 动态链接库与 Python 扩展
   - 许可证类型: GNU Lesser General Public License Version 3 (LGPL-3.0)
   - 本地许可文件名: LICENSE-PySide6-LGPL-3.0.txt
   - 官方许可全文: https://www.gnu.org/licenses/lgpl-3.0.html
   - 许可说明地址: https://doc.qt.io/qtforpython-6/licenses.html
   - 源代码获取渠道: https://code.qt.io/cgit/pyside/pyside-setup.git
   - 合规说明:
     PySide6 及其依赖的 Qt 库采用 LGPL-3.0 许可证。本软件采用动态链接方式
     加载 PySide6 库，并在分发包中附带 LGPL-3.0 许可证正文与开源源码获取指引，
     符合 LGPL-3.0 动态链接的合规要求。

3. 7-Zip / 7za.exe
   - 对应文件/目录: runtime/7za.exe（仅在超大文件 TOC 原位更新时调用）
   - 许可证类型: GNU Lesser General Public License Version 2.1 or later (LGPL-2.1+)
   - 本地许可文件名: LICENSE-7Zip-LGPL-2.1.txt
   - 官方许可全文: https://www.7-zip.org/license.txt
   - 许可说明地址: https://www.gnu.org/licenses/old-licenses/lgpl-2.1.html
   - 源代码获取渠道: https://www.7-zip.org/download.html
   - 合规说明:
     7-Zip / 7za.exe 采用 LGPL 许可证（部分模块含 unRAR 限制或 BSD 3-clause 许可）。
     分发独立 7za.exe 二进制程序时，必须随包附带其许可证全文与官方源码获取地址。

4. pypdf
   - 对应文件/目录: pypdf Python 依赖包（用于 PDF 页码截取与切片）
   - 许可证类型: BSD 3-Clause License
   - 本地许可文件名: LICENSE-pypdf-BSD-3.txt
   - 官方许可全文: https://raw.githubusercontent.com/py-pdf/pypdf/main/LICENSE
   - 源代码获取渠道: https://github.com/py-pdf/pypdf
   - 合规说明:
     pypdf 采用宽松的 BSD-3 许可证，分发二进制产物时需保留其版权声明与免责条款。

5. Windows Doc Tool (wdt) 应用程序本体
   - 对应文件/目录: src/ 源码及编译出的 wdt.exe / wdt 控制台与 GUI 入口
   - 许可证类型: 自有专有许可（All Rights Reserved）
   - 本地许可文件名: LICENSE-wdt.txt

--------------------------------------------------------------------------------
二、打包分发检查项 (Release Checklist)
--------------------------------------------------------------------------------

在执行 build_scripts/build_exe.py 及 build_scripts/installer.iss 打包前，
发布维护者必须确保本目录下已放置以下文本文件：
  [ ] licenses/README.txt (本文档)
  [ ] licenses/LICENSE-LibreOffice-MPL-2.0.txt
  [ ] licenses/LICENSE-PySide6-LGPL-3.0.txt
  [ ] licenses/LICENSE-7Zip-LGPL-2.1.txt
  [ ] licenses/LICENSE-pypdf-BSD-3.txt
  [ ] licenses/LICENSE-wdt.txt

Inno Setup 安装脚本会将 licenses/ 目录完整打包至安装根目录下的 licenses/ 子目录中。
