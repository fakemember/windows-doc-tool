# Phase 5: Windows 打包与交付详细实操手册

本手册为 **Phase 5（打包与交付）** 的标准 SOP（标准作业程序）。按照本指引，可在任意一台 64 位 Windows 10/11 电脑上，将本项目打包为可直接分发给终端用户的绿色目录或单一安装包（`.exe`）。

---

## 最终交付物规格

打包完成后将产出两类交付物（位于 `dist/` 目录）：
1. **绿色免安装包（One-Dir 文件夹 / Portable ZIP，强烈推荐给无管理员权限环境）**：`dist/wdt/`（压缩为 `WindowsDocTool_v0.1.0_portable.zip`），内含 `wdt.exe`、`runtime/libreoffice/`（裁剪版）、`runtime/7za.exe`、`licenses/`。解压到任意目录（如桌面/下载/D盘）双击即可秒开，**零 UAC 提权提示、零注册表写入、完全不需要管理员权限**。
2. **标准 Windows 安装包（Setup EXE）**：`dist/WindowsDocTool_Setup_v0.1.0.exe`（体积约 90~130MB），已配置 `PrivilegesRequired=lowest`，默认安装至当前用户的 `AppData\Local\Programs`，**普通受限员工账号双击即可直接安装，不会弹出黄色/蓝色管理员密码拦截窗口**。

---

## 针对“无管理员权限（No Admin Access）”的特别指引

如果你的 Windows 电脑是公司受限机（没有 Admin 权限）：

### 场景 A：只想在这台电脑上【运行使用】软件
- **直接使用绿色免安装包（Portable ZIP）**：拷贝 `WindowsDocTool_v0.1.0_portable.zip` 过去解压，直接双击 `wdt.exe`。它是一个自包含的独立目录，所有临时沙箱都在系统默认允许当前用户读写的 `%TEMP%` 目录下运行，**100% 不需要管理员权限**。
- **即使使用 `.exe` 安装包**：因为我们在 `installer.iss` 里配置了 `PrivilegesRequired=lowest`，安装程序绝不会向 Windows 申请 UAC 管理员提权，普通用户账号即可顺利装入个人目录并生成桌面图标。

### 场景 B：需要在这台电脑上【亲自打包构建】软件
在受限机上打包也完全可行：
1. **安装 Python**：从 Python 官网下载安装程序，运行到第一步时选择 **"Install for me only"（仅为当前用户安装）**，安装路径为 `C:\Users\<用户名>\AppData\Local\Programs\Python`，不需要管理员权限。
2. **免装 Inno Setup**：如果在受限机上无法安装 Inno Setup 软件，**直接跳过第五步**！执行完第四步（PyInstaller 构建出 `dist\wdt`）后，在 PowerShell 中直接运行一条内置压缩命令：
   ```powershell
   Compress-Archive -Path dist\wdt\* -DestinationPath dist\WindowsDocTool_v0.1.0_portable.zip
   ```
   产出的 `WindowsDocTool_v0.1.0_portable.zip` 就是完美的开箱即用终极交付物。
3. **最轻松捷径**：直接使用我们提供的 GitHub Actions（云端是微软托管的高配 Windows 环境，自带管理员权限），在云端自动编译出 `WindowsDocTool_v0.1.0_portable.zip`，再下载复制到受限电脑上即可。

---

## 准备工作（Windows 环境要求）

在执行打包前，请确保打包机已具备以下工具：

1. **Python 3.10 ~ 3.12 64-bit**：
   - 官网下载：[https://www.python.org/downloads/](https://www.python.org/downloads/)
   - 安装时勾选 **"Add Python to PATH"**。
2. **Inno Setup 6（安装包编译器，免费开源）**：
   - 官方下载：[https://jrsoftware.org/isdl.php](https://jrsoftware.org/isdl.php)
   - 安装默认路径通常为 `C:\Program Files (x86)\Inno Setup 6\ISCC.exe`。
3. **7-Zip Extra 独立程序（`7za.exe`）**：
   - 官方下载：[https://www.7-zip.org/download.html](https://www.7-zip.org/download.html)（下载 7-Zip Extra / 7za standalone）。
4. **官方 LibreOffice Portable（便携版）**：
   - 官方下载：[https://www.libreoffice.org/download/portable-versions/](https://www.libreoffice.org/download/portable-versions/)（下载 `.paf.exe`，例如 LibreOffice 24.x 或 25.x Portable）。

---

## 标准打包 6 步流程

### 第一步：获取代码与初始化虚拟环境

打开 Windows 终端（PowerShell 或 CMD）：

```powershell
# 1. 进入代码仓库根目录
cd C:\path\to\windows-doc-tool

# 2. 创建并激活虚拟环境
python -m venv .venv
.\.venv\Scripts\activate

# 3. 安装项目依赖（含 GUI 与构建工具）
pip install -e ".[gui]"
pip install pyinstaller
```

---

### 第二步：解压并裁剪 LibreOffice 便携版

1. **解压官方便携版**：
   双击运行下载的 `LibreOfficePortable_xx.paf.exe`，安装解压到临时目录（如 `C:\temp\LibreOfficePortable`）。
   解压后核心目录通常位于：`C:\temp\LibreOfficePortable\App\libreoffice`（该目录下应包含 `program\` 和 `share\` 文件夹）。

2. **执行自动化裁剪脚本**：
   在项目根目录下运行裁剪脚本，裁剪工具会自动执行安全黑名单裁剪，并**自动发起无头 RTF→DOCX 转换冒烟测试**，验证裁剪未损坏核心组件：

   ```powershell
   python build_scripts\slim_libreoffice.py --source "C:\temp\LibreOfficePortable\App\libreoffice" --output "runtime\libreoffice"
   ```

   - **预期输出**：
     - 输出裁剪规则命中统计（已安全移除 help/template/extensions 及非 en/zh 语言包）。
     - 显示 `[冒烟测试] 启动无头转换验证: RTF -> DOCX... ✓ 转换成功`。
     - 打印体积优化统计：未压缩体积从 ~480MB 缩减至 ~190MB（降幅约 60%）。

---

### 第三步：补充辅助运行时与许可证文件

1. **放置 `7za.exe`**：
   将下载解压出的 `7za.exe` 复制到项目 `runtime\` 根目录：
   ```powershell
   # 确认文件存在
   Test-Path runtime\7za.exe   # 应返回 True
   ```

2. **补充开源许可证文件（合规必需）**：
   参考 `licenses\README.txt` 中的指引，下载以下官方协议纯文本并放入 `licenses\` 目录：
   - `licenses\LICENSE.LibreOffice.MPL2.txt`（[MPL-2.0](https://www.mozilla.org/en-US/MPL/2.0/)）
   - `licenses\LICENSE.PySide6.LGPL3.txt`（[LGPL-3.0](https://www.gnu.org/licenses/lgpl-3.0.html)）
   - `licenses\LICENSE.7Zip.LGPL21.txt`（[7-Zip License](https://www.7-zip.org/license.txt)）
   - `licenses\LICENSE.pypdf.BSD3.txt`（[BSD-3-Clause](https://raw.githubusercontent.com/py-pdf/pypdf/main/LICENSE)）

---

### 第四步：使用 PyInstaller 进行 One-Dir 构建

运行自动化打包脚本：
```powershell
python build_scripts\build_exe.py --entry src\wdt\main.py --runtime-dir runtime
```

> **注意**：
> 本项目**严格禁止**使用 `--onefile` 单文件模式！因为内置便携运行时上百兆，单文件模式每次冷启动都会将所有组件解压至系统的 `%TEMP%` 目录，不仅每次启动卡顿数秒，而且频繁占用大量磁盘 I/O。
> 采用 `--onedir` 模式可实现秒级启动与即时响应。

构建完成后，检查 `dist\wdt\` 目录：
```powershell
# 验证主程序与内置运行时已收集完毕
Test-Path dist\wdt\wdt.exe
Test-Path dist\wdt\runtime\libreoffice\program\soffice.exe
```

此时可直接双击运行 `dist\wdt\wdt.exe` 验证图形界面是否秒开。

---

### 第五步：使用 Inno Setup 编译安装程序

调用 Inno Setup 命令行编译器 `ISCC.exe`：

```powershell
# 如果 Inno Setup 在默认路径：
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=0.1.0 build_scripts\installer.iss
```

- **构建结果**：
  在 `dist\` 目录下生成：`WindowsDocTool_Setup_v0.1.0.exe`。
- **安装包特性**：
  - 采用 LZMA2 超高压缩算法，最终体积仅 **90~130MB**。
  - 普通用户权限即可安装（默认装入用户的 Local AppData 或 Program Files，不强制要求 UAC 提权）。
  - 安装后自动生成桌面图标与开始菜单项。

---

### 第六步：干净 Windows 虚拟机最终验收（Checklist）

找一台**全新的、未安装 Python、未安装 Office、未安装外部 LibreOffice** 的干净 Windows 10/11 虚拟机（或沙箱 Windows Sandbox）：

1. **安装验证**：
   - 双击 `WindowsDocTool_Setup_v0.1.0.exe`，安装过程流畅无卡顿、无缺失 DLL 弹窗。
2. **环境诊断测试**：
   - 打开 CMD 运行：`wdt.exe detect`。
   - 验证：输出应显示 `[转换引擎] 状态: 已就绪，来源: bundled`，7-Zip 为 `可用`。
3. **GUI 界面验证**：
   - 双击桌面快捷方式启动图形界面；
   - 拖拽 1 个 `.rtf` 文件到界面，选择 `经典规约` 样式，勾选同时导出 `PDF`，点击“开始处理”；
   - 验证任务进度条推进，产出 `.docx` 与 `.pdf`，耗时正常，排版与目录完整。
4. **大文件极限压测**：
   - 将 200MB+ 真实 Listing 文档拖入，执行端到端处理，监控任务管理器：内存峰值在可控范围（<2GB），无僵尸 `soffice.bin` 进程残留。
5. **卸载验证**：
   - 从 Windows 设置中卸载，程序文件与快捷方式完全移除，无流氓服务残留。

---

## 备选方案：通过 GitHub Actions 云端全自动打包

如果你不想在本地配置 Windows 环境，可以在 GitHub 上创建私有仓库并推送代码，配合 CI 脚本直接在 GitHub 云端 Windows 虚拟机中自动完成上述打包流程并输出 `.exe` 安装包下载链接。
