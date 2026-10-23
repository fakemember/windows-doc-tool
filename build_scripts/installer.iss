; =====================================================================
; Windows Doc Tool (wdt) - Inno Setup 6 安装包编译脚本
; =====================================================================
; 说明：
; 本脚本用于将 PyInstaller one-dir 产物、裁剪后的便携式 LibreOffice 运行时
; 以及合规许可证文本打包为 Windows 原生安装程序（.exe 安装包）。
;
; 编译方式示例（通过 ISCC 命令行）：
;   ISCC.exe /DAppVersion=0.1.0 /DSourceDir=..\dist\wdt build_scripts\installer.iss
;
; 宏定义参数说明（可通过命令行 /D 传入覆盖）：
;   - AppVersion: 软件版本号（必需从命令行传入，默认兜底 "0.1.0"）
;   - AppName:    软件显示名称（默认 "Windows Doc Tool"）
;   - AppExeName: 主程序可执行文件名（默认 "wdt.exe"）
;   - SourceDir:  PyInstaller 打包产物目录（默认 "..\dist\wdt"）
;   - RuntimeDir: 裁剪版便携运行时目录（默认 "..\runtime"）
;   - LicenseDir: 许可证文件目录（默认 "..\licenses"）
;   - OutputDir:  安装包生成输出目录（默认 "..\dist"）
; =====================================================================

#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

#ifndef AppName
  #define AppName "Windows Doc Tool"
#endif

#ifndef AppExeName
  #define AppExeName "wdt.exe"
#endif

#ifndef SourceDir
  #define SourceDir "..\dist\wdt"
#endif

#ifndef RuntimeDir
  #define RuntimeDir "..\runtime"
#endif

#ifndef LicenseDir
  #define LicenseDir "..\licenses"
#endif

#ifndef OutputDir
  #define OutputDir "..\dist"
#endif

[Setup]
; 应用程序唯一 GUID 标识符
AppId={{5A8E192E-73C1-4D3B-A954-58D9C03C2809}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Windows Doc Tool Team
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir={#OutputDir}
OutputBaseFilename=WindowsDocTool_Setup_v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
; 采用普通用户权限模式，无需管理员 UAC 提权，不写系统注册表
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
ArchitecturesInstallIn64BitMode=x64compatible
; 中文与英文消息文件
MessagesFile=compiler:Default.isl,compiler:Languages\ChineseSimplified.isl

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Tasks]
; 桌面快捷方式创建选项
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 1. PyInstaller one-dir 打包主程序产物
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 2. 便携运行时组件 (LibreOffice 裁剪版 + 7za)
Source: "{#RuntimeDir}\*"; DestDir: "{app}\runtime"; Flags: ignoreversion recursesubdirs createallsubdirs
; 3. 合规许可证文本
Source: "{#LicenseDir}\*"; DestDir: "{app}\licenses"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 开始菜单快捷方式
Name: "{autoprograms}\{#AppName}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autoprograms}\{#AppName}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
; 桌面快捷方式（受 Tasks 控制）
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; 安装完成后可选启动程序
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

; [UninstallRun]
; 本产品完全绿色化运行，不注册全局系统服务与 COM 组件，不写注册表启动项，
; 卸载时直接删除安装目录文件即可干净移除，故无 [UninstallRun] 指令。
