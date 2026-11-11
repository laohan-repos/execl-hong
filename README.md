# Excel 小工具箱

面向 Windows 桌面版 Microsoft Excel 的 VBA 加载项，提供删除空行、一键去重和美化表头三个功能，操作前自动备份工作表。

源码、构建脚本及测试工作簿位于 [`ExcelToolbox/`](ExcelToolbox/)。完整使用说明、构建要求和验证范围见 [项目文档](ExcelToolbox/README.md)。

安装包下载：[`ExcelToolboxSetup.exe`](https://github.com/laohan-repos/execl-hong/releases/download/v1.0.0/ExcelToolboxSetup.exe)，完整附件和更新说明见 [v1.0.0 Release](https://github.com/laohan-repos/execl-hong/releases/tag/v1.0.0)。本地构建产物位于 `ExcelToolbox/dist/`。

本次交付的是在目标电脑自动生成加载项的安装包。运行环境需要 Windows 10/11 和桌面版 Microsoft Excel 2010 或以上；仅有 WPS 无法安装。安装包包含全部构建资源，无需联网、Python 或 Inno Setup；安装前会调用本机 Excel 生成 XLAM 并执行回归测试，成功后才安装。

本机实测 `Excel.Application` 指向 WPS，因此未进行 Microsoft Excel 内的完整安装和功能验证。已验证 EXE 编译、内嵌资源提取与哈希，以及不支持的 Office 环境会在安装前停止。

重新构建此安装包：`powershell.exe -NoProfile -ExecutionPolicy Bypass -File ExcelToolbox\build\Build-Installer.ps1 -IsccPath "C:\路径\ISCC.exe"`。此构建模式只需要 Inno Setup 6。

如需预编译 XLAM 的传统安装包，在装有 Microsoft Excel 和 Inno Setup 6 的机器上运行 `ExcelToolbox\Build.cmd`；只需 XLAM 时可运行 `Build.cmd -SkipInstaller`。
