# Excel 小工具箱

面向 Windows 桌面版 Microsoft Excel 的 VBA 加载项，提供删除空行、一键去重和美化表头三个功能，操作前自动备份工作表。

源码、构建脚本及测试工作簿位于 [`ExcelToolbox/`](ExcelToolbox/)。完整使用说明、构建要求和验证范围见 [项目文档](ExcelToolbox/README.md)。

在安装了桌面版 Excel 和 Inno Setup 6 的 Windows 机器上，双击 [`ExcelToolbox/Build.cmd`](ExcelToolbox/Build.cmd) 即可构建；只需 XLAM 时可运行 `Build.cmd -SkipInstaller`。

当前仓库提供源码，尚未交付或验证 Windows 安装包。Windows Excel 的实际运行与安装验证仍需在目标环境完成。
