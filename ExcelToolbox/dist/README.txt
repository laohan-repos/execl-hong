Excel 小工具箱 — EXE 交付说明（2026-11-11）

运行 ExcelToolboxSetup.exe 安装。
需要 Windows 10/11 和桌面版 Microsoft Excel 2010 或以上，不支持 WPS。

这是在目标电脑生成 XLAM 的安装包：先调用本机 Excel 编译并回归测试，
成功后再安装文件与注册加载项。目标电脑无需联网、Python 或 Inno Setup。
安装前请保存并关闭所有 Excel，生成及测试可能需要几分钟。

安装时临时开启 VBA 工程对象模型访问，完成后还原原设置。
日志：%LOCALAPPDATA%\ExcelToolbox\Logs
安装目录：%APPDATA%\Microsoft\AddIns\ExcelToolbox

本机仅有 WPS，已经验证打包、资源提取与哈希，以及不支持环境的正确退出。
尚未在 Microsoft Excel 上完成安装、卸载、功能区与功能运行验证。
安装包未签名。

完整说明见项目 ExcelToolbox\README.md。
