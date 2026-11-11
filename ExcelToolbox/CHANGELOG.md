# 更新日志

## v1.0.0 — 2026-11-11

首次发布 Windows EXE 安装包，提供删除空行、一键去重和美化表头三个 Excel 加载项功能，操作前自动备份工作表。

- 新增 `Build-Installer.ps1`：构建机无需安装 Microsoft Excel，也能生成包含 VBA、Ribbon 和测试数据的安装包。
- 安装时在目标电脑调用 Microsoft Excel 生成 XLAM，并执行自动化回归；成功后才复制文件及注册加载项。
- 修复安装脚本的 Pascal 注释语法错误，补充 32/64 位 Office 注册表视图检查，拒绝把 WPS 的 COM 别名识别为 Microsoft Excel。
- 新增构建转录日志及纯 PowerShell Ribbon 注入选项，确保目标电脑无需 Python。
- 将临时 VBA 工程访问设置纳入还原流程，补充安装前须知、用户指南和实际验证范围。

运行要求：Windows 10/11，桌面版 Microsoft Excel 2010 或以上。不支持 WPS、网页版 Excel 或仅 Office 运行时。安装前请保存并关闭所有 Excel 窗口。

安装时会临时启用当前用户的 VBA 工程对象模型访问并还原原设置；安装目录会登记为 Excel 受信任位置。企业策略禁止相关能力时可能无法安装或加载。

已验证 EXE 编译、六个内嵌构建资源提取及 SHA256、缺少受支持 Office 时在安装前停止，以及构建脚本拒绝 WPS。**尚未在 Microsoft Excel 上完成安装、卸载、功能区显示及功能运行验证。** 安装包未签名。
