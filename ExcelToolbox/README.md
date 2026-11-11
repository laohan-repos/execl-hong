# Excel 小工具箱 v1.0.0

面向 Windows 桌面版 Microsoft Excel 的 VBA 加载项。所有操作针对当前活动的普通工作簿，实际改动前会确认并复制整张工作表作为备份。

| 功能 | 行为 |
|---|---|
| 删除空行 | 删除当前表 UsedRange 中既无值也无公式的整行 |
| 一键去重 | 按选中连续区域的全部列组合判重，不区分大小写，每组保留第一条；可指定首行为表头 |
| 美化表头 | 选中区域首行设为蓝底、白色粗体、居中并自动调整选中列宽 |

去重删除工作表整行，请让选区覆盖需要保留的全部列。恢复数据请从橙色标签的 `备份_` 工作表复制，不要依赖 Ctrl+Z；请自行保存工作簿。

## 本次 EXE 交付（2026-11-11）

产物：`dist/ExcelToolboxSetup.exe`。

本机 Windows 的 `Excel.Application` 实测指向 WPS 12.0，未安装可用于构建的 Microsoft Excel。为生成可分发的 EXE，本次使用 **在目标电脑生成加载项** 的安装模式：

1. 安装前检查 Microsoft Excel；需要 Windows 10/11 和桌面版 Microsoft Excel 2010 或以上，不支持 WPS、网页版 Excel或仅 Office 运行时。
2. 提取内嵌 VBA、Ribbon、构建脚本及测试工作簿到安装程序的临时目录。
3. 通过目标电脑的 Excel COM 生成真正的 `ExcelToolbox.xlam`，注入 Ribbon，并执行三个功能及状态恢复的自动化回归。
4. 仅在生成及测试成功后，才安装文件、登记加载项和受信任位置。

目标电脑无需 Python、Inno Setup 或联网。安装生成阶段可能需要几分钟，请先保存并关闭所有 Excel 窗口。

生成期间临时启用当前用户的「信任对 VBA 工程对象模型的访问」，随后还原原设置；企业策略禁止访问时无法生成。构建日志保存在 `%LOCALAPPDATA%\ExcelToolbox\Logs`。

已完成的本机验证：Inno Setup 6.7.4 实际编译成功；内嵌的六个构建资源按预期目录提取且 SHA256 全部一致；生成的 EXE 在缺少受支持 Excel 时正确退出且未创建安装目录；主构建脚本在修改 VBA 工程访问设置前拒绝 WPS。

**验证范围：尚未在 Microsoft Excel 上完成安装、卸载、功能区显示及功能运行验证。安装时的自动化回归也需要在目标 Excel 环境实际执行，本机打包检查不能代替它。** 安装包未签名。

## 安装与使用

双击 `dist/ExcelToolboxSetup.exe`，阅读须知后安装。完成后重新打开 Microsoft Excel，顶部出现「Excel 小工具箱」选项卡。

安装目录：`%APPDATA%\Microsoft\AddIns\ExcelToolbox`。仅安装到当前用户，无需管理员权限。

安装程序登记该目录为 Excel 受信任位置，不修改全局宏禁用或受保护视图设置。若企业策略禁止宏或受信任位置，仍可能无法加载。

菜单未出现时，运行开始菜单中的「重新注册加载项（菜单未出现时）」并重启 Excel；也可在 Excel → 文件 → 选项 → 加载项 → 管理「Excel 加载项」→ 转到 → 浏览，选择安装目录中的 `ExcelToolbox.xlam` 并勾选。

卸载前关闭 Excel，在 Windows 设置的已安装应用中卸载「Excel 小工具箱」。卸载移除工具文件、加载项注册和工具目录的受信任位置，用户工作簿及其中的备份表不受影响。

## 重新构建

### 本机无需 Excel 的安装包

仅需 Windows PowerShell 5.1 和 Inno Setup 6：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File build\Build-Installer.ps1 -IsccPath "C:\路径\ISCC.exe"
```

编译器位于项目的 `build\sfx\InnoSetup\ISCC.exe` 或常见安装位置时可省略 `-IsccPath`。该脚本向安装脚本传入 `/DBuildOnInstall`，生成在目标电脑编译 XLAM 的安装包。

### 预编译 XLAM 的传统安装包

构建机需要 Microsoft Excel 2010+ 和 Inno Setup 6。关闭 Excel 后运行 `Build.cmd`，将执行 XLAM 生成、Ribbon 注入、EXE 编译及 Excel COM 回归。

参数：`-SkipInstaller` 只生成 XLAM，`-SkipTests` 跳过 COM 回归，`-SkipXlam` 复用已有 XLAM，`-IsccPath` 指定编译器。

底层 `build\Build-ExcelToolbox.ps1` 还支持 `-PowerShellRibbon` 强制使用内置的 Ribbon 注入实现，以及 `-LogPath` 保存转录日志。PowerShell 脚本保持 UTF-8 BOM，以兼容 Windows PowerShell 5.1；VBA 源码通过 UTF-8 解码后写入 COM，避免中文乱码。

两种安装包模式使用相同输出文件名；重新构建会替换 `dist\ExcelToolboxSetup.exe`。

## 源码与测试数据

- `src/vba/`：三个功能、备份与确认、目标工作簿守卫、Ribbon 回调。
- `src/ribbon/customUI.xml`：功能区定义。
- `src/installer/ExcelToolboxSetup.iss`：Inno Setup 安装脚本，支持预编译与目标电脑生成两种模式。
- `src/installer/UserGuide.txt`：安装后提供的用户说明。
- `build/Build-ExcelToolbox.ps1`：Excel COM 构建及回归。
- `build/Build-Installer.ps1`：无本机 Excel 的安装包构建入口。
- `build/inject_ribbon.py`：可选 Python Ribbon 注入器；另有 PowerShell 内置实现。
- `build/Register-ExcelAddin.vbs`：可选重新注册工具。
- `tests/ExcelToolbox_TestData.xlsx`：删除空行、去重、美化表头与说明四张工作表。

COM 回归期望：空行测试 10 行变 7 行；去重测试 8 行变 4 行且保留首条；美化测试蓝底白字粗体；每次实际改动创建一张备份。另有 32 组清理宏状态恢复场景，覆盖正常处理、无需处理、工作表保护和备份失败。

已知功能边界：不专门处理筛选、合并单元格或 Excel 表格的删除约束；保护或共享限制可能导致备份或删除失败。UsedRange 中仅有格式的行视为空行。纯 VBA 无原生 Declare，设计上不依赖 Office 位数，但 32/64 位的实际运行仍需在对应 Excel 环境验证。
