# Excel 小工具箱 v1.0.0

给 Microsoft Office Excel 加一个「Excel 小工具箱」功能区选项卡，三个按钮：

| 按钮 | 作用 |
|---|---|
| **删除空行** | 删除当前工作表中「整行既无值、也无公式」的行 |
| **一键去重** | 对选中区域按所有列组合去重，重复行只保留第一次出现的那条 |
| **美化表头** | 选中区域首行设为蓝色背景 + 白色粗体 + 居中，并自动调整列宽 |

技术栈：**VBA 核心 + XLAM 加载项 + Ribbon XML(customUI) + Inno Setup 6 EXE 安装包**。
所有操作只作用于用户当前打开的工作簿，不会修改加载项自身文件。

---

## 0. 先说清楚：本次在什么环境里开发、哪些是真做出来的

按你的要求我先做了环境检查，真实结果如下（不是推测）：

| 检查项 | 实测结果 |
|---|---|
| 操作系统 | **macOS 15.3.2 (Darwin arm64)** —— 不是 Windows |
| Microsoft Office Excel | **仅 Excel for Mac 16.85**，没有 Windows 版 Excel |
| Excel COM 自动化 | **不可用**（macOS 无 COM，`pywin32` 不存在） |
| Office 版本 / 位数 | Mac 版 16.85 / arm64；**无法读取 Windows Office 的位数** |
| Inno Setup / ISCC | **未安装**（Windows 专有，无 Mac 版，也没有 Wine） |
| Windows 虚拟机 | Parallels Desktop 已装，但无任何 `.pvm` 虚拟机，`prlctl` 无法初始化 |
| XLAM + Ribbon 生成条件 | 只能做 OOXML / XML 层级验证；`vbaProject.bin` 必须由 Excel 编译，本机无法产出 |
| 可用工具 | Python 3.13 + openpyxl 3.1.5、zip/unzip、系统自带真实 `Solver.xlam` 可作 OOXML 样本 |

**由此产生的硬约束（必须坦白）：**

- `.xlam` 里的 `vbaProject.bin` 是 Excel 的 VBA 编译器输出的二进制工程，
  纯文本 `.bas` 无法等价替代，**只有 Windows Excel 通过 COM 才能生成**。
- Inno Setup 的 `ISCC.exe` 只有 Windows 版，**EXE 只能在 Windows 上编译**。
- 所以我**没有**在这台 Mac 上伪造一个能点但没功能的空壳 `.xlam`，也**没有**谎称生成了 `.exe`。

**我的做法**：把所有源码、脚本、测试数据和安装包逻辑全部写完并尽最大可能真实验证，
然后把剩下的两步做成 **Windows 上双击一个 `Build.cmd` 全自动完成**，
你仍然不需要打开 VBA 编辑器、不需要手动复制任何代码。

### 已真实验证 / 待 Windows 验证

| 步骤 | 状态 | 证据 |
|---|---|---|
| 环境检查 | ✅ 真实验证 | 本报告第 0 节的实测输出 |
| 项目目录创建 | ✅ 真实验证 | 见下方目录结构 |
| 完整 VBA 源码（3 个功能 + Ribbon 回调 + 备份/确认） | ✅ 已产出 | `src/vba/*.bas`、`*.cls`；源码级走查（过程签名、回调对应、无 Declare、无 32/64 位依赖） |
| 测试用 Excel 工作簿 | ✅ **已生成并自检通过** | `tests/ExcelToolbox_TestData.xlsx`；openpyxl 读回断言：10 行中第 4/7/8 行为空、8 行去重数据含 4 种唯一组合、4 张工作表齐全 |
| Ribbon XML | ✅ **已校验** | well-formed、命名空间 2006/01 正确、4 个按钮 + 5 个回调（`rxDeleteEmptyRows` / `rxRemoveDuplicates` / `rxBeautifyHeader` / `rxAbout` / `rxOnLoad`）与 VBA 中 `Public Sub` 一一对应 |
| Ribbon 注入 OOXML 包的逻辑 | ✅ **在真实 XLAM 上实测通过** | 用系统自带 `Solver.xlam`（真实的 12 部件 OPC 包、含 435,712 B 的 `vbaProject.bin`）跑 `build/inject_ribbon.py`：新增 `customUI/customUI.xml`、追加 `ui/extensibility` 关系、追加 Content-Type Override；全部 XML 部件通过 well-formed 校验；**重复注入幂等**（部件稳定 13 个、无重复关系）；**除 `.rels` 和 `[Content_Types].xml` 外所有原有部件字节完全未变** |
| Inno Setup 安装脚本（检测/注册/受信任位置/卸载） | ✅ 已产出，⏳ 未编译 | `src/installer/ExcelToolboxSetup.iss`（Inno Setup 6 Pascal Script） |
| 生成 `ExcelToolbox.xlam` | ⏳ 需 Windows Excel | `build/Build-ExcelToolbox.ps1` 第 2 步自动完成 |
| 生成 `ExcelToolboxSetup.exe` | ⏳ 需 Windows + Inno Setup 6 | `build/Build-ExcelToolbox.ps1` 第 4 步自动完成 |
| 安装 / 卸载 / Ribbon 显示 / 3 个功能 COM 回归测试 | ⏳ 需 Windows | `build/Build-ExcelToolbox.ps1` 第 5 步已写好自动化断言，Windows 上自动跑 |

> ⏳ = 代码已就绪但未在本机执行过。我不会把「写好脚本」说成「已验证通过」。

---

## 1. 目录结构

```
ExcelToolbox/
├── README.md                         本文件
├── Build.cmd                         ★ Windows 上一键构建入口（双击这个）
├── build/
│   ├── Build-ExcelToolbox.ps1        主构建脚本：环境检查→XLAM→Ribbon→EXE→COM 测试
│   ├── inject_ribbon.py              Ribbon customUI 注入器（跨平台，本机已实测）
│   ├── make_test_workbook.py         生成测试工作簿（openpyxl，本机已实测）
│   └── Register-ExcelAddin.vbs       菜单未出现时的可选重注册工具（英文界面）
├── src/
│   ├── vba/
│   │   ├── modToolbox.bas            三个功能 + 备份/确认/目标工作簿校验
│   │   ├── modRibbon.bas             Ribbon onLoad / onAction 回调
│   │   └── ThisWorkbook.cls          加载项工作簿事件
│   ├── ribbon/customUI.xml           自定义功能区定义
│   └── installer/
│       ├── ExcelToolboxSetup.iss     Inno Setup 6 脚本
│       └── InfoBefore.txt            安装前安全须知
├── tests/ExcelToolbox_TestData.xlsx  ★ 测试用工作簿（已真实生成）
└── dist/
    ├── README.txt                    产物目录说明
    ├── ExcelToolbox.xlam             ← Windows 上运行 Build.cmd 后生成
    └── ExcelToolboxSetup.exe         ← Windows 上运行 Build.cmd 后生成
```

---

## 2. 在 Windows 上生成两个成品（一次性）

### 需要准备

1. Windows 10/11
2. **桌面版 Microsoft Office Excel**（Office 2010 / 14.0 及以上；32 位 / 64 位都可以，
   加载项是纯 VBA，无 `Declare`、无 ActiveX，位数无关）
3. **Inno Setup 6**（官网免费下载，装时保持默认路径即可）
4. 把整个 `ExcelToolbox` 文件夹拷过去，**保持目录结构**
5. **关闭所有 Excel 窗口**

### 操作

> **双击 `Build.cmd`**，然后按一次任意键。
> 不需要右键管理员运行，不需要打开 Excel，不需要打开 VBA 编辑器，不需要导入宏。

脚本会自动按顺序做完并打印 `[OK]` / `[FAIL]`：

1. 检查 Excel / COM / 版本 / 位数 / Inno Setup / VBA 工程访问条件
2. 校验测试工作簿
3. 临时开启「信任对 VBA 工程对象模型的访问」→ 用 Excel COM 新建工作簿 →
   导入 `modToolbox.bas`、`modRibbon.bas` → 写入 `ThisWorkbook` 事件代码 →
   设 `IsAddin = True` → `SaveAs(..., 55)` 生成 `dist\ExcelToolbox.xlam` →
   **立即还原**该信任设置
4. 注入 `customUI.xml`（优先用已验证的 `inject_ribbon.py`，无 Python 时用内置 PowerShell 实现），
   校验部件 / 关系 / Content-Type / `vbaProject.bin` / 全部 XML
5. 调用 `ISCC.exe` 编译 `dist\ExcelToolboxSetup.exe`
6. 用 Excel COM 打开加载项 + 测试工作簿，跑 3 个功能的自动化断言（含备份工作表数量），
   打印结果和 SHA256

可用参数：`Build.cmd -SkipInstaller`（只出 XLAM）、`-SkipTests`、`-SkipXlam`。

默认构建必须生成安装包：找不到 `ISCC.exe` 时会在环境检查阶段以退出码 1 结束，
不会把已有的旧 EXE 当成本次产物。可用 `Build.cmd -IsccPath "C:\路径\ISCC.exe"` 指定编译器；
只有显式使用 `-SkipInstaller` 时才允许仅生成 XLAM 并返回成功。

构建脚本保存为 **UTF-8 with BOM**，项目的 `.editorconfig` 会保持该编码，
以兼容 `Build.cmd` 调用的 Windows PowerShell 5.1。VBA 模块也会显式按 UTF-8 读取，
再通过 COM 写入工程，避免文件导入时按系统 ANSI 误读中文。

---

## 3. 最终用户的安装体验（这就是你要的四步）

1. 双击 **`ExcelToolboxSetup.exe`**
2. 看一眼「安装前须知」→ 点 **下一步 / 安装**（默认装到当前用户，不弹 UAC）
3. 打开 **Microsoft Office Excel**
4. 顶部功能区出现 **「Excel 小工具箱」** 选项卡，直接用三个按钮

### 可能出现的「安全提示」，如实说明

| 现象 | 原因 | 怎么处理 |
|---|---|---|
| Windows SmartScreen 显示「未知发布者 / Windows 已保护你的电脑」 | 安装包**未做代码签名**（本项目没有证书） | 企业分发建议用代码签名证书签名；个人使用点「更多信息 → 仍要运行」 |
| 首次打开 Excel 弹出「安全警告：宏已被禁用」或菜单不出现 | 组策略关闭了「受信任位置」，或安装时 Excel 没关干净 | 见第 6 节手动回退；或让 IT 把安装目录加入白名单 |
| 公司策略开启了「禁用所有受信任位置」/「阻止来自 Internet 的宏」 | 企业安全策略，工具无法也不应绕过 | 联系 IT：对本目录做策略白名单，或用可信证书签名 EXE + XLAM 后重新分发 |

### 关于宏安全（重点，没有绕过）

- **没有**关闭宏安全设置、**没有**修改「禁用所有宏」「受保护的视图」等全局策略、
  **没有**要求用户打开 VBA 编辑器或导入宏。
- 安装程序做的唯一一件事是把工具箱**自己的安装目录**登记为 Excel 的
  **受信任位置 (Trusted Location)**，这是微软官方推荐的加载项部署方式：
  `HKCU\Software\Microsoft\Office\<ver>\Excel\Security\Trusted Locations\LocationExcelToolbox`。
- 加载项注册用两条标准的当前用户注册表项：
  `...\Excel\Options\OPEN<n>`（自动打开，脚本会找第一个空闲下标，**不覆盖**其它加载项）
  和 `...\Excel\Add-in Manager`（在加载项列表里登记）。
- 卸载时会删除这两条注册表项、受信任位置键、所有文件，并清理可能残留的 `OPEN*` 项。
- 构建/测试脚本里的 `Application.AutomationSecurity = 1` **只作用于脚本临时启动、
  随后立即 Quit 的隐藏 Excel 实例**，不写注册表、不持久化、不改变用户 Excel 的信任中心配置；
  它不属于最终用户的安装环节。

---

## 4. 三个功能的确切行为与数据可恢复性

Excel 的 VBA 删除动作**不会进入撤销栈**，所以 `Ctrl+Z` 不可靠。为此：

- **删除空行**：统计 UsedRange 内 `CountA = 0` 的行 → 弹出确认（显示将删几行）→
  把当前整张工作表复制到同一工作簿末尾，命名为 `备份_删除空行_YYMMDD_HHNNSS`（橙色标签）
  → 自下而上删除 → 报告删除行数与备份表名。
- **一键去重**：要求先选中连续矩形区域 → 询问「第一行是否为表头」→
  用 `Scripting.Dictionary` 生成带长度前缀的行指纹（避免分隔符碰撞，错误值用显示文本），
  不区分大小写 → 从下往上删除重复行，**每组保留第一条** →
  执行前同样确认 + 全表备份。注意：**删除的是工作表整行**，对话框里已明确提示，
  请先让选区覆盖所有需要保留的列。
- **美化表头**：选中区域第一行 → 蓝底 `RGB(0,112,192)`、白字加粗、水平垂直居中、
  底部细边框、行高 24、选中区域 `EntireColumn.AutoFit` → 同样先确认 + 备份。
- 恢复数据：从 `备份_xxx` 工作表把内容复制回原工作表即可（不再需要时会一直保留，
  你可以自行删除）。
- 所有入口都有统一守卫：只取 `ActiveWorkbook`，若当前是加载项自身
  (`wb Is ThisWorkbook`)、其它加载项 (`wb.IsAddin`) 或没有打开工作簿，会直接提示退出，
  **不会改动工具箱自身文件**。
- 表格/受保护/共享工作簿导致备份创建失败时，操作直接放弃，不改动任何数据。

---

## 5. 卸载

- 开始菜单 → 「Excel 小工具箱」→ **卸载 Excel 小工具箱**
  （或 设置 → 应用 → 安装的应用 → Excel 小工具箱 → 卸载）
- 卸载前请先关闭所有 Excel 窗口，卸载程序会检测并提示。
- 卸载会删除 `%APPDATA%\Microsoft\AddIns\ExcelToolbox` 整个目录、
  上述注册表项与受信任位置。用户工作簿和里面的备份工作表**不会被删除**。

---

## 6. 需要人工介入时，明确点哪里

正常情况下**不需要**。只有以下两种例外：

**A. 构建时若脚本报「无法访问 VBA 工程对象模型」（权限被组策略锁死）**
> Excel → **文件** → **选项** → **信任中心** → **信任中心设置…** → **宏设置** →
> 勾选 **「信任对 VBA 工程对象模型的访问」** → 确定 → 关闭 Excel → 重新双击 `Build.cmd`。
> （脚本正常时会自动临时打开并还原这项，无需你操作。）

**B. 安装后 Excel 顶部没有出现「Excel 小工具箱」**
> 先试开始菜单 → 「Excel 小工具箱」→ **重新注册加载项（菜单未出现时）**，然后重启 Excel。
>
> 仍不行则手动勾选一次（这是唯一需要点 Excel 的回退路径）：
> Excel → **文件** → **选项** → **加载项** → 底部「管理」下拉选 **Excel 加载项** →
> 点 **转到…** → 点 **浏览…** → 定位到
> `%APPDATA%\Microsoft\AddIns\ExcelToolbox\ExcelToolbox.xlam` → 确定 →
> 在列表里勾选 **ExcelToolbox** → 确定 → 重启 Excel。

---

## 7. 测试工作簿与自动化断言

`tests/ExcelToolbox_TestData.xlsx` 含 4 张表：

| 工作表 | 数据 | 期望 |
|---|---|---|
| 删除空行测试 | 10 行，第 4/7/8 行完全为空 | 删 3 行剩 7 行，新增 1 张备份表 |
| 去重测试 | 8 行无表头数据，4 种唯一组合 | 剩 4 行，第一行为 `张三/北京/100`，新增 1 张备份表 |
| 美化表头测试 | A1:E6 | A1 背景 `12611584`、字色 `16777215`、加粗，新增 1 张备份表 |
| 说明 | 人工测试指引 | — |

`Build-ExcelToolbox.ps1` 第 5 步会用 Excel COM 静默调用
`Toolbox_DeleteEmptyRows(True)` / `Toolbox_RemoveDuplicates(True)` /
`Toolbox_BeautifyHeader(True)` 并逐条断言（`Silent=True` 只跳过对话框，备份照常创建）。

删除空行和去重会保存并恢复执行前的计算模式、事件开关和屏幕刷新开关，
恢复逻辑也覆盖备份创建失败及删除出错的退出路径。
COM 回归另外增加了 **32 组状态恢复场景**：两个清理宏 × 四种原始状态组合
（自动/手动计算，事件与刷新开关组合）× 正常完成、无需处理、工作表保护、备份失败。
每组同时检查状态、处理后的数据行数与备份数量。这些新增 COM 断言需在 Windows Excel 上执行；
macOS 上的静态校验不代表它们已通过。

---

## 8. 已知限制

- 仅支持 **Windows 桌面版 Microsoft Office Excel**（Office 2010+）。
  网页版 Excel、WPS、Mac 版 Excel、仅装 Office 运行时均不支持（VBA + Ribbon + 注册表机制所限）。
- `删除空行` 依据 `UsedRange` 判断；若工作表有大量仅带格式的「幽灵单元格」，
  UsedRange 会偏大，但不影响正确性（这些行本来就会被判定为空）。
- `一键去重` 删除整行，不处理筛选状态、合并单元格、Excel 表格(ListObject) 的特殊约束；
  遇到这些情况会走错误提示并保留备份，不会静默改坏数据。
- `customUI.xml` 的 `imageMso` 图标名在不同 Office 版本可能有个别不显示，
  **不影响按钮功能**；需要换图标直接改 `src/ribbon/customUI.xml` 里对应的 `imageMso` 即可。
- 安装包默认**未代码签名**，企业分发建议签名。
- 本机的 Excel for Mac 不能用来生成 Windows 版 XLAM（VBA 工程二进制不通用），
  也不能编译 Inno Setup；因此第 2 节必须在 Windows 上执行。

---

## 9. 一句话总结

源码、菜单定义、安装逻辑、测试数据和「一键构建 + 一键测试」链路都已完整交付；
在本机（macOS）已真实验证了环境检查、测试工作簿、Ribbon XML 和在**真实 XLAM 包**上的
Ribbon 注入与幂等性。把整个文件夹拷到 Windows 机器上双击 `Build.cmd`，
即可拿到你要的 `ExcelToolbox.xlam` 与 `ExcelToolboxSetup.exe`，
并最终用户只需「双击 EXE → 安装 → 打开 Excel → 看到菜单」四步。
