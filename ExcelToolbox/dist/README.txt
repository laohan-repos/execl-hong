本目录用于存放最终交付的两个二进制产物：

    ExcelToolbox.xlam          Excel 加载项（含编译后的 VBA 工程 + Ribbon 菜单）
    ExcelToolboxSetup.exe      Inno Setup 制作的安装包

这两个文件需要在 Windows 上生成，原因：
  * .xlam 内的 vbaProject.bin 是 Excel 的 VBA 编译器产出的二进制，
    只有 Excel（Windows 版，通过 COM 自动化）能生成，纯源码无法替代。
  * Inno Setup 的 ISCC.exe 只有 Windows 版，无法在 macOS/Linux 上编译 EXE。

生成方式（在 Windows 上，装好桌面版 Excel + Inno Setup 6 后）：
  双击项目根目录的 Build.cmd，等待脚本跑完即可。
  脚本会自动完成环境检查 -> 编译 XLAM -> 注入 Ribbon -> 编译 EXE -> COM 自动化测试。

生成前请先把整个 ExcelToolbox 文件夹拷到 Windows 机器（保留目录结构）。

本仓库目前已提供（可直接使用/审阅）：
  src/vba/*.bas, *.cls        完整 VBA 源码
  src/ribbon/customUI.xml     功能区菜单定义
  src/installer/*.iss         Inno Setup 安装脚本（检测/注册/受信任位置/卸载）
  tests/ExcelToolbox_TestData.xlsx   测试用工作簿
  build/*.ps1, *.py, *.vbs    一键构建、注入、测试、注册工具
  README.md                   完整说明与验证报告

不要把一个「没有 vbaProject.bin 的空壳 xlam」当成成品：那样菜单会出现但按钮无效。
所以本目录不放假文件。
