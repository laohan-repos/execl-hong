@echo off
chcp 65001 >nul 2>&1
title Excel 小工具箱 - 一键构建
echo.
echo   ============================================================
echo    Excel 小工具箱 v1.0.0  一键构建
echo    Windows + 桌面版 Excel + Inno Setup 6
echo   ============================================================
echo.
echo   本脚本会自动完成：
echo     1. 环境检查（Excel / COM / 版本 / 位数 / Inno Setup）
echo     2. 用 Excel 把 VBA 编译进 ExcelToolbox.xlam
echo     3. 注入 Ribbon 菜单
echo     4. 编译 ExcelToolboxSetup.exe
echo     5. 用 Excel COM 跑三个功能的自动化测试
echo.
echo   请先关闭所有 Excel 窗口，然后按任意键继续...
pause >nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build\Build-ExcelToolbox.ps1" %*
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" (
  echo   [完成] 产物位于 %~dp0dist
) else (
  echo   [失败] 退出码 %RC%，请向上滚动查看 [FAIL] 行。
)
echo.
pause
exit /b %RC%
