<#
.SYNOPSIS
    Excel 小工具箱 —— Windows 一键构建脚本

.DESCRIPTION
    在「Windows + 桌面版 Excel + Inno Setup 6」上自动完成：
      0. 环境检查（Excel / COM / 版本 / 位数 / Inno Setup / VBA 工程访问）
      1. 生成（或校验）测试工作簿
      2. 用 Excel COM 把 .bas / .cls 编译进真正的 ExcelToolbox.xlam
      3. 把 customUI.xml 注入 XLAM 的 OPC 包（Ribbon 菜单）
      4. 用 ISCC.exe 编译 ExcelToolboxSetup.exe
      5. 用 Excel COM 跑三个功能的自动化回归测试
      6. 输出产物清单与 SHA256

    用户不需要打开 VBA 编辑器，不需要手动导入任何代码。

.PARAMETER SkipInstaller   只生成 XLAM，不编译 EXE
.PARAMETER SkipTests       跳过 COM 回归测试
.PARAMETER SkipXlam        已有 dist\ExcelToolbox.xlam 时跳过重建
.PARAMETER IsccPath        手工指定 ISCC.exe 路径

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\Build-ExcelToolbox.ps1
    # 或直接双击项目根目录的 Build.cmd
#>
[CmdletBinding()]
param(
    [switch] $SkipInstaller,
    [switch] $SkipTests,
    [switch] $SkipXlam,
    [string] $IsccPath = ''
)

$ErrorActionPreference = 'Stop'
$script:Failed = 0

# COM 后期绑定在部分非英语区域设置下会出问题，统一成 en-US
try { [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US') } catch { }

$Root       = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$SrcVba     = Join-Path $Root 'src\vba'
$RibbonXml  = Join-Path $Root 'src\ribbon\customUI.xml'
$Iss        = Join-Path $Root 'src\installer\ExcelToolboxSetup.iss'
$Dist       = Join-Path $Root 'dist'
$TestXlsx   = Join-Path $Root 'tests\ExcelToolbox_TestData.xlsx'
$XlamPath   = Join-Path $Dist 'ExcelToolbox.xlam'
$ExePath    = Join-Path $Dist 'ExcelToolboxSetup.exe'
$PyInjector = Join-Path $Root 'build\inject_ribbon.py'
$PyTestBook = Join-Path $Root 'build\make_test_workbook.py'

function Write-Step($msg) { Write-Host ''; Write-Host ("=== {0} ===" -f $msg) -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host ("  [OK]   {0}" -f $msg) -ForegroundColor Green }
function Write-Info($msg) { Write-Host ("  [INFO] {0}" -f $msg) -ForegroundColor Gray }
function Write-Warn($msg) { Write-Host ("  [WARN] {0}" -f $msg) -ForegroundColor Yellow }
function Write-Fail($msg) { Write-Host ("  [FAIL] {0}" -f $msg) -ForegroundColor Red; $script:Failed++ }
function Assert-True($cond, $msg) { if ($cond) { Write-OK $msg } else { Write-Fail $msg } }

# ============================================================================
# 0. 环境检查
# ============================================================================
Write-Step '0/6 环境检查'

if ($env:OS -ne 'Windows_NT') {
    throw "当前不是 Windows。本脚本需要 Windows + 桌面版 Excel + Inno Setup 6。"
}
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
Write-Info ("OS           : {0} ({1})" -f $os.Caption, $os.OSArchitecture)

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "检测到 Excel 正在运行。请保存工作并关闭所有 Excel 窗口后再运行本脚本。"
}

# Excel COM 可用性 + 版本
$xlProbe = $null
try {
    $xlProbe = New-Object -ComObject Excel.Application
    $script:OfficeVer = [string]$xlProbe.Version          # 例: 16.0
    $script:ExcelPath = [string]$xlProbe.Path
    Write-OK ("Excel COM    : 可用，Office 版本 {0}" -f $script:OfficeVer)
    Write-Info ("Excel 安装目录: {0}" -f $script:ExcelPath)
} catch {
    throw ("Excel COM 自动化不可用：{0}" -f $_.Exception.Message)
} finally {
    if ($xlProbe) { try { $xlProbe.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlProbe) }
}

# Office 位数
$bit = 'unknown'
try {
    $p = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name Platform -ErrorAction Stop).Platform
    if ($p -eq 'x86') { $bit = '32-bit' } elseif ($p -eq 'x64') { $bit = '64-bit' }
} catch { }
if ($bit -eq 'unknown') {
    if ($script:ExcelPath -match '\(x86\)') { $bit = '32-bit' }
    elseif ($script:ExcelPath -match 'Program Files') { $bit = '64-bit' }
}
Write-Info ("Office 位数   : {0}" -f $bit)

# VBA 工程对象模型访问（仅构建期临时开启，构建结束立即还原）
$script:SecKey = "HKCU:\Software\Microsoft\Office\$($script:OfficeVer)\Excel\Security"
Write-Info ("VBA 工程注册表键: {0}" -f $script:SecKey)

# Inno Setup
function Find-Iscc {
    $c = @()
    if ($IsccPath) { $c += $IsccPath }
    $found = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($found) { $c += $found.Source }
    $c += @(
        'C:\Program Files\Inno Setup 6\ISCC.exe',
        'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        'C:\Program Files\Inno Setup 5\ISCC.exe',
        'C:\Program Files (x86)\Inno Setup 5\ISCC.exe'
    )
    foreach ($k in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1')) {
        if (Test-Path $k) {
            $loc = (Get-ItemProperty $k -ErrorAction SilentlyContinue).InstallLocation
            if ($loc) { $c += (Join-Path $loc 'ISCC.exe') }
        }
    }
    foreach ($item in $c) { if ($item -and (Test-Path $item)) { return $item } }
    return $null
}
$script:Iscc = Find-Iscc
if ($SkipInstaller) { Write-Info '按 -SkipInstaller 仅构建 XLAM，不要求 Inno Setup' }
elseif ($script:Iscc) { Write-OK ("Inno Setup   : {0}" -f $script:Iscc) }
else {
    Write-Fail '未找到 ISCC.exe，无法生成安装包。请安装 Inno Setup 6 或使用 -IsccPath 指定路径；只需 XLAM 时请显式使用 -SkipInstaller。'
    exit 1
}

$pyCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pyCmd) { $pyCmd = Get-Command py -ErrorAction SilentlyContinue }
if ($pyCmd) { Write-OK ("Python       : {0}（用于已验证的 Ribbon 注入器）" -f $pyCmd.Source) }
else { Write-Info 'Python       : 未安装，Ribbon 注入改用内置 PowerShell 实现' }

New-Item -ItemType Directory -Path $Dist -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $Root 'tests') -Force | Out-Null

# ============================================================================
# 1. 测试工作簿
# ============================================================================
Write-Step '1/6 测试工作簿'
if (-not (Test-Path $TestXlsx)) {
    if ($pyCmd) {
        & $pyCmd.Source $PyTestBook $TestXlsx
        if ($LASTEXITCODE -ne 0) { throw '测试工作簿生成失败' }
    } else {
        throw "缺少测试工作簿 $TestXlsx，且本机无 Python。请从项目包中恢复该文件。"
    }
}
Assert-True (Test-Path $TestXlsx) ("测试工作簿存在: {0}" -f $TestXlsx)

# ============================================================================
# 2. 生成真正的 XLAM（Excel 编译 VBA）
# ============================================================================
Write-Step '2/6 用 Excel COM 生成 ExcelToolbox.xlam'

$script:oldVBOM = $null
$script:hadVBOM = $false
function Enable-VbomTemporarily {
    if (-not (Test-Path $script:SecKey)) { New-Item -Path $script:SecKey -Force | Out-Null }
    $p = Get-ItemProperty -Path $script:SecKey -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($p) { $script:hadVBOM = $true; $script:oldVBOM = $p.AccessVBOM }
    New-ItemProperty -Path $script:SecKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Info '已临时开启「信任对 VBA 工程对象模型的访问」（构建结束会还原）'
}
function Restore-Vbom {
    try {
        if ($script:hadVBOM) { Set-ItemProperty -Path $script:SecKey -Name AccessVBOM -Value $script:oldVBOM }
        else { Remove-ItemProperty -Path $script:SecKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        Write-Info '已还原 VBA 工程访问设置'
    } catch { Write-Warn ("还原 AccessVBOM 失败: {0}" -f $_.Exception.Message) }
}

$xlamBuilt = $false
if ($SkipXlam -and (Test-Path $XlamPath)) {
    Write-Info '按 -SkipXlam 跳过重建，使用现有 XLAM'
    $xlamBuilt = $true
} else {
    Enable-VbomTemporarily
    $xl = $null; $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.AskToUpdateLinks = $false

        $wb = $xl.Workbooks.Add()
        try { $comps = $wb.VBProject.VBComponents }
        catch {
            throw ("无法访问 VBA 工程对象模型。请在 Excel 中打开：文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 勾选「信任对 VBA 工程对象模型的访问」，然后重跑本脚本。原始错误：{0}" -f $_.Exception.Message)
        }

        foreach ($f in @('modToolbox.bas', 'modRibbon.bas')) {
            $p = Join-Path $SrcVba $f
            if (-not (Test-Path $p)) { throw "缺少 VBA 源文件: $p" }
            # 显式解码 UTF-8，通过 COM 字符串写入，避免 Import 按系统 ANSI 误读中文。
            $moduleCode = (Get-Content -Path $p -Encoding UTF8 |
                Where-Object { $_ -notmatch '^\s*Attribute\s+' }) -join "`r`n"
            $component = $comps.Add(1)     # vbext_ct_StdModule
            $component.Name = [System.IO.Path]::GetFileNameWithoutExtension($p)
            [void]$component.CodeModule.AddFromString($moduleCode)
            Write-OK ("已导入 {0}" -f $f)
        }

        # ThisWorkbook 事件代码：只取 Option Explicit 起往后的正文，去掉类文件头部
        $clsPath = Join-Path $SrcVba 'ThisWorkbook.cls'
        $lines = @(Get-Content -Path $clsPath -Encoding UTF8)
        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*Option Explicit') { $start = $i; break }
        }
        if ($start -lt 0) { throw 'ThisWorkbook.cls 缺少 Option Explicit 标记行' }
        $twCode = ($lines[$start..($lines.Count - 1)] -join "`r`n")
        [void]$comps.Item('ThisWorkbook').CodeModule.AddFromString($twCode)
        Write-OK '已写入 ThisWorkbook 事件代码'

        try { $wb.IsAddin = $true }
        catch { Write-Warn 'IsAddin 设置失败，将由 SaveAs 文件格式自动处理' }

        if (Test-Path $XlamPath) { Remove-Item $XlamPath -Force }
        $wb.SaveAs($XlamPath, 55)   # 55 = xlOpenXMLAddIn (.xlam)
        Write-OK ("XLAM 已生成: {0}" -f $XlamPath)
        $xlamBuilt = $true
    } finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($xl) { try { $xl.Quit() } catch { } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Restore-Vbom
    }
}
Assert-True ($xlamBuilt -and (Test-Path $XlamPath)) ("XLAM 文件存在，大小 {0} bytes" -f (Get-Item $XlamPath).Length)

# ============================================================================
# 3. 注入 Ribbon customUI
# ============================================================================
Write-Step '3/6 注入 Ribbon 菜单 (customUI.xml)'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { throw "缺少 OOXML 部件: $name" }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}
function Set-ZipEntryText($zip, [string]$name, [string]$text) {
    $old = $zip.GetEntry($name)
    if ($old) { $old.Delete() }
    $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $sw.Write($text) } finally { $sw.Dispose() }
}
function Add-CustomUIPartPS([string]$pkg, [string]$xmlFile) {
    $bytes = [System.IO.File]::ReadAllBytes($xmlFile)
    $offset = 0
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $offset = 3   # 去掉 UTF-8 BOM，避免写入 OOXML 部件时出现前导字符
    }
    $fs = [System.IO.File]::Open($pkg, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update, $false)
        try {
            $ct = Read-ZipEntryText $zip '[Content_Types].xml'
            $rl = Read-ZipEntryText $zip '_rels/.rels'

            if ($ct -notmatch 'customUI/customUI\.xml') {
                $ov = '<Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>'
                $ct = [regex]::Replace($ct, '(</\s*(?:\w+:)?Types\s*>)', ($ov + '$1'))
                Set-ZipEntryText $zip '[Content_Types].xml' $ct
            }
            if ($rl -notmatch 'customUI/customUI\.xml') {
                $rid = 'rIdExcelToolboxUI'; $n = 1
                while ($rl -match ('Id="' + [regex]::Escape($rid) + '"')) { $n++; $rid = "rIdExcelToolboxUI$n" }
                $rel = '<Relationship Id="' + $rid + '" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI.xml"/>'
                $rl = [regex]::Replace($rl, '(</\s*(?:\w+:)?Relationships\s*>)', ($rel + '$1'))
                Set-ZipEntryText $zip '_rels/.rels' $rl
            }
            $existing = $zip.GetEntry('customUI/customUI.xml')
            if ($existing) { $existing.Delete() }
            $e = $zip.CreateEntry('customUI/customUI.xml', [System.IO.Compression.CompressionLevel]::Optimal)
            $s = $e.Open(); $s.Write($bytes, $offset, $bytes.Length - $offset); $s.Dispose()
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}
function Test-CustomUIPart([string]$pkg) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($pkg)
    try {
        $names = @($z.Entries | ForEach-Object { $_.FullName })
        $hasPart = $names -contains 'customUI/customUI.xml'
        $hasVba  = $names -contains 'xl/vbaProject.bin'
        $ct = Read-ZipEntryText $z '[Content_Types].xml'
        $rl = Read-ZipEntryText $z '_rels/.rels'
        $hasCt = $ct -match 'customUI/customUI\.xml'
        $hasRl = $rl -match 'office/2007/relationships/ui/extensibility'
        foreach ($n in $names) {
            if ($n -like '*.xml' -or $n -like '*.rels') {
                $txt = Read-ZipEntryText $z $n
                try { $null = [xml]$txt } catch { throw "XML 部件格式错误: $n" }
            }
        }
        return @{ Part = $hasPart; Relationship = $hasRl; ContentType = $hasCt; Vba = $hasVba; Parts = $names.Count }
    } finally { $z.Dispose() }
}

$usedPython = $false
if ($pyCmd -and (Test-Path $PyInjector)) {
    & $pyCmd.Source $PyInjector $XlamPath $RibbonXml
    if ($LASTEXITCODE -eq 0) { $usedPython = $true }
    else { Write-Warn 'Python 注入器失败，回退到 PowerShell 实现' }
}
if (-not $usedPython) { Add-CustomUIPartPS $XlamPath $RibbonXml }
Write-Info ("注入实现: {0}" -f $(if ($usedPython) { 'inject_ribbon.py（已验证版）' } else { 'PowerShell 内置' }))

$chk = Test-CustomUIPart $XlamPath
Assert-True $chk.Part         'customUI/customUI.xml 部件存在'
Assert-True $chk.Relationship '_rels/.rels 中 ui/extensibility 关系存在'
Assert-True $chk.ContentType  '[Content_Types].xml 中 Override 存在'
Assert-True $chk.Vba          'xl/vbaProject.bin 存在（VBA 已编译进加载项）'
Assert-True ($chk.Parts -ge 4) ("OPC 包共 {0} 个部件，全部 XML 通过 well-formed 校验" -f $chk.Parts)

# ============================================================================
# 4. 编译 EXE 安装包
# ============================================================================
Write-Step '4/6 编译 ExcelToolboxSetup.exe'
$exeBuilt = $false
if ($SkipInstaller) {
    Write-Info '按 -SkipInstaller 跳过'
} elseif (-not $script:Iscc) {
    Write-Fail '未找到 ISCC.exe，无法生成本次构建要求的安装包。'
} else {
    if (Test-Path $ExePath) { Remove-Item $ExePath -Force }
    & $script:Iscc $Iss
    if ($LASTEXITCODE -ne 0) { Write-Fail ("ISCC 编译失败，退出码 {0}" -f $LASTEXITCODE) }
    elseif (Test-Path $ExePath) {
        $exeBuilt = $true
        Write-OK ("EXE 已生成: {0} ({1:N0} bytes)" -f $ExePath, (Get-Item $ExePath).Length)
    } else { Write-Fail 'ISCC 返回成功但未找到输出 EXE' }
}

# ============================================================================
# 5. COM 自动化回归测试
# ============================================================================
Write-Step '5/6 Excel COM 自动化回归测试'

function Invoke-AppStateTests($xl) {
    $initial = @{
        Calculation = [int]$xl.Calculation
        EnableEvents = [bool]$xl.EnableEvents
        ScreenUpdating = [bool]$xl.ScreenUpdating
    }
    $states = @(
        @{ Calculation = -4105; EnableEvents = $true;  ScreenUpdating = $true },
        @{ Calculation = -4135; EnableEvents = $false; ScreenUpdating = $false },
        @{ Calculation = -4105; EnableEvents = $false; ScreenUpdating = $true },
        @{ Calculation = -4135; EnableEvents = $true;  ScreenUpdating = $false }
    )
    try {
        foreach ($macro in @('Toolbox_DeleteEmptyRows', 'Toolbox_RemoveDuplicates')) {
            foreach ($scenario in @('正常完成', '无需处理', '工作表保护', '备份失败')) {
                foreach ($state in $states) {
                    $wb = $null
                    try {
                        $wb = $xl.Workbooks.Add()
                        $ws = $wb.Worksheets.Item(1)
                        $ws.Activate() | Out-Null
                        $sheetCount = [int]$wb.Worksheets.Count
                        $ws.Range('A1').Value2 = 'first'
                        $ws.Range('A3').Value2 = 'last'
                        if ($macro -eq 'Toolbox_RemoveDuplicates') { $ws.Range('A2').Value2 = 'first' }
                        if ($scenario -eq '无需处理') { $ws.Range('A2').Value2 = 'distinct' }
                        $ws.Range('A1:A3').Select() | Out-Null
                        if ($scenario -eq '工作表保护') { $ws.Protect('ExcelToolboxTests') }
                        if ($scenario -eq '备份失败') { $wb.Protect('ExcelToolboxTests', $true) }

                        $xl.Calculation = $state.Calculation
                        $xl.EnableEvents = $state.EnableEvents
                        $xl.ScreenUpdating = $state.ScreenUpdating
                        $xl.Run(("'ExcelToolbox.xlam'!{0}" -f $macro), $true) | Out-Null

                        $label = '{0}/{1}/Calc={2},Events={3},Screen={4}' -f $macro, $scenario,
                            $state.Calculation, $state.EnableEvents, $state.ScreenUpdating
                        $restored = ([int]$xl.Calculation -eq $state.Calculation) -and
                            ([bool]$xl.EnableEvents -eq $state.EnableEvents) -and
                            ([bool]$xl.ScreenUpdating -eq $state.ScreenUpdating)
                        Assert-True $restored ("状态恢复: {0}" -f $label)

                        $expectedRows = if ($scenario -eq '正常完成') { 2 } else { 3 }
                        $expectedBackups = if ($scenario -in @('正常完成', '工作表保护')) { 1 } else { 0 }
                        $outcome = ([int]$ws.UsedRange.Rows.Count -eq $expectedRows) -and
                            ([int]$wb.Worksheets.Count -eq ($sheetCount + $expectedBackups)) -and
                            ($ws.Range('A1').Value2 -eq 'first') -and
                            ($ws.Range(('A{0}' -f $expectedRows)).Value2 -eq 'last')
                        Assert-True $outcome ("操作结果与备份: {0}" -f $label)
                    } finally {
                        $xl.EnableEvents = $false
                        if ($wb) { $wb.Close($false) }
                    }
                }
            }
        }
    } finally {
        $xl.Calculation = $initial.Calculation
        $xl.EnableEvents = $initial.EnableEvents
        $xl.ScreenUpdating = $initial.ScreenUpdating
    }
}

function Invoke-ComTests {
    $xl = $null; $addin = $null; $temps = @()
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        # 仅对这个临时启动、随后立即退出的隐藏 Excel 实例放开宏，
        # 不写注册表、不影响用户 Excel 的信任中心设置（该属性不持久化）。
        $xl.AutomationSecurity = 1     # msoAutomationSecurityLow

        $addin = $xl.Workbooks.Open($XlamPath)
        $ping = $xl.Run("'ExcelToolbox.xlam'!Toolbox_Ping")
        Assert-True ($ping -eq '1.0.0') ("加载项可被 COM 加载，Toolbox_Ping = {0}" -f $ping)

        # ---- 删除空行 ----
        $t1 = Join-Path $env:TEMP ("ExcelToolbox_T1_{0}.xlsx" -f [guid]::NewGuid().ToString('N').Substring(0, 8))
        $temps += $t1; Copy-Item $TestXlsx $t1
        $wb = $xl.Workbooks.Open($t1)
        try {
            $ws = $wb.Worksheets.Item('删除空行测试'); $ws.Activate() | Out-Null
            $before = [int]$ws.UsedRange.Rows.Count
            $xl.Run("'ExcelToolbox.xlam'!Toolbox_DeleteEmptyRows", $true) | Out-Null
            $after = [int]$ws.UsedRange.Rows.Count
            $baks = 0; foreach ($s in $wb.Worksheets) { if ($s.Name -like '备份_*') { $baks++ } }
            Assert-True (($before -eq 10) -and ($after -eq 7)) ("删除空行: {0} -> {1} 行（期望 10 -> 7）" -f $before, $after)
            Assert-True ($baks -eq 1) ("删除空行: 自动创建备份工作表 {0} 张（期望 1）" -f $baks)
        } finally { $wb.Close($false) }

        # ---- 一键去重 ----
        $t2 = Join-Path $env:TEMP ("ExcelToolbox_T2_{0}.xlsx" -f [guid]::NewGuid().ToString('N').Substring(0, 8))
        $temps += $t2; Copy-Item $TestXlsx $t2
        $wb = $xl.Workbooks.Open($t2)
        try {
            $ws = $wb.Worksheets.Item('去重测试'); $ws.Activate() | Out-Null
            $ws.Range('A1:C8').Select() | Out-Null
            $before = [int]$ws.UsedRange.Rows.Count
            $xl.Run("'ExcelToolbox.xlam'!Toolbox_RemoveDuplicates", $true) | Out-Null
            $after = [int]$ws.UsedRange.Rows.Count
            $first = @($ws.Range('A1').Value2, $ws.Range('B1').Value2, $ws.Range('C1').Value2) -join '/'
            $baks = 0; foreach ($s in $wb.Worksheets) { if ($s.Name -like '备份_*') { $baks++ } }
            Assert-True (($before -eq 8) -and ($after -eq 4)) ("一键去重: {0} -> {1} 行（期望 8 -> 4）" -f $before, $after)
            Assert-True ($first -eq '张三/北京/100') ("一键去重: 保留第一条 = {0}" -f $first)
            Assert-True ($baks -eq 1) ("一键去重: 自动创建备份工作表 {0} 张（期望 1）" -f $baks)
        } finally { $wb.Close($false) }

        # ---- 美化表头 ----
        $t3 = Join-Path $env:TEMP ("ExcelToolbox_T3_{0}.xlsx" -f [guid]::NewGuid().ToString('N').Substring(0, 8))
        $temps += $t3; Copy-Item $TestXlsx $t3
        $wb = $xl.Workbooks.Open($t3)
        try {
            $ws = $wb.Worksheets.Item('美化表头测试'); $ws.Activate() | Out-Null
            $ws.Range('A1:E6').Select() | Out-Null
            $xl.Run("'ExcelToolbox.xlam'!Toolbox_BeautifyHeader", $true) | Out-Null
            $cell = $ws.Range('A1')
            $bg   = [int]$cell.Interior.Color       # RGB(0,112,192) = 12611584
            $fg   = [int]$cell.Font.Color           # RGB(255,255,255) = 16777215
            $bold = [bool]$cell.Font.Bold
            $baks = 0; foreach ($s in $wb.Worksheets) { if ($s.Name -like '备份_*') { $baks++ } }
            Assert-True ($bg -eq 12611584) ("美化表头: 背景色 = {0}（期望 12611584 蓝色）" -f $bg)
            Assert-True ($fg -eq 16777215) ("美化表头: 字体色 = {0}（期望 16777215 白色）" -f $fg)
            Assert-True $bold             '美化表头: 字体已加粗'
            Assert-True ($baks -eq 1) ("美化表头: 自动创建备份工作表 {0} 张（期望 1）" -f $baks)
        } finally { $wb.Close($false) }

        # 两个清理宏：四种原始状态，覆盖成功、无需处理、删除出错及备份失败。
        Invoke-AppStateTests $xl
    } finally {
        if ($addin) { try { $addin.Close($false) } catch { } }
        if ($xl) { try { $xl.Quit() } catch { } }
        foreach ($t in $temps) { Remove-Item $t -Force -ErrorAction SilentlyContinue }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
if ($SkipTests) { Write-Info '按 -SkipTests 跳过' }
else {
    try { Invoke-ComTests }
    catch { Write-Fail ("COM 测试异常：{0}" -f $_.Exception.Message) }
}

# ============================================================================
# 6. 汇总
# ============================================================================
Write-Step '6/6 构建结果'
$artifacts = @($XlamPath, $TestXlsx, $Iss)
if ($exeBuilt) { $artifacts += $ExePath }
if (-not $SkipInstaller -and -not $exeBuilt -and $script:Failed -eq 0) {
    Write-Fail '本次构建要求生成安装包，但未生成 EXE。'
}
foreach ($f in $artifacts) {
    if (Test-Path $f) {
        $i = Get-Item $f
        $hash = (Get-FileHash $f -Algorithm SHA256).Hash.Substring(0, 16)
        Write-Host ("  {0,-28} {1,10:N0} B   sha256:{2}..." -f $i.Name, $i.Length, $hash)
    }
}
Write-Host ''
if ($script:Failed -gt 0) {
    Write-Host ("构建完成，但有 {0} 项校验未通过，请查看上面的 [FAIL]。" -f $script:Failed) -ForegroundColor Red
    exit 1
}
Write-Host '全部校验通过。' -ForegroundColor Green
if ($SkipInstaller) { Write-Host '已按 -SkipInstaller 完成 XLAM 构建，本次未生成 EXE。' -ForegroundColor Yellow }
exit 0
