<#
.SYNOPSIS
    无需本机安装 Excel，生成在目标电脑编译加载项的 EXE 安装包。
#>
[CmdletBinding()]
param([string] $IsccPath = '')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$candidates = @(
    $IsccPath,
    (Join-Path $PSScriptRoot 'sfx\InnoSetup\ISCC.exe'),
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe',
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe')
)
$command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($command) { $candidates += $command.Source }
$compiler = $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
if (-not $compiler) { throw '未找到 Inno Setup 6。请用 -IsccPath 指定 ISCC.exe。' }
$exe = Join-Path $root 'dist\ExcelToolboxSetup.exe'
New-Item -ItemType Directory -Path (Split-Path -Parent $exe) -Force | Out-Null
& $compiler '/Qp' '/DBuildOnInstall' (Join-Path $root 'src\installer\ExcelToolboxSetup.iss')
if ($LASTEXITCODE -ne 0) { throw "Inno Setup 编译失败，退出码 $LASTEXITCODE" }
if (-not (Test-Path -LiteralPath $exe)) { throw '编译器未输出 EXE。' }
Get-Item -LiteralPath $exe | Select-Object FullName, Length
Write-Output ('SHA256: ' + (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash)
Write-Host 'EXE 已生成。目标电脑需要 Microsoft Excel 2010+；安装时自动编译并测试加载项。'
