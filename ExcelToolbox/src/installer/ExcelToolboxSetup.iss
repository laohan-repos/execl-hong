; ============================================================================
;  Excel 小工具箱  --  Inno Setup 6 安装脚本
;  编译命令:  ISCC.exe ExcelToolboxSetup.iss
;  或双击项目根目录的 Build.cmd（会自动完成 XLAM 生成 + EXE 编译 + 测试）
;
;  设计要点
;   * 纯当前用户安装（PrivilegesRequired=lowest），不弹 UAC，不需要管理员。
;   * 不关闭、不降低任何全局宏安全设置。仅把工具箱自身目录登记为
;     Excel「受信任位置」（微软推荐的加载项部署方式，见 README 安全说明）。
;   * 自动检测 Office Excel 版本 / 位数 / 安装路径，检测不到则拒绝安装。
;   * 通过 OPEN<n> + Add-in Manager 注册加载项，不覆盖用户已有的 OPEN 项。
;   * 卸载时删除注册表项、受信任位置与全部文件。
; ============================================================================

#define AppName   "Excel 小工具箱"
#define AppVer    "1.0.0"
#define AddinFile "ExcelToolbox.xlam"

[Setup]
AppId={{A7E1C4D2-9B83-4F15-8D60-3C2E5F1A7B94}
AppName={#AppName}
AppVersion={#AppVer}
AppVerName={#AppName} {#AppVer}
AppPublisher=ExcelToolbox
DefaultDirName={userappdata}\Microsoft\AddIns\ExcelToolbox
DefaultGroupName={#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
#ifdef BuildOnInstall
InfoBeforeFile=InfoBeforeSource.txt
#else
InfoBeforeFile=InfoBefore.txt
#endif
OutputDir=..\..\dist
OutputBaseFilename=ExcelToolboxSetup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
MinVersion=10.0
UninstallDisplayName={#AppName} {#AppVer}
UninstallDisplayIcon={uninstallexe}
CloseApplications=no
; 若你的 Inno Setup 自带中文语言文件，可取消下面三行注释获得中文向导：
; [Languages]
; Name: "cn"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
#ifdef BuildOnInstall
; 安装前先在临时目录编译、回归；成功后才开始安装和写入注册表。
Source: "..\..\build\Build-ExcelToolbox.ps1"; DestDir: "{tmp}\ExcelToolbox\build"; Flags: dontcopy
Source: "..\vba\*.bas"; DestDir: "{tmp}\ExcelToolbox\src\vba"; Flags: dontcopy
Source: "..\vba\ThisWorkbook.cls"; DestDir: "{tmp}\ExcelToolbox\src\vba"; Flags: dontcopy
Source: "..\ribbon\customUI.xml"; DestDir: "{tmp}\ExcelToolbox\src\ribbon"; Flags: dontcopy
Source: "..\..\tests\ExcelToolbox_TestData.xlsx"; DestDir: "{tmp}\ExcelToolbox\tests"; Flags: dontcopy
Source: "{tmp}\ExcelToolbox\dist\{#AddinFile}"; DestDir: "{app}"; Flags: external ignoreversion
#else
Source: "..\..\dist\{#AddinFile}";          DestDir: "{app}"; Flags: ignoreversion
#endif
Source: "UserGuide.txt"; DestDir: "{app}"; DestName: "README.txt"; Flags: ignoreversion
Source: "..\..\build\Register-ExcelAddin.vbs"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式（打开使用说明）"; GroupDescription: "附加选项:"; Flags: unchecked

[Icons]
Name: "{group}\{#AppName} 使用说明"; Filename: "{app}\README.txt"
Name: "{group}\重新注册加载项（菜单未出现时）"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\Register-ExcelAddin.vbs"""
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName} 使用说明"; Filename: "{app}\README.txt"; Tasks: desktopicon

[Registry]
; ---- 1) 让 Excel 启动时自动打开该 XLAM -------------------------------------
; 不覆盖已有的 OPEN / OPEN1 / OPEN2 ...（可能属于其它加载项）
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVer}\Excel\Options"; \
  ValueType: string; ValueName: "{code:GetFreeOpenName}"; \
  ValueData: """{app}\{#AddinFile}"""; Flags: uninsdeletevalue

; ---- 2) 在「文件 > 选项 > 加载项 > Excel 加载项」列表中登记 -----------------
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVer}\Excel\Add-in Manager"; \
  ValueType: string; ValueName: "{app}\{#AddinFile}"; ValueData: ""; Flags: uninsdeletevalue

; ---- 3) 受信任位置：只信任本工具箱目录，不动全局宏安全级别 -----------------
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVer}\Excel\Security\Trusted Locations\LocationExcelToolbox"; \
  ValueType: string; ValueName: "Path"; ValueData: "{app}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVer}\Excel\Security\Trusted Locations\LocationExcelToolbox"; \
  ValueType: dword; ValueName: "AllowSubfolders"; ValueData: 0
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVer}\Excel\Security\Trusted Locations\LocationExcelToolbox"; \
  ValueType: string; ValueName: "Description"; ValueData: "Excel 小工具箱 安装目录"

; ============================================================================
;  安装逻辑（Pascal Script）
; ============================================================================
[Code]
var
  gOfficeVer : String;
  gOfficeBit : String;
  gExcelExe  : String;
  gOpenName  : String;
#ifdef BuildOnInstall
  gBuildReady: Boolean;
#endif

{ ---------- Office 版本 / 位数 / 路径检测 ---------- }

function IsExcelPath(Path: String): Boolean;
var
  VersionMS, VersionLS: Cardinal;
  Major: Integer;
begin
  Result := False;
  if (CompareText(ExtractFileName(Path), 'EXCEL.EXE') <> 0) or
     not FileExists(Path) or
     (Pos('kingsoft', Lowercase(Path)) > 0) or
     (Pos('wps', Lowercase(Path)) > 0) then Exit;
  if not GetVersionNumbers(Path, VersionMS, VersionLS) then Exit;
  Major := VersionMS shr 16;
  if Major < 14 then Exit;
  gExcelExe := Path;
  gOfficeVer := IntToStr(Major) + '.0';
  Result := True;
end;

function FindExcelInView(RootKey: Integer): Boolean;
var
  i: Integer;
  Path: String;
  Versions: array[0..2] of String;
begin
  Result := False;
  if RegQueryStringValue(RootKey,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', Path) then
  begin
    Path := RemoveQuotes(Trim(Path));
    if IsExcelPath(Path) then
    begin
      Result := True;
      Exit;
    end;
  end;
  Versions[0] := '16.0'; Versions[1] := '15.0'; Versions[2] := '14.0';
  for i := 0 to 2 do
  begin
    if RegQueryStringValue(RootKey,
      'SOFTWARE\Microsoft\Office\' + Versions[i] + '\Excel\InstallRoot', 'Path', Path) then
    begin
      if IsExcelPath(AddBackslash(RemoveQuotes(Trim(Path))) + 'EXCEL.EXE') then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

function DetectExcel(): Boolean;
begin
  gOfficeVer := '';
  gOfficeBit := '';
  gExcelExe  := '';

  { 分别检查两个注册表视图，避免把 WPS 的 Excel.Application 别名当成 Office。 }
  Result := FindExcelInView(HKLM32) or FindExcelInView(HKCU32);
  if Result then gOfficeBit := '32-bit';
  if not Result and IsWin64 then
  begin
    Result := FindExcelInView(HKLM64) or FindExcelInView(HKCU64);
    if Result then gOfficeBit := '64-bit';
  end;
  if Result then
  begin
    Result := RegKeyExists(HKCR32, 'Excel.Application');
    if not Result and IsWin64 then Result := RegKeyExists(HKCR64, 'Excel.Application');
  end;
end;

// ---------- 供 [Registry] 使用的 code 常量函数 ----------

function GetOfficeVer(Param: String): String;
var
  dummy: Boolean;
begin
  if gOfficeVer = '' then dummy := DetectExcel();
  Result := gOfficeVer;
end;

{ 找第一个空闲的 OPEN / OPEN1 / OPEN2 ...，并缓存，保证安装与卸载一致 }
function GetFreeOpenName(Param: String): String;
var
  i: Integer;
  n, key: String;
  dummy: Boolean;
begin
  if gOpenName <> '' then
  begin
    Result := gOpenName;
    Exit;
  end;
  if gOfficeVer = '' then dummy := DetectExcel();
  key := 'Software\Microsoft\Office\' + gOfficeVer + '\Excel\Options';
  i := 0;
  while True do
  begin
    if i = 0 then n := 'OPEN' else n := 'OPEN' + IntToStr(i);
    if not RegValueExists(HKCU, key, n) then Break;
    i := i + 1;
    if i > 200 then RaiseException('Excel 的 OPEN 注册项已满，无法安全注册加载项。');
  end;
  gOpenName := n;
  Result := n;
end;

{ ---------- Excel 正在运行时要求先关闭 ---------- }

function AskCloseExcel(): Boolean;
var
  r: Integer;
begin
  Result := False;
  while FindWindowByClassName('XLMAIN') <> 0 do
  begin
    if WizardSilent then Exit;
    r := MsgBox('检测到 Microsoft Excel 正在运行。' #13#10 #13#10 +
                '请先保存好工作，并关闭所有 Excel 窗口，再继续。' #13#10 #13#10 +
                '关闭后点【确定】继续，点【取消】退出。',
                mbConfirmation, MB_OKCANCEL);
    if r <> IDOK then Exit;
  end;
  Result := True;
end;

function InitializeSetup(): Boolean;
var
  dummy: Boolean;
begin
  Result := False;
  dummy := DetectExcel();
  if not dummy then
  begin
    Log('Unsupported Office: ' + gOfficeVer + '; ' + gExcelExe);
    if not WizardSilent then
    MsgBox('未检测到受支持的 Microsoft Office Excel，安装程序将退出。' #13#10 #13#10 +
           '本工具需要桌面版 Microsoft Office Excel（Office 2010 / 14.0 及以上）。' #13#10 +
           '网页版 Excel、WPS、仅安装 Office 运行时均不受支持。' #13#10 #13#10 +
           '检测到的版本号 : ' + gOfficeVer + #13#10 +
           '检测到的位数   : ' + gOfficeBit + #13#10 +
           '检测到的 EXCEL : ' + gExcelExe,
           mbError, MB_OK);
    Exit;
  end;
  if not AskCloseExcel() then Exit;
  Result := True;
end;

#ifdef BuildOnInstall
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  BuildRoot, BuildScript, BuildLog, Params, OldCaption: String;
  ExitCode: Integer;
begin
  Result := '';
  if gBuildReady then Exit;
  if not AskCloseExcel() then
  begin
    Result := '请保存并关闭所有 Excel 窗口后重试。';
    Exit;
  end;
  OldCaption := WizardForm.PreparingLabel.Caption;
  WizardForm.PreparingLabel.Caption := '正在生成并验证 Excel 加载项，请稍候…';
  try
    ExtractTemporaryFiles('{tmp}\ExcelToolbox\*');
    BuildRoot := ExpandConstant('{tmp}\ExcelToolbox');
    BuildScript := BuildRoot + '\build\Build-ExcelToolbox.ps1';
    BuildLog := ExpandConstant('{localappdata}\ExcelToolbox\Logs\build-') +
      GetDateTimeString('yyyymmdd-hhnnss', '', '') + '.log';
    Params := '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' +
      BuildScript + '" -SkipInstaller -PowerShellRibbon -LogPath "' + BuildLog + '"';
    Log('Building add-in; transcript: ' + BuildLog);
    if not Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
                Params, BuildRoot, SW_HIDE, ewWaitUntilTerminated, ExitCode) then
      Result := '无法启动 Windows PowerShell。错误：' + SysErrorMessage(ExitCode)
    else if ExitCode <> 0 then
      Result := '加载项生成或测试失败，尚未安装。请查看日志：' + #13#10 + BuildLog
    else if not FileExists(BuildRoot + '\dist\{#AddinFile}') then
      Result := '未生成加载项，尚未安装。请查看日志：' + #13#10 + BuildLog
    else
      gBuildReady := True;
  except
    Result := '准备安装失败：' + GetExceptionMessage;
  end;
  WizardForm.PreparingLabel.Caption := OldCaption;
end;
#endif

function InitializeUninstall(): Boolean;
begin
  Result := AskCloseExcel();
end;

{ ---------- 卸载时清理可能由「重新注册加载项」写入的 OPEN 项 ---------- }

procedure RemoveStaleOpenValues();
var
  i: Integer;
  n, key, v: String;
  dummy: Boolean;
begin
  if gOfficeVer = '' then dummy := DetectExcel();
  if gOfficeVer = '' then Exit;
  key := 'Software\Microsoft\Office\' + gOfficeVer + '\Excel\Options';
  for i := 0 to 200 do
  begin
    if i = 0 then n := 'OPEN' else n := 'OPEN' + IntToStr(i);
    if RegQueryStringValue(HKCU, key, n, v) then
      if Pos('ExcelToolbox.xlam', v) > 0 then
        RegDeleteValue(HKCU, key, n);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then RemoveStaleOpenValues();
end;
