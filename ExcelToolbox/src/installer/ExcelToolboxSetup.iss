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
InfoBeforeFile=InfoBefore.txt
OutputDir=..\..\dist
OutputBaseFilename=ExcelToolboxSetup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayName={#AppName} {#AppVer}
UninstallDisplayIcon={uninstallexe}
CloseApplications=no
; 若你的 Inno Setup 自带中文语言文件，可取消下面三行注释获得中文向导：
; [Languages]
; Name: "cn"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
Source: "..\..\dist\{#AddinFile}";          DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\README.md";                  DestDir: "{app}"; DestName: "README.txt"; Flags: ignoreversion
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

{ ---------- Office 版本 / 位数 / 路径检测 ---------- }

function VerFromCurVer(): String;
var
  s: String;
begin
  Result := '';
  if RegQueryStringValue(HKCR, 'Excel.Application\CurVer', '', s) then
  begin
    { 典型值: Excel.Application.16 }
    if Pos('.16', s) > 0 then Result := '16.0'
    else if Pos('.15', s) > 0 then Result := '15.0'
    else if Pos('.14', s) > 0 then Result := '14.0'
    else if Pos('.12', s) > 0 then Result := '12.0'
    else if Pos('.11', s) > 0 then Result := '11.0';
  end;
end;

function VerFromClickToRun(): String;
var
  s: String;
begin
  Result := '';
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
                         'VersionToReport', s) then
  begin
    { 典型值: 16.0.14326.20504 }
    if Copy(s, 1, 3) = '16.' then Result := '16.0'
    else if Copy(s, 1, 3) = '15.' then Result := '15.0'
    else if Copy(s, 1, 3) = '14.' then Result := '14.0';
  end;
end;

function VerFromInstallRoot(): String;
var
  i: Integer;
  p: String;
  v: array[0..4] of String;
begin
  Result := '';
  v[0] := '16.0'; v[1] := '15.0'; v[2] := '14.0'; v[3] := '12.0'; v[4] := '11.0';
  for i := 0 to 4 do
  begin
    if RegQueryStringValue(HKLM,
      'SOFTWARE\Microsoft\Office\' + v[i] + '\Excel\InstallRoot', 'Path', p) then
    begin
      if (gExcelExe = '') and FileExists(p + 'EXCEL.EXE') then
        gExcelExe := p + 'EXCEL.EXE';
      if Result = '' then Result := v[i];
    end;
    { 64 位系统上的 32 位 Office }
    if RegQueryStringValue(HKLM,
      'SOFTWARE\WOW6432Node\Microsoft\Office\' + v[i] + '\Excel\InstallRoot', 'Path', p) then
    begin
      if (gExcelExe = '') and FileExists(p + 'EXCEL.EXE') then
        gExcelExe := p + 'EXCEL.EXE';
      if gOfficeBit = '' then gOfficeBit := '32-bit';
      if Result = '' then Result := v[i];
    end;
  end;
end;

function DetectExcel(): Boolean;
var
  s: String;
begin
  gOfficeVer := '';
  gOfficeBit := '';
  gExcelExe  := '';

  { EXCEL.EXE 路径 }
  if RegQueryStringValue(HKLM,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', s) then
    if FileExists(s) then gExcelExe := s;
  if (gExcelExe = '') and RegQueryStringValue(HKCU,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', s) then
    if FileExists(s) then gExcelExe := s;

  { 版本号，按可靠度依次尝试 }
  gOfficeVer := VerFromCurVer();
  if gOfficeVer = '' then gOfficeVer := VerFromClickToRun();
  if gOfficeVer = '' then gOfficeVer := VerFromInstallRoot();

  { 位数 }
  if gOfficeBit = '' then
    if RegQueryStringValue(HKLM,
      'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', s) then
    begin
      if s = 'x86' then gOfficeBit := '32-bit'
      else if s = 'x64' then gOfficeBit := '64-bit';
    end;
  if gOfficeBit = '' then
  begin
    if Pos('(x86)', gExcelExe) > 0 then gOfficeBit := '32-bit'
    else if Pos('Program Files', gExcelExe) > 0 then gOfficeBit := '64-bit'
    else gOfficeBit := 'unknown';
  end;

  { 必须能找到 Excel 的 COM 注册，否则加载项一定无法工作 }
  Result := (gOfficeVer <> '') and RegKeyExists(HKCR, 'Excel.Application');
end;

{ ---------- 供 [Registry] 使用的 {code:...} 函数 ---------- }

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
    if i > 200 then Break;
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
    MsgBox('未检测到 Microsoft Office Excel，安装程序将退出。' #13#10 #13#10 +
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
