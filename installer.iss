; Bill Scrapper - Inno Setup Installer Script
; Produces a modern, professional, per-user Windows Setup Wizard

#ifndef AppVersion
#define AppVersion "1.3.0"
#endif

#define AppName "Bill Scrapper"
#define AppPublisher "TN Developers Co."
#define AppExeName "BillScrapper.exe"

[Setup]
AppId={{D37F2C5A-8E9B-4D1C-A63E-5B9C8E7D1A2F}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=mailto:tahir.naqash4@gmail.com
AppSupportURL=https://wa.me/923339962474
DefaultDirName={localappdata}\Programs\BillScrapper
DisableDirPage=yes
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=BillScrapper-Setup
SetupIconFile=assets\billscrapper.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
CloseApplications=force
RestartApplications=no
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}"
Filename: "{app}\{#AppExeName}"; Flags: nowait shellexec; Check: ShouldLaunchSilent; WorkingDir: "{app}"

[Code]
function SetEnvironmentVariable(lpName, lpValue: string): BOOL;
external 'SetEnvironmentVariableW@kernel32.dll stdcall';

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // Clear any inherited PyInstaller variables so the relaunched app unpacks cleanly
  SetEnvironmentVariable('_MEIPASS2', '');
  SetEnvironmentVariable('_MEIPASS', '');
  SetEnvironmentVariable('PYTHONHOME', '');
  SetEnvironmentVariable('PYTHONPATH', '');

  // Terminate any running or lingering instances of BillScrapper.exe
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM BillScrapper.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(400);
  Result := True;
end;

function ShouldLaunchSilent: Boolean;
begin
  Result := WizardSilent or (ExpandConstant('{param:AUTOUPDATE|0}') = '1');
end;

