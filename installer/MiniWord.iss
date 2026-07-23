; MiniWord installer script (Inno Setup 6)
; Build: ISCC.exe MiniWord.iss

#define MyAppName "MiniWord"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Levitd"
#define MyAppURL "https://github.com/Levitd/MiniWord"
#define MyAppExeName "MiniWord.exe"
#define PublishDir "..\MiniWord\bin\Release\net6.0-windows\win-x64\publish"

[Setup]
AppId={{8F2A1C64-3D5B-4E9A-9C77-A1B2C3D4E5F6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=Output
OutputBaseFilename=MiniWordSetup-{#MyAppVersion}
SetupIconFile=..\MiniWord\MiniWord.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#PublishDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[CustomMessages]
english.DotNetMissing=MiniWord requires the .NET 6 Desktop Runtime (or newer), which was not found on this computer.%n%nOpen the download page now? You can continue the installation and install the runtime later.
russian.DotNetMissing=Для работы MiniWord требуется .NET 6 Desktop Runtime (или новее), который не найден на этом компьютере.%n%nОткрыть страницу загрузки? Установку можно продолжить, а рантайм поставить позже.

[Code]
function IsDotNetDesktopInstalled(): Boolean;
var
  SharedDir: string;
  FindRec: TFindRec;
  MajorStr: string;
  Major: Integer;
begin
  Result := False;
  SharedDir := ExpandConstant('{commonpf64}\dotnet\shared\Microsoft.WindowsDesktop.App');
  if not DirExists(SharedDir) then
    SharedDir := ExpandConstant('{commonpf32}\dotnet\shared\Microsoft.WindowsDesktop.App');
  if not DirExists(SharedDir) then
    Exit;

  if FindFirst(SharedDir + '\*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0)
          and (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          MajorStr := FindRec.Name;
          if Pos('.', MajorStr) > 0 then
            MajorStr := Copy(MajorStr, 1, Pos('.', MajorStr) - 1);
          Major := StrToIntDef(MajorStr, 0);
          if Major >= 6 then
          begin
            Result := True;
            Exit;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function InitializeSetup(): Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;
  if not IsDotNetDesktopInstalled() then
  begin
    if MsgBox(CustomMessage('DotNetMissing'), mbConfirmation, MB_YESNO) = IDYES then
      ShellExec('open', 'https://dotnet.microsoft.com/download/dotnet/6.0/runtime',
        '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  end;
end;
