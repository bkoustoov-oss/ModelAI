#define ProductName "LAYA Agentic Studio"
#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

[Setup]
AppId={{D52B1989-6A31-4A3A-84D0-0B4A7029208A}
AppName={#ProductName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\LAYA Agentic Studio
DefaultGroupName={#ProductName}
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=LAYA-Windows-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\LAYAAgenticStudio.exe

[Files]
Source: "..\dist\LAYAAgenticStudio\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userprograms}\{#ProductName}"; Filename: "{app}\LAYAAgenticStudio.exe"
Name: "{userdesktop}\{#ProductName}"; Filename: "{app}\LAYAAgenticStudio.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\LAYAAgenticStudio.exe"; Description: "Launch {#ProductName}"; Flags: postinstall nowait skipifsilent
