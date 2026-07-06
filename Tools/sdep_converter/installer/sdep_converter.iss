#define MyAppName "sdep_converter"
#define MyAppVersion "0.1.0"
#define MyAppExeName "sdep_converter.exe"

[Setup]
AppId={{ee97b988-65e9-45d6-89ad-a09b55383ccb}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Solar Design Engineering Package
DefaultDirName={autopf}\SDEP Converter
DefaultGroupName=SDEP Converter
OutputDir=..\target\installer
OutputBaseFilename=sdep_converter_setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\target\release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\schemas\*"; DestDir: "{app}\schemas"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\src\Data\*"; DestDir: "{app}\Data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
