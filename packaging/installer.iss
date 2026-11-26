; Compile from the repo root after building the exe:
;   iscc /DAppVersion=1.5 packaging\installer.iss
#ifndef AppVersion
  #define AppVersion "1.5"
#endif

[Setup]
AppName=Mail AI
AppVersion={#AppVersion}
AppPublisher=mailai.uk
AppPublisherURL=https://mailai.uk
DefaultDirName={autopf}\Mail AI
DefaultGroupName=Mail AI
OutputDir=..\installer_output
OutputBaseFilename=MailAI_Setup
SetupIconFile=..\app\assets\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "..\dist\Mail AI v{#AppVersion}.exe"; DestDir: "{app}"; DestName: "Mail AI.exe"; Flags: ignoreversion

[Icons]
Name: "{group}\Mail AI"; Filename: "{app}\Mail AI.exe"
Name: "{group}\Uninstall Mail AI"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Mail AI"; Filename: "{app}\Mail AI.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Mail AI.exe"; Description: "Launch Mail AI"; Flags: nowait postinstall skipifsilent