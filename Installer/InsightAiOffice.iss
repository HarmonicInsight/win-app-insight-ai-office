; Insight AI Office Inno Setup Installer Script
; Requires Inno Setup 6.x - https://jrsoftware.org/isinfo.php
;
; Build with:  ISCC.exe InsightAiOffice.iss

#define MyAppName "Insight AI Office"
#define MyAppVersion "1.0.3"
#define MyAppPublisher "HARMONIC insight"
#define MyAppExeName "InsightAiOffice.exe"
#define MyAppURL "https://github.com/HarmonicInsight/win-app-insight-ai-office"
#define PublishDir "..\publish"

[Setup]
AppId={{D0E1F2A3-4B5C-6D7E-8F9A-B0C1D2E3F4A5}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\Output
OutputBaseFilename=InsightAiOffice_Setup_{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupLogging=yes
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
Root: HKLM; Subkey: "SOFTWARE\HarmonicInsight\Insight AI Office"; ValueName: "InstallPath"; ValueType: string; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\HarmonicInsight\Insight AI Office"; ValueName: "Version"; ValueType: string; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\HarmonicInsight\IAOF"
Type: filesandordirs; Name: "{userappdata}\HarmonicInsight\IAOF"
