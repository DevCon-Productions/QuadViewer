; DevCon QuadViewer Installer Script
; Built with Inno Setup 6

[Setup]
AppName=DevCon QuadViewer
AppVersion=1.1
AppPublisher=DevCon Productions
AppPublisherURL=https://github.com/DevConProductions
AppCopyright=Copyright (C) 2026 DevCon Productions
DefaultDirName={autopf}\DevCon QuadViewer
DefaultGroupName=DevCon QuadViewer
UninstallDisplayIcon={app}\DevConQuadViewer.exe
OutputDir=installer_output
OutputBaseFilename=DevConQuadViewerSetup
SetupIconFile=quadviewer.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
; Main executable and all PyInstaller output
Source: "dist\DevConQuadViewer\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\DevCon QuadViewer"; Filename: "{app}\DevConQuadViewer.exe"; IconFilename: "{app}\_internal\quadviewer.ico"
Name: "{group}\Uninstall DevCon QuadViewer"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DevCon QuadViewer"; Filename: "{app}\DevConQuadViewer.exe"; IconFilename: "{app}\_internal\quadviewer.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\DevConQuadViewer.exe"; Description: "Launch DevCon QuadViewer"; Flags: nowait postinstall skipifsilent
