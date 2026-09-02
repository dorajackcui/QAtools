#ifndef AppVersion
  #error AppVersion must be supplied by the build script
#endif
#ifndef SourceDir
  #error SourceDir must be supplied by the build script
#endif

[Setup]
AppId={{9A854BDD-9184-4B8D-9622-130C28E39182}
AppName=QAtools
AppVersion={#AppVersion}
AppVerName=QAtools {#AppVersion}
AppPublisher=QAtools
DefaultDirName={localappdata}\Programs\QAtools
DefaultGroupName=QAtools
PrivilegesRequired=lowest
UsePreviousAppDir=yes
DisableProgramGroupPage=yes
OutputBaseFilename=QAtools-Setup
SetupIconFile=QAtools.ico
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\QAtools.exe
WizardStyle=modern
MinVersion=10.0

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\QAtools"; Filename: "{app}\QAtools.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\QAtools"; Filename: "{app}\QAtools.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\QAtools.exe"; Description: "启动 QAtools"; Flags: nowait postinstall skipifsilent
