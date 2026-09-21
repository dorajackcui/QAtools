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
UninstallDisplayIcon={app}\QAtools-icon-{#AppVersion}.ico
WizardStyle=modern
MinVersion=10.0
ChangesAssociations=yes

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; Flags: unchecked
Name: "shellmenu"; Description: "添加 Excel 文件和文件夹右键菜单"

[Registry]
; Own only these verbs; leave Excel/WPS defaults and other applications untouched.
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.Workflow"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.Workflow"; ValueType: string; ValueName: ""; ValueData: "一键质量检查"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.Workflow"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.Workflow"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.Workflow\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --qa-workflow ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.Workflow"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.Workflow"; ValueType: string; ValueName: ""; ValueData: "一键质量检查"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.Workflow"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.Workflow"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.Workflow\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --qa-workflow ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.FrenchNbsp"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: ""; ValueData: "法语 NBSP 修复"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\QAtools.FrenchNbsp\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --nbsp-restore ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.FrenchNbsp"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: ""; ValueData: "法语 NBSP 修复"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.FrenchNbsp"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\QAtools.FrenchNbsp\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --nbsp-restore ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Compatibility"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Compatibility"; ValueType: string; ValueName: ""; ValueData: "兼容性重存"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Compatibility"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Compatibility"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Compatibility\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --compatibility-dir ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Merge"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Merge"; ValueType: string; ValueName: ""; ValueData: "合并表格"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Merge"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Merge"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Merge\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --merge-dir ""%1"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Untranslated"; ValueType: none; Flags: deletekey; Tasks: not shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Untranslated"; ValueType: string; ValueName: ""; ValueData: "统计未翻译"; Flags: uninsdeletekey; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Untranslated"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\QAtools-icon-{#AppVersion}.ico"""; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Untranslated"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"; Tasks: shellmenu
Root: HKCU; Subkey: "Software\Classes\Directory\shell\QAtools.Untranslated\command"; ValueType: string; ValueName: ""; ValueData: """{app}\QAtools.exe"" --untranslated-dir ""%1"""; Tasks: shellmenu

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "QAtools.ico"; DestDir: "{app}"; DestName: "QAtools-icon-{#AppVersion}.ico"; Flags: ignoreversion

[InstallDelete]
; Only obsolete files owned by QAtools; preserve user files and configuration.
Type: files; Name: "{app}\QAtools-CLI.exe"
Type: files; Name: "{app}\QAtools-CLI.cmd"
Type: filesandordirs; Name: "{app}\_internal\numpy"
Type: filesandordirs; Name: "{app}\_internal\numpy.libs"
Type: filesandordirs; Name: "{app}\_internal\numpy-*.dist-info"
Type: filesandordirs; Name: "{app}\_internal\Pythonwin"
Type: filesandordirs; Name: "{app}\_internal\yaml"

[Icons]
Name: "{group}\QAtools"; Filename: "{app}\QAtools.exe"; WorkingDir: "{app}"; IconFilename: "{app}\QAtools-icon-{#AppVersion}.ico"
Name: "{autodesktop}\QAtools"; Filename: "{app}\QAtools.exe"; WorkingDir: "{app}"; IconFilename: "{app}\QAtools-icon-{#AppVersion}.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\QAtools.exe"; Description: "启动 QAtools"; Flags: nowait postinstall skipifsilent
