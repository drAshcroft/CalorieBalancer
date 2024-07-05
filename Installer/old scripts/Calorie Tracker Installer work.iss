; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=Calorie Tracker
AppVerName=Calorie Tracker 1.0
AppPublisher=Calorie Balance Diet
AppPublisherURL=http://www.CalorieBalanceDiet.com
AppSupportURL=http://www.CalorieBalanceDiet.com
AppUpdatesURL=http://www.CalorieBalanceDiet.com
DefaultDirName={pf}\Calorie Tracker
DefaultGroupName=Calorie Tracker
OutputDir=C:\Documents and Settings\bashc\Desktop\Final Calories\Installer
licensefile=C:\Documents and Settings\bashc\Desktop\Final Calories\EULA.txt
OutputBaseFilename=Setup Calorie Tracker
SetupIconFile=C:\Documents and Settings\bashc\Desktop\Final Calories\Icons\Food.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; begin VB system files
; (Note: Scroll to the right to see the full lines!)
Source: "C:\Program Files\Inno Setup 5\vbfiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "C:\Program Files\Inno Setup 5\vbfiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Program Files\Inno Setup 5\vbfiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Program Files\Inno Setup 5\vbfiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Program Files\Inno Setup 5\vbfiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Program Files\Inno Setup 5\vbfiles\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
; end VB system files
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Calorie Tracker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\EULA.txt"; DestDir: "{app}";
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\New Folder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\COMDLG32.OCX"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\dao360.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\MSFLXGRD.OCX"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\mshtml.tlb"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\msscript.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
Source: "C:\Documents and Settings\bashc\Desktop\Final Calories\Installer\System32\RICHTX32.OCX"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver onlyifdoesntexist
; NOTE: Don't use "Flags: ignoreversion" on any shared system files


[INI]
Filename: "{app}\Calorie Tracker.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.CalorieBalanceDiet.com"

[Icons]
Name: "{group}\Calorie Tracker"; Filename: "{app}\Calorie Tracker.exe"
Name: "{group}\{cm:ProgramOnTheWeb,Calorie Tracker}"; Filename: "{app}\Calorie Tracker.url"
Name: "{userdesktop}\Calorie Tracker"; Filename: "{app}\Calorie Tracker.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Calorie Tracker.exe"; Description: "{cm:LaunchProgram,Calorie Tracker}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\Calorie Tracker.url"

