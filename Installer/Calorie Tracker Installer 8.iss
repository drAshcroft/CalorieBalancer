; InnoScript Version 7.2  Build 4  Trial
; Randem Systems, Inc.
; Copyright 2003-2007
; Website:  http://www.randem.com
; Support:  http://www.randem.com/cgi-bin/discus/discus.cgi
; OS: Windows NT 6.0 build 6000 ()

; Date: October 16, 2007

;              VB Runtime Files Folder:   C:\Users\Public\shawntel\Randem Systems\InnoScript\InnoScript 7\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Users\shawntel\Desktop\Final Calories 2.9\Project1.vbp
; Inno Setup Script Output File (.iss):   C:\Users\shawntel\Desktop\Final Calories 2.9\Installer\Calorie Tracker Installer 5 Release.iss
;         Script Template Files (.tpl):   C:\Users\Public\shawntel\Randem Systems\InnoScript\InnoScript 7\Templates\Release.tpl
;                                     :   C:\Users\Public\shawntel\Randem Systems\InnoScript\InnoScript 7\Templates\VBRuntime.tpl
;                                     :   C:\Users\Public\shawntel\Randem Systems\InnoScript\InnoScript 7\Templates\Vista.tpl

; ------------------------
;        References
; ------------------------

; Visual Basic runtime objects and procedures - (MSVBVM60.DLL)
; OLE Automation - (STDOLE2.TLB)
; Microsoft DAO 3.6 Object Library - (dao360.dll)
; Microsoft Internet Controls - (ieframe.dll)
; Microsoft HTML Object Library - (mshtml.tlb)
; Microsoft Shell Controls And Automation - (SHELL32.dll)


; --------------------------
;        Components
; --------------------------

; Microsoft Script Control 1.0 - (msscript.ocx)
; Microsoft Common Dialog Control 6.0 (SP6) - (COMDLG32.OCX)
; Microsoft Windows Common Controls 6.0 (SP6) - (MSCOMCTL.OCX)
; Microsoft FlexGrid Control 6.0 (SP6) - (MSFLXGRD.OCX)
; Microsoft Rich Textbox Control 6.0 (SP6) - (RICHTX32.OCX)
; Microsoft Windows Common Controls-2 6.0 (SP6) - (MSCOMCT2.OCX)


[Setup]
AppName=Calorie Balance Tracker
AppVerName=Calorie Balance Tracker 4.0.5
AppPublisher=Calorie Balance Diet
AppPublisherURL=http://www.CalorieBalanceDiet.com
AppSupportURL=http://www.CalorieBalanceDiet.com
AppUpdatesURL=http://www.CalorieBalanceDiet.com
DefaultDirName={code:GetAppFolder}\Calorie Balance\Calorie Balance Tracker
DefaultGroupName=Calorie Balance Tracker
OutputDir=C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer
licensefile=C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\EULA.txt
OutputBaseFilename=Setup Calorie Balance Tracker
SetupIconFile=C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Icons\Food.ico
Compression=lzma
SolidCompression=yes

AppId=CalorieBalanceTracker

AppVersion=4.0.5
VersionInfoVersion=4.0.5
AllowNoIcons=no
MinVersion=4.0,4.0
privilegesRequired=none

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: AutoOSUpdater; Description: Install MDAC's for Database Operations; GroupDescription: Install MDAC's:
Name: ScriptingRuntime; Description: Install Microsoft's Scripting Runtime; GroupDescription: Install Scripting Runtime:

[Files]
; begin VB system files
; (Note: Scroll to the right to see the full lines!)
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\msvbvm60.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\oleaut32.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\olepro32.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\asycfilt.dll; DestDir: {sys}; Flags:  sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\stdole2.tlb; DestDir: {sys}; Flags:  uninsneveruninstall sharedfile regtypelib onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\comcat.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6

Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\CheckSystem\WebUpdater.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion; Tasks: AutoOSUpdater
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\DhtmlEd.msi; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion; MinVersion: 0,6.0; Tasks: AutoOSUpdater

Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\scripten.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion nocompression; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; Tasks: ScriptingRuntime
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\scr56en.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion nocompression; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; Tasks: ScriptingRuntime

;Source: c:\program files\common files\microsoft shared\dao\dao360.dll; DestDir: {cf}\microsoft shared\dao\; Flags:  regserver restartreplace sharedfile;
;Source: msscript.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: "C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Calorie Balance Tracker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\EULA.txt"; DestDir: "{app}";
Source: "C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\New Folder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs


Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\COMDLG32.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\MSCOMCT2.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\MSFLXGRD.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\RICHTX32.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Administrator\Desktop\Free Calories 4.0\Installer\System32\MSCOMCTL.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;


[INI]
Filename: {app}\Calorie Balance Tracker.url; Section: InternetShortcut; Key: URL; String: http://www.caloriebalancediet.com

[Icons]
Name: {group}\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; WorkingDir: {app}
Name: {group}\{cm:ProgramOnTheWeb, CalorieTracker}; Filename: {app}\Calorie Balance Tracker.url
Name: {group}\{cm:UninstallProgram, CalorieTracker}; Filename: {uninstallexe}
Name: {commondesktop}\Calorie Tracker; Filename: {app}\Calorie Balance Tracker.exe; Tasks: desktopicon; WorkingDir: {app}
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; Tasks: quicklaunchicon; WorkingDir: {app};

[Run]
Filename: "msiexec.exe"; Parameters: "/i ""{tmp}\DhtmlEd.msi"" /passive /norestart"  ;MinVersion: 0,6; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: AutoOSUpdater
Filename: {tmp}\WebUpdater.exe; Parameters: {app}; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: AutoOSUpdater
;Filename: {tmp}\VB_DCOM_MDAC_JET_AutoSetup.exe; Parameters: /NORESTART /SILENT; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: AutoOSUpdater
Filename: {tmp}\scr56en.exe; Parameters: /r:n ; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {tmp}\scripten.exe; Parameters: /r:n ; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {app}\Calorie Balance Tracker.exe; Description: {cm:LaunchProgram, CalorieTracker}; Flags: nowait postinstall skipifsilent; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\Calorie Balance Tracker.url

[Registry]
;Root: HKCU; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: RUNASADMIN; Flags: uninsdeletevalue
Root: HKLM; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: WINXPSP2; Flags: uninsdeletevalue
Root: HKCR; Subkey: ".cbm"; ValueType: string; ValueName: ""; ValueData: "caloriebalancetracker"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "caloriebalancetracker"; ValueType: string; ValueName: ""; ValueData: "Calorie Balance Tracker"; Flags: uninsdeletekey
Root: HKCR; Subkey: "caloriebalancetracker\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\calorie balance tracker.EXE,0"
Root: HKCR; Subkey: "caloriebalancetracker\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\calorie balance tracker.EXE"" ""%1"""

[Code]
function GetAppFolder(Param: String): String;
begin
  if InstallOnThisVersion('0,6', '0,0') = irInstall then
    Result := 'C:\Users\Public\'
  else
    Result := ExpandConstant('{pf}');
end;

[UninstallRun]
Filename: "{app}\Calorie Balance Tracker.exe"; Parameters: "/ExitSurvey"; RunOnceId: ExitSurvey


