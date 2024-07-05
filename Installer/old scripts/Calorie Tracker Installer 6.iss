; InnoScript Version 7.2  Build 4  Trial
; Randem Systems, Inc.
; Copyright 2003-2007
; Website:  http://www.randem.com
; Support:  http://www.randem.com/cgi-bin/discus/discus.cgi
; OS: Windows NT 6.0 build 6000 ()

; Date: October 16, 2007

;              VB Runtime Files Folder:   C:\Users\Public\shawntel\Randem Systems\InnoScript\InnoScript 7\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Users\shawntel\Desktop\Final Calories 2.5\Project1.vbp
; Inno Setup Script Output File (.iss):   C:\Users\shawntel\Desktop\Final Calories 2.5\Installer\Calorie Tracker Installer 5 Release.iss
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
AppVerName=Calorie Balance Tracker 2.8
AppPublisher=Calorie Balance Diet
AppPublisherURL=http://www.CalorieBalanceDiet.com
AppSupportURL=http://www.CalorieBalanceDiet.com
AppUpdatesURL=http://www.CalorieBalanceDiet.com
DefaultDirName={pf}\Calorie Balance Tracker
DefaultGroupName=Calorie Balance Tracker
OutputDir=C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer
licensefile=C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\EULA.txt
OutputBaseFilename=Setup Calorie Balance Tracker
SetupIconFile=C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Icons\Food.ico
Compression=lzma
SolidCompression=yes

AppId=CalorieBalanceTracker

AppVersion=2.8.0
VersionInfoVersion=2.8.0
AllowNoIcons=no
MinVersion=4.0,4.0
privilegesRequired=none

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: AutoOSUpdater; Description: Install MDAC's for Database Operations; GroupDescription: Install MDAC's:
Name: ScriptingRuntime; Description: Install Microsoft's Scripting Runtime; GroupDescription: Install Scripting Runtime:

[Files]
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\vb_dcom_mdac_jet_autosetup.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion nocompression; Tasks: AutoOSUpdater
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\scripten.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion nocompression; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; Tasks: ScriptingRuntime
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\scr56en.exe; DestDir: {tmp}; Flags:  deleteafterinstall ignoreversion nocompression; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; Tasks: ScriptingRuntime
;Source: c:\program files\common files\microsoft shared\dao\dao360.dll; DestDir: {cf}\microsoft shared\dao\; Flags:  regserver restartreplace sharedfile;
;Source: msscript.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: "C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Calorie Balance Tracker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\EULA.txt"; DestDir: "{app}";
Source: "C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\New Folder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs


Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\System32\COMDLG32.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\System32\MSCOMCT2.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\System32\MSFLXGRD.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\System32\RICHTX32.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\System32\MSCOMCTL.OCX; DestDir: {sys}; Flags:  regserver restartreplace sharedfile;


[INI]
Filename: {app}\Calorie Balance Tracker.url; Section: InternetShortcut; Key: URL; String: http://www.caloriebalancediet.com

[Icons]
Name: {group}\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; WorkingDir: {app}
Name: {group}\{cm:ProgramOnTheWeb, CalorieTracker}; Filename: {app}\Calorie Balance Tracker.url
Name: {group}\{cm:UninstallProgram, CalorieTracker}; Filename: {uninstallexe}
Name: {commondesktop}\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; Tasks: desktopicon; WorkingDir: {app}
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; Tasks: quicklaunchicon; WorkingDir: {app};

[Run]
Filename: {tmp}\VB_DCOM_MDAC_JET_AutoSetup.exe; Parameters: /NORESTART /VERYSILENT; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: AutoOSUpdater
Filename: {tmp}\scr56en.exe; Parameters: /r:n /q:1; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {tmp}\scripten.exe; Parameters: /r:n /q:1; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {app}\Calorie Balance Tracker.exe; Description: {cm:LaunchProgram, CalorieTracker}; Flags: nowait postinstall skipifsilent; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\Calorie Balance Tracker.url

[Registry]
Root: HKCU; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: RUNASADMIN; Flags: uninsdeletevalue
Root: HKLM; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: WINXPSP2; Flags: uninsdeletevalue

[Code]
function GetAppFolder(Param: String): String;
begin
  if InstallOnThisVersion('0,6', '0,0') = irInstall then
    Result := 'C:\Users\Public\' + ExpandConstant('{username}')
  else
    Result := ExpandConstant('{pf}');
end;

[UninstallRun]
Filename: "{app}\Calorie Balance Tracker.exe"; Parameters: "/ExitSurvey"; RunOnceId: ExitSurvey


