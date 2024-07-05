; InnoScript Version 7.0  Build 0  Trial
; Randem Systems, Inc.
; Copyright 2003-2007
; website:  http://www.randem.com
; support:  http://www.innoscript.com/cgi-bin/discus/discus.cgi

; Date: June 06, 2007

;              VB Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript\InnoScript 7.0\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Project1.vbp
; Inno Setup Script Output File (.iss):   C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\Installer\Test Release.iss

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
; Microsoft FlexGrid Control 6.0 (SP3) - (MSFLXGRD.OCX)
; Microsoft Rich Textbox Control 6.0 (SP6) - (RICHTX32.OCX)
; Microsoft Windows Common Controls-2 6.0 (SP4) - (MSCOMCT2.OCX)


[Setup]

;-----------------------------------------------------------------------------------------------------
; Taken from VBP Project File Parameters AppName, AppName AppVersion and Company
;-----------------------------------------------------------------------------------------------------

AppName=CalorieTracker
AppVerName=CalorieTracker 2.5.0
AppPublisher=Sundance

;-----------------------------------------------------------------------------------------------------

AppPublisherURL=http://www.caloriebalancediet.com
AppSupportURL=http://www.caloriebalancediet.com
AppVersion=2.5.0
VersionInfoVersion=2.5.0
AllowNoIcons=no
LicenseFile=C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\EULA.txt
DefaultGroupName=Calorie Balance Tracker
DefaultDirName={code:GetAppFolder}\calorie balance tracker
AppCopyright=
PrivilegesRequired=None
MinVersion=4.0,4.0
Compression=lzma
OutputBaseFilename=SetupCalorieBalanceTrackerRelease

[Tasks]
Name: desktopicon; Description: Create a &Desktop Icon; GroupDescription: Additional Icons:

[Files]
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\msvbvm60.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\oleaut32.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\olepro32.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\asycfilt.dll; DestDir: {sys}; Flags:  sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\stdole2.tlb; DestDir: {sys}; Flags:  uninsneveruninstall sharedfile regtypelib onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 7.0\vb 6 redist files\comcat.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist; OnlyBelowVersion: 0,6
Source: c:\documents and settings\brian\desktop\final calories 2.5\calorie balance tracker.exe; DestDir: {app}; Flags:  ignoreversion onlyifdoesntexist;
Source: ieframe.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: mshtml.tlb; DestDir: {sys}; Flags:  uninsneveruninstall sharedfile regtypelib onlyifdoesntexist;
Source: msscript.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: comdlg32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: mscomctl.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: msflxgrd.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: richtx32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: mscomct2.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile onlyifdoesntexist;
Source: psapi.dll; DestDir: {sys}; Flags:  sharedfile; 
Source: iertutil.dll; DestDir: {sys}; Flags:  sharedfile; 
Source: C:\Documents and Settings\Brian\Desktop\Final Calories 2.5\New Folder\Resources\; DestDir: {app}\Resources; Flags: ignoreversion

[INI]
Filename: {app}\Calorie Balance Tracker.url; Section: InternetShortcut; Key: URL; String: http://www.caloriebalancediet.com

[Icons]
Name: {group}\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; WorkingDir: {app}
Name: {group}\CalorieTracker on the Web; Filename: {app}\Calorie Balance Tracker.url
Name: {group}\Uninstall CalorieTracker; Filename: {uninstallexe}
Name: {userdesktop}\CalorieTracker; Filename: {app}\Calorie Balance Tracker.exe; Tasks: desktopicon; WorkingDir: {app}

[Run]
Filename: {app}\Calorie Balance Tracker.exe; Description: Launch CalorieTracker; Flags: nowait postinstall skipifsilent; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\Calorie Balance Tracker.url

[Registry]
Root: HKCU; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: RUNASADMIN
Root: HKLM; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueType: string; ValueName: {app}\Calorie Balance Tracker.exe; ValueData: WINXPSP2
