VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Web Downloader"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Continue"
      Height          =   855
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3840
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Check Computer's Readiness"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download and Save MDAC"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

Private Function GetVersionInfo() As String
    Dim myOS As OSVERSIONINFOEX
    Dim bExInfo As Boolean
    Dim sOS As String

    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    'try win2000 version
    If GetVersionEx(myOS) = 0 Then
        'if fails
        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Microsoft Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
   
    With myOS
        'is version 4
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            'nt platform
            Select Case .dwMajorVersion
            Case 3, 4
                sOS = "Microsoft Windows NT"
            Case 5
                sOS = "Microsoft Windows 2000"
            End Select
            If bExInfo Then
                'workstation/server?
                If .wProductType = VER_NT_SERVER Then
                    sOS = sOS & " Server"
                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
                    sOS = sOS & " Domain Controller"
                ElseIf .wProductType = VER_NT_WORKSTATION Then
                    sOS = sOS & " Workstation"
                End If
            End If
           
            'get version/build no
            sOS = sOS & " Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
           
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            'get minor version info
            If .dwMinorVersion = 0 Then
                sOS = "Microsoft Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                sOS = "Microsoft Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                sOS = "Microsoft Windows Millenium"
            Else
                sOS = "Microsoft Windows 9?"
            End If
            'get version/build no
            sOS = sOS & "Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & StripTerminator(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
        End If
    End With
    GetVersionInfo = sOS
End Function
Private Function StripTerminator(sString As String) As String
    StripTerminator = Left$(sString, InStr(sString, Chr$(0)) - 1)
End Function

Private Sub Command1_Click()
On Error GoTo errhandl
'CD.CancelError = True
'CD.InitDir = App.Path & "\VB_DCOM_MDAC_JET_AutoSetup.exe"
'CD.FileName = "VB_DCOM_MDAC_JET_AutoSetup.exe"
'CD.ShowSave


Shell "explorer http://www.caloriebalancediet.com/mdac.asp", vbNormalFocus
 '  Call DownloadFile("http://www.caloriebalancediet.com/installers/VB_DCOM_MDAC_JET_AutoSetup.exe", CD.FileName)
 DoEvents
End
errhandl:
End Sub

Private Sub Command2_Click()
   DoEvents
   Call DownloadFile("http://www.caloriebalancediet.com/installers/VB_DCOM_MDAC_JET_AutoSetup.exe", App.Path & "\VB_DCOM_MDAC_JET_AutoSetup.exe")
   Shell App.Path & "\VB_DCOM_MDAC_JET_AutoSetup.exe /NORESTART ", vbNormalFocus
End Sub

Private Sub Command3_Click()
  Unload Me
  End
End Sub

Private Sub Command4_Click()

On Error Resume Next


Set j = CreateObject("DAO.dbEngine.36")

If j Is Nothing Then
  Set j = GetObject("DAO.dbEngine.36")
  If j Is Nothing Then
     Command1.Enabled = True
     Command2.Enabled = True
     'Command3.Enabled = True
     Label1.Caption = "Your computer is missing the Microsoft Data Access Components" & vbCrLf & _
     "Please choose a method to download this file to the right.  Please keep in mind that this " & _
     "is a very large file and will take up to 15 minutes to download."
     
     
     'Call UpdateComputer
     'GoTo s
     Shell "explorer http://www.caloriebalancediet.com/BadDatabase.asp", vbNormalFocus
     End
     Exit Sub
  End If
End If

Set db = j.OpenDatabase(Interaction.Command$ & "\resources\proto.mdb", True, False)

If db Is Nothing Or Err.Number <> 0 Then
     Command1.Enabled = True
     Command2.Enabled = True
     'Command3.Enabled = True
     Label1.Caption = "Your computer is missing the Microsoft Data Access Components" & vbCrLf & _
     "Please choose a method to download this file to the right.  Please keep in mind that this " & _
     "is a very large file and will take up to 15 minutes to download."
     Shell "explorer http://www.caloriebalancediet.com/BadDatabase.asp", vbNormalFocus
     End
     Exit Sub
End If

s:
Set db = Nothing
Set j = Nothing
Label1.Caption = "   Your computer has all the componites already installed.  Please continue."
Command5.Visible = True
End
End Sub

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
On Error GoTo errhandl
    Dim lngRetVal As Long
    Dim O As Boolean
    O = False
r:
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
   ' MsgBox LocalFilename
    If lngRetVal = 0 Then
       DownloadFile = True
    ElseIf O = False Then
       MsgBox "You must be connected to the internet for this operation to work. " & vbCrLf & "Please connect your computer to the internet and then click OK", vbOKOnly, ""
       O = True
       GoTo r
      
    End If
errhandl:
End Function
Private Function UpdateComputer()
   'Call MsgBox("Your system needs a few componites downloaded.  Please be patient while we download the required files.", vbOKOnly, "")
   Call DownloadFile("http://www.caloriebalancediet.com/installers/VB_DCOM_MDAC_JET_AutoSetup.exe", App.Path & "\VB_DCOM_MDAC_JET_AutoSetup.exe")
   'Shell App.Path & "\VB_DCOM_MDAC_JET_AutoSetup.exe /NORESTART ", vbNormalFocus

End Function

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
  Command4_Click
End Sub
