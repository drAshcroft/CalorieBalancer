VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FNewUser 
   Caption         =   "New Profile"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser Display 
      Height          =   5655
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame FWEB 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      Begin VB.CommandButton CBack 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmNewUser.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Last Page"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   615
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmNewUser.frx":6424
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Next Page"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   1320
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label Label1 
      Height          =   1050
      Left            =   5610
      TabIndex        =   0
      Top             =   975
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "FNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PageShown As Long
Public newUserDone As Boolean
Public WebSign As Boolean
Public DoSummary As Boolean
Private mASP As New ASP
Dim FreePass As Boolean
Dim CancelAll As Boolean
Dim URLs As New Collection, UrlIndex As Long
Public DoNewUser As Boolean
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub MakeReport(Filename As String)

On Error GoTo errhandl

Dim GS As uGraphSurface
Call InitASP(SC, GS, mASP) '

Dim template As String
Dim ff As Long

'A.SetScript SC
ff = FreeFile
template = mASP.LoadTemplate(App.path & "\resources\temp\", Filename)
Open App.path & "\resources\temp\common.htm" For Output As #ff
Print #ff, template
Close #ff
FreePass = True
Display.Navigate2 App.path & "\resources\temp\common.htm"

Dim i As Long
UrlIndex = URLs.Count
For i = 1 To URLs.Count
  If Filename = URLs(i) Then
     UrlIndex = i
     GoTo Exitsub
  End If
Next i
If UrlIndex = URLs.Count Then
   URLs.Add Filename
   UrlIndex = URLs.Count
End If
Exitsub:

If UrlIndex = URLs.Count Then
   Command1.Enabled = False
Else
  Command1.Enabled = True
End If
If UrlIndex = 1 Then
  CBack.Enabled = False
Else
  CBack.Enabled = True
End If

Exit Sub
errhandl:


End Sub

Private Sub CBack_Click()
On Error Resume Next
If UrlIndex - 1 > 0 Then
  Call ProcessForm("Back.htm")
  MakeReport URLs(UrlIndex - 1)
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If UrlIndex + 1 <= URLs.Count Then
  Call ProcessForm("forward.htm")
  MakeReport URLs(UrlIndex + 1)
End If
End Sub

Private Function ProcessForm(URL As Variant) As Boolean
On Error Resume Next

   On Error GoTo errhandl
   Dim junk As String
   mASP.SaveForm Display.document, ProcessForm
   
   junk = "false"
   junk = LCase$(mASP.Form("webcheck"))
   
   ProcessForm = False
   
   If junk = "true" Then
     ' ProcessForm = WebCheck
   End If
   
   Dim i As Long
   i = InStr(1, URL, "done.htm", vbTextCompare)
   If i <> 0 Then
     ProcessForm = True
     
     For i = 1 To 100000
       DoEvents
     Next i
     Dim temp As Recordset
     Set temp = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
     Call SaveSetting(App.Title, "usersettings", "user", CurrentUser.Username)
     Call SaveSetting(App.Title, "Usersettings", "password", temp("password"))
     Dim pss As String
     pss = temp("password")
     temp.Close
     Set temp = Nothing
     If DoSummary And Not DoNewUser Then
        MsgBox "Edits has been saved.  Thank you.", vbOKOnly, ""
     Else
        MsgBox "Your profile has been loaded.  Thank you.", vbOKOnly, ""
     End If
     DoNewUser = False
     newUserDone = True
     Unload Me
     If frmLogin.Visible Then
       frmLogin.txtUserName = CurrentUser.Username
       frmLogin.txtPassword = pss
       frmLogin.AutoLogin = True
       frmLogin.ZOrder
     End If
     Exit Function
   End If

   
   Exit Function
errhandl:
MsgBox "There has been an untracable error.  Please check all fields and try again.", vbOKOnly, ""
ProcessForm = True

End Function
Private Sub Display_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next

If InStr(1, URL, "BMR.htm", vbTextCompare) <> 0 Then
  Call OpenURL(URL & "")
  
  Cancel = True
  Exit Sub
End If
If CancelAll Then
   Cancel = True
   Exit Sub
End If
If FreePass Then
   FreePass = False
   Exit Sub
End If

   Cancel = ProcessForm(URL)
   If Not Cancel Then
     Dim Parts() As String
     If InStr(1, URL, "how to get started", vbTextCompare) = 0 Then
        Parts = Split(URL, "\")
       
        MakeReport App.path & "\resources\new user\" & Parts(UBound(Parts))
        Cancel = True
     End If
   End If

End Sub


Private Sub Display_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If InStr(1, URL, "common", vbTextCompare) <> 0 Then
    If DoNewUser Then
      Call ProcessForm("done.htm")
    End If
'    Stop
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
   FWEB.Visible = True
   Label1.Visible = True
   If DoSummary Then
     MakeReport App.path & "\resources\new user\summary.htm"
  
   Else
     MakeReport App.path & "\resources\new user\summary.htm"
   End If
   Call Form_Resize
End Sub

Private Sub Form_Resize()

On Error GoTo errhandl
If FirstRun Then
Display.Move 0, 0, ScaleWidth, ScaleHeight - 50  ' CDone.Height - 100

Else


Display.Move 0, FWEB.Height, ScaleWidth, ScaleHeight - FWEB.Height - 50 ' CDone.Height - 100
FWEB.Move (Me.ScaleWidth - FWEB.Width) / 2, 0
Label1.Move Me.ScaleWidth / 2 - Label1.Width / 2, Me.ScaleHeight / 2
End If
errhandl:

End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

     Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
