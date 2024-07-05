VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.CommandButton Command3 
      Caption         =   "Download Internet Profile"
      Height          =   510
      Left            =   3930
      TabIndex        =   9
      Top             =   480
      Width           =   1725
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   30
      Left            =   105
      TabIndex        =   8
      Top             =   960
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
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
   Begin VB.CommandButton Command2 
      Caption         =   "Get/Enter Serial"
      Height          =   345
      Left            =   3945
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CNewUSer 
      Caption         =   "Make New User"
      Height          =   375
      Left            =   3915
      TabIndex        =   3
      Top             =   90
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2490
      TabIndex        =   4
      Tag             =   "Cancel"
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   1290
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   990
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public RemainingDays As Long

Public USER As String
Public ok As Boolean
Public NewUSer As Boolean
Public SeriesDay As Long

Private UrlMode As Integer

Private Sub CNewUSer_Click()
  On Error Resume Next
  NewUSer = True
  Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
Call OpenURL(App.path & "\Resources\New User\congratulations.htm")
End Sub

Private Sub Command2_Click()
On Error GoTo errhandl2
Dim Ret
errhandl:
Ret = InputBox("Please enter serial number", "")

Dim Serial As String
Serial = "1w24f5h794e3"
If LCase$(Ret) = Serial Then
   SaveSetting App.Title, "Settings", "PD", True
   SaveSetting App.Title, "Settings", "Webmember", False
   
   Paid = True
   
   UrlMode = 1
   
   MsgBox "Thank you.  You have successfully entered the serial number", vbOKOnly, ""
   
   Call Form_Load
   
   
ElseIf Ret <> "" Then
   MsgBox "There is an error please try again.", vbOKOnly, ""
   GoTo errhandl
End If

errhandl2:
End Sub

Private Sub Command3_Click()

If txtUserName.Text = "" Then
   MsgBox "Please enter your username and password to the left and then try again.", vbOKOnly, ""
   Exit Sub

End If

         Me.MousePointer = 11
         UrlMode = 2
         WB.Navigate2 "http://www.caloriebalancediet.com/sync.asp?username=" & txtUserName.Text _
          & "&password=" & txtPassword.Text

End Sub

Private Sub Form_Load()
On Error Resume Next
'    Dim sBuffer As String
    Dim lSize As Long
    WB.Navigate2 "about:blank"
    
    'get the last username
    USER = GetSetting(App.Title, "UserSettings", "User", "")
    Lside = Len(USER)
    If Lside <> 0 Then
        txtUserName.Text = USER
    Else
        txtUserName.Text = vbNullString
    End If
    
    'make sure that the controls are enabled
    cmdOK.Enabled = True
    CNewUSer.Enabled = True
 
    'extort money from demos
    If Paid = False Then
       If Me.RemainingDays > 0 Then
           Caption = "Demo - " & Me.RemainingDays & " Days Left"
       Else
           Caption = "Trial period finished"
           cmdOK.Enabled = False
           CNewUSer.Enabled = False
       End If
       Command2.Enabled = True
       'Me.Height = cmdOK.Top + cmdOK.Height + 700 + Command1.Height
    Else
       Caption = "Login"
    End If
    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 50
    cmdCancel.Top = cmdOK.Top
End Sub



Private Sub cmdCancel_Click()
On Error Resume Next
    ok = False
    NewUSer = False
    Unload Me
End Sub


Private Sub cmdOK_Click()
 On Error GoTo errhandl
    If Trim$(txtUserName.Text) = "" Then Exit Sub
    Dim temp As Recordset
    Dim Pass As String
    Set temp = DB.OpenRecordset("Select password from profiles where user = '" & txtUserName.Text & "';", dbOpenDynaset)
    
    If temp.EOF Then
      Dim Ret As VbMsgBoxResult
      Ret = MsgBox("Your profile is not on this computer." & vbCrLf & "Do you wish to download your internet profile?" _
           , vbYesNo, "Login Failed")
       If Ret = vbYes Then
         Me.MousePointer = 11
         UrlMode = 2
         WB.Navigate2 "http://www.caloriebalancediet.com/sync.asp?username=" & txtUserName.Text _
          & "&password=" & txtPassword.Text
       End If
       temp.Close
       Set temp = Nothing
       Exit Sub
    End If
    Pass = temp("password")
    temp.Close
    Set temp = Nothing
    
    
    If txtPassword.Text = Pass Then
        frmMain.LastUser = txtUserName.Text
        USER = txtUserName.Text
        ok = True
        Unload Me
    Else
errhandl:
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub


Private Sub Form_Resize()
Dim w As Single
Dim L As Single
w = CNewUSer.Width
L = CNewUSer.Left
'Command1.Move L, Command1.Top, w
Command2.Move L, Command2.Top, w
Command3.Move L, Command3.Top, w


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  SaveSetting App.Title, "UserSettings", "User", txtUserName.Text
End Sub

Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
If InStr(1, URL, "dnserror.htm", vbTextCompare) <> 0 Then
   MsgBox "Cannot find server.  Please connect to the internet", vbOKOnly, ""
End If
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error GoTo errhandl
  If InStr(1, URL, "caloriebalancediet", vbTextCompare) = 0 Then
    Exit Sub
  End If
  
 Dim Document As HTMLDocument
  Set Document = WB.Document
  DoEvents
  
  If UrlMode = 2 Then
  
  Dim temp As Recordset
  Set temp = DB.OpenRecordset("Select * from profiles;", dbOpenDynaset)

  
  
  Dim junk As String, lines() As String, parts() As String
  Dim i As Long, numbers As Boolean
  temp.AddNew
 ' MsgBox "Downloadcomplete"
  junk = Document.Body.innerText
  'MsgBox junk
  lines = Split(junk, vbCrLf)
  If Trim$(lines(0)) <> "Record" Then
errhandl:
     MsgBox "No profile is found on the internet", vbOKOnly, ""
     Exit Sub
  End If
  For i = 1 To UBound(lines)
    lines(i) = Trim$(lines(i))
    If lines(i) = "---Numbers---" Then
       numbers = True
    Else
       parts = Split(lines(i), "=")
       parts(0) = Trim$(parts(0))
       parts(1) = Trim$(parts(1))
       If parts(0) = "birthdate" And parts(1) = "" Then parts(1) = Date
       If numbers Then
          temp(parts(0)) = Val(parts(1))
       Else
          temp(parts(0)) = parts(1)
       End If
    End If
  Next i
  temp("StartWeight") = temp("WEight")
  temp("Calpound") = 3500
  temp("carbs") = 0.5
  temp("fat") = 0.25
  temp("Protein") = 0.25
  temp("fiber") = 0.5
  temp("sugar") = 0.1
  temp("OtherWatches") = "Calcium,Iron,Sodium,Vitamin C,Cholesterol"
  temp("Caloriesbalance") = "V2"
  CurrentUser.Username = temp("User")
  temp.Update
  temp.Close
  
  Set temp = Nothing
  Set Document = Nothing
  Me.MousePointer = 0
  cmdOK_Click
  DoEvents
  MsgBox "Please review your user information and then choose a plan", vbOKOnly, ""
  FNewUser.USER = CurrentUser.Username
  FNewUser.PageShown = 0
  FNewUser.Show vbModal, frmMain


  Else
     
 
  End If
End Sub

