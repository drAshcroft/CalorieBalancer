VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Premade Info File"
      Height          =   615
      Left            =   3720
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox txtUserName 
      Height          =   2010
      ItemData        =   "frmLogin3.frx":0000
      Left            =   1200
      List            =   "frmLogin3.frx":0002
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download Web Profile"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CNewUSer 
      Caption         =   "Setup Username, Password, and Calories with the Body Wizard"
      Height          =   1215
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2505
      TabIndex        =   4
      Tag             =   "Cancel"
      Top             =   3240
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Tag             =   "OK"
      Top             =   3240
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   2
      Top             =   2355
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remember the password"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5760
      X2              =   5640
      Y1              =   360
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5760
      X2              =   5640
      Y1              =   120
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   6000
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Here to set up a free new account.   This will give you a username and password."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6000
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Tag             =   "&Password:"
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Tag             =   "&User Name:"
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public RemainingDays As Long

Public USER As String
Public ok As Boolean
Public NewUSer As Boolean
Public SeriesDay As Long
Public AutoLogin As Boolean

Public LaunchNewUser As Boolean

Private UrlMode As Integer

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub CNewUSer_Click()
  On Error Resume Next
  FNewUserD.Show vbModal, Me
  If FNewUserD.ShowSummary Then FUserSummary.Show vbModal, Me
  If AutoLogin Then
     cmdOK_Click
  End If
  'Me.Hide
End Sub

Private Sub Command1_Click()


    On Error GoTo Err_Proc
HelpWindowHandle = htmlHelpTopic(frmMain.hWnd, HelpPath, 0, "newHTML/login.htm")
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmLogin", "Command1_Click", Err.Description
    Resume Exit_Proc


End Sub



Private Sub Command2_Click()


    On Error GoTo Err_Proc
    Dim i As Long
    Dim Username As String
    Username = ""
    For i = 0 To txtUserName.ListCount - 1
       If txtUserName.Selected(i) Then
           Username = txtUserName.List(i)
       End If
    Next i
If Username = "" Or txtPassword = "" Then
   MsgBox "Please enter in your username and password from the webpage." & vbCrLf & "If you have not signed up yet, please to the website.", vbOKOnly, ""
   Exit Sub
End If
Me.WindowState = 2

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmLogin", "Command2_Click", Err.Description
    Resume Exit_Proc


End Sub
Public Sub RefreshList()
    txtUserName.Clear
    Call Form_Load
End Sub

Private Sub Command3_Click()
On Error GoTo errhandl
CD.CancelError = True
CD.Filter = "Calorie Balance File (*.cbm)|*.cbm"
CD.ShowOpen
Call REadScriptMod.UpdateScript(CD.Filename)
Call Form_Load
errhandl:
End Sub

Private Sub Form_Load()
On Error Resume Next
'    Dim sBuffer As String
    Dim lSize As Long
    Label1.ZOrder
    If FreeVersion Then
      Command2.Visible = False
    End If
    
    txtPassword.Left = txtUserName.Left
    txtPassword.Width = txtUserName.Width
  
    USER = GetSetting(App.Title, "UserSettings", "User", "")
    Check1.Value = GetSetting(App.Title, "UserSettings", "RemeberPassword", 0)
    If Check1.Value = 1 Or FirstRun Then
         txtPassword.Text = GetSetting(App.Title, "UserSettings", "Password", "")
    End If
    Dim Lside As Long
    Lside = Len(USER)
    
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("select * from profiles;", dbOpenDynaset)
    txtUserName.Clear
    txtUserName.AddItem "Please select an user"
    While Not RS.EOF
      If IsNull(RS("user")) = False Then
        If Not RS("user") = "average" Then
           txtUserName.AddItem Trim(RS("user"))
        End If
      End If
      RS.MoveNext
    Wend
    
    Dim i As Long
    txtUserName.Selected(txtUserName.ListCount - 1) = True
    For i = 0 To txtUserName.ListCount - 1
       If LCase$(txtUserName.List(i)) = LCase$(USER) Then txtUserName.Selected(i) = True
    Next i
    
    'make sure that the controls are enabled
    cmdOK.Enabled = True
    CNewUSer.Enabled = True
 
    'extort money from demos
   
    If Paid = False Then
       If Me.RemainingDays > 0 Then
           Caption = "Login " ' & Me.RemainingDays & " Days Left"
       Else
       '    Caption = "Trial period finished"
       '    cmdOK.Enabled = False
       '    CNewUSer.Enabled = False
           'Command3.Enabled = False
       End If
       
       
       
       
       
       'Me.Height = cmdOK.Top + cmdOK.Height + 700 + Command1.Height
    Else
       Caption = "Login"
    End If
    
    Dim Hints As New Collection
    Hints.Add "When adding meals to calorie counter, pressing shift and then dragging the meal will allow two of the same meal to be placed."
    Hints.Add "When in the meal planner, pressing shift and then dragging a meal on the board will allow you to move a meal from one cell to another."
    Hints.Add "You can print shopping lists from the file menu."
    Hints.Add "You can delete the user from the file menu."
    Hints.Add "You can download example meal plans from the webpage."
    Hints.Add "You can upload recipes to the website, and download those recipes that have been already donated."
    Hints.Add "By clicking options, you can change the display to show the percentage of calories consumed by each item."
    Hints.Add "Click make into a meal to turn any meal in the calorie counter into a meal."
    Hints.Add "You can keep an eating journal in the journal function to record all your emotions and triggers."
    Hints.Add "Can't find a food? Click internet search from the new food toolbar, or big search from the search function and the program will read the needed food right off the internet."
    Hints.Add "You can black out a cell in the exercise tracker by entering ' - ' in the cell."
    Hints.Add "You disable a cell by entering ' * ' before the planned exercise.  Then clicking on the box later will make the exercise 'active'."
    Hints.Add "You can add new exercises to the exercise tracker by looking for similar exercises from the list."
    
    Randomize Timer
    Label1.Caption = "   " & Hints(1 + Int((Hints.Count) * Rnd()))
   
'    Command2.Enabled = True
'       Command2.Visible = True
    
'    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 50
'    cmdCancel.Top = cmdOK.Top
    If LaunchNewUser Then
      LaunchNewUser = False
      FNewUserD.Show vbModal, Me
    End If
  
End Sub



Private Sub cmdCancel_Click()

On Error Resume Next
    If Not DoDebug Then HtmlHelp Me.hWnd, "", HH_CLOSE_ALL, 0&
    ok = False
    NewUSer = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()

 On Error GoTo errhandl
 
    Dim i As Long
    Dim Username As String
    Username = ""
    For i = 0 To txtUserName.ListCount - 1
       If txtUserName.Selected(i) Then
           If i = 0 Then
             MsgBox "You must select a username.  Please click a username to select.", vbOKOnly, ""
             Exit Sub
           End If
           Username = txtUserName.List(i)
       End If
    Next i
    
    If Trim$(Username) = "" Then
    
       MsgBox "You must select a username.  Please click a username to select.", vbOKOnly, ""
       Exit Sub
    End If
    Dim temp As Recordset
    Dim Pass As String
    Set temp = DB.OpenRecordset("Select password from profiles where user = '" & Trim$(Username) & "';", dbOpenDynaset)
    
    If temp.EOF Then
      
      Call MsgBox("Your profile is not on this computer." & vbCrLf & _
      "Please check your caplock key and try again." & vbCrLf & _
      "Also make sure that you are using your username and not the full name that you have entered", vbYesNo, "")
    
       Exit Sub
    End If
    Pass = temp("password")
    temp.Close
    Set temp = Nothing
    
    
    If LCase$(txtPassword.Text) = LCase$(Pass) Then
        frmMain.LastUser = Username
        USER = Username
        ok = True
        Me.Hide
    Else
errhandl:
        MsgBox "The username and password do not match, please try again" & vbCrLf & "check to see if your caps loc is on.", , "Login"
        
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        Me.MousePointer = 0
    End If
End Sub


Private Sub Form_Resize()
Dim W As Single
Dim L As Single
On Error Resume Next
W = CNewUSer.Width
L = CNewUSer.Left

If FirstRun Then
 '  Me.Width = Label2.Left + Label2.Width + 200
 '  Label2.Caption = "Click Here to set up a free new account." & vbCrLf & "   This will give you a username and password."
 '  Label2.Visible = True
 '  Line1.Visible = True
 '  Line2.Visible = True
 '  Line3.Visible = True
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  If Not DoDebug Then HtmlHelp Me.hWnd, "", HH_CLOSE_ALL, 0&
  SaveSetting App.Title, "UserSettings", "User", txtUserName.Text
  SaveSetting App.Title, "UserSettings", "RemeberPassword", Check1.Value
  If Check1.Value = 1 Then
    SaveSetting App.Title, "UserSettings", "Password", txtPassword.Text
  Else
    SaveSetting App.Title, "UserSettings", "Password", ""
  End If
  CloseProgram = True
'  End
End Sub





Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
