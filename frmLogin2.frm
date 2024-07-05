VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   5610
   Icon            =   "frmLogin2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Remember the password"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   45
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   435
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   360
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   1320
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2505
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   1320
      Width           =   1140
   End
   Begin VB.CommandButton CNewUSer 
      Caption         =   "Make New User"
      Height          =   375
      Left            =   3810
      TabIndex        =   4
      Top             =   0
      Width           =   1755
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
      TabIndex        =   9
      Top             =   1800
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Tag             =   "&User Name:"
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Tag             =   "&Password:"
      Top             =   450
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
HelpWindowHandle = htmlHelpTopic(frmMain.hWnd, HelpPath, _
         0, "newHTML/login.htm")
End Sub



Private Sub Form_Load()
On Error Resume Next
'    Dim sBuffer As String
    Dim lSize As Long
  '  wb.Navigate2 "about:blank"
    
    'get the last username
    USER = GetSetting(App.Title, "UserSettings", "User", "")
    Check1.Value = GetSetting(App.Title, "UserSettings", "RemeberPassword", 0)
    If Check1.Value = 1 Or FirstRun Then
         txtPassword.Text = GetSetting(App.Title, "UserSettings", "Password", "")
    End If
    Dim Lside As Long
    Lside = Len(USER)
    If Lside <> 0 Then
        txtUserName.Text = USER
    Else
        txtUserName.Text = ""
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
End Sub



Private Sub cmdCancel_Click()

On Error Resume Next
    If Not DoDebug Then HtmlHelp Me.hWnd, "", HH_CLOSE_ALL, 0&
    ok = False
    NewUSer = False
    Unload Me
End Sub


Private Sub cmdOK_Click()

 On Error GoTo errhandl
    If Trim$(txtUserName.Text) = "" Then Exit Sub
    Dim temp As Recordset
    Dim Pass As String
    Set temp = DB.OpenRecordset("Select password from profiles where user = '" & Trim$(txtUserName.Text) & "';", dbOpenDynaset)
    
    If temp.EOF Then
      
      Call MsgBox("Your profile is not on this computer." & vbCrLf & _
      "Please check your caplock key and try again.", vbYesNo, "")
    
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
        MsgBox "Invalid Login, please try again" & vbCrLf & "check to see if your caps loc is on.", , "Login"
        
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
          Me.MousePointer = 0
    End If
End Sub


Private Sub Form_Resize()
Dim W As Single
Dim l As Single
W = CNewUSer.Width
l = CNewUSer.Left
'WB.Move 0, 0
'WB2.Move WB.Width, 0
'Command1.Move L, Command1.Top, w
'Command2.Move L, Command2.Top, W
'Command3.Move L, Command3.Top, W


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
End Sub


