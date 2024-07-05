VERSION 5.00
Begin VB.Form frmSplash2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Edition - Calorie Balance"
   ClientHeight    =   6855
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   6405
   ClipControls    =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Buy Now"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
End
Attribute VB_Name = "frmSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public BaseUrl As String
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub Command1_Click()


    On Error GoTo Err_Proc
Me.Hide
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "Command1_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command2_Click()
    Dim buyURL As String
    buyURL = regGetBuyURL("Calorie Balance Diet", "Calorie Balance Tracker", "0")
    
    If buyURL = "" Then
        ' BuyURL doesn't exsits in registry, default it
          If RegNow Then
             buyURL = "https://www.regnow.com/softsell/nph-softsell.cgi?item=16565-1"
          Else
             buyURL = Branding("paypage")
          End If
    End If
  Call OpenURL(buyURL)

  
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "Command2_Click", Err.Description
    Resume Exit_Proc


End Sub



Private Sub UnlockProgram()


    On Error GoTo Err_Proc

Dim email As String
Dim unlockC As String
Dim sum As Long, junk As String
Dim JunkO As String, i As Long

Success:
   SaveSetting App.Title, "Settings", "PD", True
   SaveSetting App.Title, "Settings", "Webmember", False
   Paid = True
   MsgBox "Thank you.  You have successfully entered the serial number", vbOKOnly, ""
   Unload Me

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "UnlockProgram", Err.Description
    Resume Exit_Proc


End Sub
Private Sub Command4_Click()


    On Error GoTo Err_Proc
End
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "Command4_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_Load()
  Command2.Visible = RegNow

    On Error GoTo Err_Proc
    Caption = "Free Version - " & Branding("caption")
    If FreeVersion Then
     ' Command3.Visible = False
    End If
    If frmLogin.RemainingDays > 0 Then
           'Caption = "Demo - " & Me.RemainingDays & " Days Left"
    Else
           'Caption = "Trial period finished"
         '  Command1.Enabled = False
           'CNewUSer.Enabled = False
    End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "Form_Load", Err.Description
    Resume Exit_Proc


End Sub


Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)


    On Error GoTo Err_Proc
If InStr(1, URL, "aaa", vbTextCompare) <> 0 Then
  Cancel = True
    Dim buyURL As String
    buyURL = regGetBuyURL("Calorie Balance Diet", "Calorie Balance Tracker", "0")
    
    If buyURL = "" Then
        ' BuyURL doesn't exsits in registry, default it
          If RegNow Then
             buyURL = "https://www.regnow.com/softsell/nph-softsell.cgi?item=16565-1"
          Else
             buyURL = Branding("paypage")
          End If
    End If
  Call OpenURL(buyURL)
End If
If InStr(1, URL, "unlock", vbTextCompare) <> 0 Then
'  Cancel = True
  'Call Command3_Click
End If

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "wb_BeforeNavigate2", Err.Description
    Resume Exit_Proc


End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)


    On Error GoTo Err_Proc
 
   Dim j As String, jj As String
  
   
     
  If InStr(1, URL, "checkserial", vbTextCompare) <> 0 Then
   
    
    If j = "01" Then
       MsgBox "You must enter an unlock code into the appropriate box.", vbOKOnly, ""
    ElseIf j = "1" Then
       MsgBox "You have entered an incorrect unlock code.  Please try again.", vbOKOnly, ""
    ElseIf j = "2" Then
       UnlockProgram
    ElseIf j = "3" Then
       MsgBox "This serial number has been removed due to piracy.  If you are legitamate, please send an email to support@CalorieBalanceDiet.com.", vbOKOnly, ""
       OpenURL "mailto:support@caloriebalancediet.com"
    End If
  End If

errhandl:
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSplash", "WB_DocumentComplete", Err.Description
    Resume Exit_Proc


End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
