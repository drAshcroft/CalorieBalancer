VERSION 5.00
Begin VB.Form FNewExercise 
   Caption         =   "New Exercise"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5805
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin CalorieBalance.AutoCompleteEX EList 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox TNumber 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Example: 200 calories/hr"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
      Begin VB.TextBox TFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Example:  0.06* minutes*weight"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "I have a number for calories burned in 1 hour (150 lb. person)"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "I have a formula"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "This exercise is like:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin CalorieBalance.AutoCompleteEX LList 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   $"FNewExercise.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Exercise Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FNewExercise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub SetDB()
On Error Resume Next
   EList.AddDataBase DB
   LLike.AddDataBase DB
End Sub

Private Sub Command1_Click()
On Error GoTo errhandl
Dim temp As Recordset
Dim temp2 As Recordset
Dim ExName As String, junk As String

ExName = Replace(EList.Text, "'", "''")

Set temp = DB.OpenRecordset("Select * from abbrevexercise " _
   & "where exercisename = '" & ExName & "';", dbOpenDynaset)
If temp.EOF Then
  Set temp2 = DB.OpenRecordset("Select max(index) as MAXit from abbrevexercise;", dbOpenDynaset)
  temp.AddNew
  temp("index") = temp2("Maxit") + 1
  temp2.Close
  Set temp2 = Nothing
Else
  temp.Edit
End If
temp("exercisename") = ExName
temp("usage") = 5
temp("Increase") = 1
If Option1.Value = True Then
    junk = Replace(LLike.Text, "'", "''")
    Set temp2 = DB.OpenRecordset("Select * from abbrevexercise " _
      & "where exercisename = '" & junk & "';", dbOpenDynaset)
    If temp2.EOF Then
       MsgBox "I do not recognize the similar exercise" & vbCrLf _
       & "Please make sure that you have selected a name from the list", vbOKOnly, ""
       Exit Sub
    End If
    temp("Formula") = temp2("Formula")
    temp("increase") = temp2("Increase")
    temp2.Close
    Set temp2 = Nothing
ElseIf Option2.Value = True Then
    junk = TFormula.Text
    If junk = "" Then
       MsgBox "Please enter a formula into the provided box" & vbCrLf _
       & "Example:  0.6 * weight * minutes" & vbCrLf _
       & "Example:  10 * minutes", vbOKOnly, ""
       Exit Sub
    End If
    junk = Replace(junk, "minutes", "par0", , , vbTextCompare)
    junk = Replace(junk, "minute", "par0", , , vbTextCompare)
    temp("Formula") = junk
ElseIf Option3.Value = True Then
    
    junk = Val(TNumber.Text) / 60 / 150
    If junk = 0 Then
       MsgBox "Please enter number of calories burned for a 150 lb person over 1 hour.", vbOKOnly, ""
       Exit Sub
    End If
    temp("Formula") = Round(junk, 6) & "*par0*weight"
End If
temp.Update
temp.Close
Set temp = Nothing


Unload Me
Exit Sub
errhandl:
MsgBox "Cannot save exercise. Please check all the fields.", vbOKOnly, ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
HelpWindowHandle = htmlHelpTopic(frmMain.hWnd, HelpPath, _
         0, "newHTML/Exercise_tracker.htm#NewExercise")

'Call HTMLHelp(frmMain.hWnd, App.HelpFile, 0, ByVal 0)
If Err.Number <> 0 Then Call MsgBox("Cannot find help file.", vbOKOnly, "")
End Sub

Private Sub EList_ExitFocus()
On Error Resume Next
EList.CloseBox
End Sub

Private Sub Form_Load()
On Error Resume Next
EList.AddDataBase DB
LLike.AddDataBase DB
LLike.ZOrder
EList.ZOrder

End Sub

Private Sub LLike_ExitFocus()
On Error Resume Next
LLike.CloseBox
End Sub

Private Sub Option1_Click()
On Error Resume Next
'LLike.Enabled = True
'TFormula.Enabled = False
'TNumber.Enabled = False

LLike.Visible = True
LLike.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Option2_Click()
'LLike.Enabled = False
'TFormula.Enabled = True
'TNumber.Enabled = False
On Error Resume Next
LLike.Visible = False
LLike.Enabled = False
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Option3_Click()
'LLike.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = True
On Error Resume Next
LLike.Visible = False
LLike.Enabled = False
Frame1.Visible = False
Frame2.Visible = True

End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
