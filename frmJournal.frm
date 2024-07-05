VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmJournal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Daily Journal"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11430
   ClipControls    =   0   'False
   Icon            =   "frmJournal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11430
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Daily Journal"
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      LargeChange     =   10
      Left            =   11160
      Max             =   100
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.FileListBox templates 
      Height          =   1260
      Left            =   6360
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin CalorieBalance.uGraphSurface GS 
      Height          =   4500
      Left            =   7200
      TabIndex        =   1
      Top             =   -840
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7938
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CurrentY        =   317
      FillColor       =   255
      ScaleHeight     =   240
      ScaleWidth      =   307
      BeginProperty AxisFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty NumberFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CalorieBalance.uGraphSurface GS2 
      Height          =   4500
      Left            =   7800
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7938
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CurrentY        =   317
      FillColor       =   255
      ScaleHeight     =   240
      ScaleWidth      =   307
      BeginProperty AxisFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty NumberFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      Timeout         =   1000000000
   End
   Begin CalorieBalance.JournalPage JournalPage2 
      Height          =   13620
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   24024
   End
   Begin CalorieBalance.JournalPage JournalPage1 
      Height          =   13620
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   24024
   End
   Begin VB.Menu mnuFle 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuFoodLog 
         Caption         =   "Food Log"
         Index           =   0
      End
      Begin VB.Menu mnuTodaNuts 
         Caption         =   "Today's Nutrition"
      End
      Begin VB.Menu mnuProgReport 
         Caption         =   "Progress Report"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
'Property Variables:
Dim FreePass As Boolean
Dim mASP As New ASP
Dim PassPhrase As String
Public ShowMode As Boolean


Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage



Private Sub AllUsers_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
Dim i As Long, USER As String, j As Long
If InStr(1, URL, "newpass", vbTextCompare) <> 0 Then
   Cancel = True
   i = InStr(1, URL, "_") + 1
   j = InStr(i, URL, ".")
   USER = Mid$(URL, i, j - i)
   Dim RS As Recordset
   Dim Newpass As String
   Newpass = InputBox("Please enter new password.", "New Password", "")
   If Newpass = "" Then Exit Sub
   
   Set RS = DB.OpenRecordset("select * from profiles where user='" & USER & "'", dbOpenDynaset)
   RS.Edit
   RS("password") = Newpass
   RS.Update
   RS.Close
   Set RS = Nothing
End If
If InStr(1, URL, "delete", vbTextCompare) <> 0 Then
   Cancel = True
   i = InStr(1, URL, "_") + 1
   j = InStr(i, URL, ".")
   USER = Mid$(URL, i, j - i)
   
   Dim ret As VbMsgBoxResult
   ret = MsgBox("Are you sure that you wish to delete this user?", vbYesNoCancel, "Delete User")
   If ret = vbYes Then
    Dim T As TableDefs, jj As TableDef, junk As String
    Dim temp As Recordset

    Set T = DB.TableDefs
    On Error GoTo NextTable
    For Each jj In T
      junk = jj.Name
      For i = 0 To jj.Fields.Count - 1
          If LCase$(jj.Fields(i).Name) = "user" Then
            GoTo ClearTable
          End If
      Next i
      GoTo NextTable
ClearTable:
      Set temp = DB.OpenRecordset("select * from " & junk & " where user='" & USER & "';", dbOpenDynaset)
      While Not temp.EOF
       temp.Delete
       temp.MoveNext
      Wend
      temp.Close
NextTable:
      Err.Clear
    Next
    ShowMode = True
    Call Form_Load
   End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call JournalPage1.SetDay(DisplayDate)
If ShowMode Then
      DoEvents
      ShowMode = False
Else
     
     GS2.Width = Me.ScaleX(200, vbPixels, vbTwips)
     GS2.Height = Me.ScaleY(175, vbPixels, vbTwips)
     DoEvents

     Caption = "Journal"
'     Me.WindowState = 2
     mnuPrint.Visible = True
     mnuReports.Visible = True
     
     Call Form_Resize
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
VScroll1.Move ScaleWidth - VScroll1.Width, 0, VScroll1.Width, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
   JournalPage1.SaveJournal
End Sub

Private Sub JournalPage1_CloseRequest()
    Unload Me
End Sub

Private Sub mnu1_Click(Index As Integer)
   On Error Resume Next
   
End Sub



Private Sub mnuExit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub MakeReport(Filename As String)
On Error GoTo errhandl
Call InitASP(SC, GS, mASP) '

Dim Template As String
Dim ff As Long

'A.SetScript SC
ff = FreeFile
mASP.PassPhrase = PassPhrase
Template = mASP.LoadTemplate(App.path & "\resources\temp\", Filename)
Open App.path & "\resources\temp\common.htm" For Output As #ff
Print #ff, Template
Close #ff
FreePass = True
Call OpenURL(App.path & "\resources\temp\common.htm", vbMaximizedFocus)
Exit Sub
errhandl:

End Sub

Private Sub mnuFoodLog_Click(Index As Integer)
     Call MakeReport(App.path & "\resources\templates\food log.htm")
End Sub

Private Sub mnuPrint_Click()
On Error Resume Next
  JournalPage1.SaveJournal
  JournalPage2.SetDay (JournalPage1.CurrentDate)
  'IoxContainer1.Visible = False
  JournalPage2.Visible = True
  
  Call Me.PrintForm
  Unload Me
End Sub

Private Sub mnuProgReport_Click()
Call MakeReport(App.path & "\resources\templates\Progress Report.htm")

End Sub

Private Sub mnuSave_Click()
On Error Resume Next
JournalPage1.SaveJournal
Call MsgBox("Page is Saved", vbOKOnly, "")
Unload Me
End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

       Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function

Private Sub mnuTodaNuts_Click()
 Call MakeReport(App.path & "\resources\templates\Today's Nutrition.htm")
End Sub

Private Sub VScroll1_Change()
JournalPage1.Top = (Height - JournalPage1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
JournalPage1.Top = (Height - JournalPage1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
