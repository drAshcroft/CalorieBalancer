VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3300
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Loading Menu Planner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public RetDialog As VbMsgBoxResult
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub ShowIt(mode As Integer, Optional Cap As String = "", Optional Label As String, _
                  Optional C1 As String, Optional C2 As String, Optional C3 As String)
On Error GoTo errhandl
If mode = 0 Then
   'Me.ControlBox = False
   Label1.Visible = True
   If Cap <> "" Then Label1.Caption = Cap
   Label2.Visible = False
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
Else
   'Me.ControlBox = True
   Label1.Visible = False
   Label2.Visible = True
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Caption = Cap
   Label2.Caption = Label
   Command1.Caption = C1
   Command2.Caption = C2
   Command3.Caption = C3
End If
errhandl:
End Sub

Private Sub Command1_Click()
On Error Resume Next
RetDialog = vbYes
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
RetDialog = vbNo
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
RetDialog = vbIgnore
Unload Me
End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

     Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
