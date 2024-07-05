VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4950
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8610
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Internet Options"
      Height          =   1695
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   4815
   End
   Begin VB.ListBox LWatch 
      Height          =   3960
      ItemData        =   "frmOptions.frx":000C
      Left            =   120
      List            =   "frmOptions.frx":000E
      MultiSelect     =   1  'Simple
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show watch columns as percent of daily value"
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.PictureBox PicOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nutrition Columns"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Watches As Collection
Public ok As Boolean


Private Sub cmdCancel_Click()
    ok = False
    'FoodGenerator1.CancelPressed
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'FoodGenerator1.OkPressed
    ok = True
    
    Me.Hide
End Sub



Private Sub Form_Load()
    Dim temp As Recordset, i As Long
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Set temp = DB.OpenRecordset("Select * from abbrev where index =-204", dbOpenDynaset)
    For i = 12 To temp.Fields.Count - 1
       LWatch.AddItem temp.Fields(i).Name
    Next i
    For i = 0 To UBound(WatchHeads)
       For j = 0 To LWatch.ListCount - 1
          If LWatch.List(j) = WatchHeads(i) Then
             LWatch.Selected(j) = True
          End If
       Next j
    Next i
End Sub


Private Sub Form_Unload(Cancel As Integer)
'FoodGenerator1.CancelPressed
End Sub




