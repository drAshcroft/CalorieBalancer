VERSION 5.00
Begin VB.UserControl NewFood 
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   ScaleHeight     =   9675
   ScaleWidth      =   10140
   Begin VB.Frame FRecipeButtons 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   9120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton CView 
         Caption         =   "View"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Instructions"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Nutrition"
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ingredients"
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   9120
      Width           =   3975
   End
End
Attribute VB_Name = "NewFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim ErrRaised As Boolean


Private Sub Command5_Click()
   
    WB.Visible = False
    WBPrint.Visible = False
    Ingred.Visible = True
    Instructions.Visible = False
    
End Sub

Public Sub SetPopUp(PopUp As Object)
  Call Ingred.SetPopUpMenu(PopUp)
End Sub


