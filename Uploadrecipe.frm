VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Uploadrecipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Recipe"
   ClientHeight    =   9135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser RR 
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   13573
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
      Location        =   ""
   End
   Begin CalorieTracker.AutoComplete Recipename 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Recipe Name"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Uploadrecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Recipename.RecipeOnly = True
RR.Navigate2 App.Path & "\resources\temp\Recipe_Submit.htm"
End Sub

Private Sub OKButton_Click()


Me.Hide
End Sub
