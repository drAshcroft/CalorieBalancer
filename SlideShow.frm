VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form SlideShow 
   Caption         =   "Introduction"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10575
      ExtentX         =   18653
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
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   7920
      Width           =   2055
   End
End
Attribute VB_Name = "SlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WB.Navigate2 App.path & "\resources\new user\how to use\quickstart.html"
End Sub

Private Sub Form_Resize()
Command1.Move (ScaleWidth - Command1.Width) / 2, ScaleHeight - Command1.Height - 100
WB.Move 0, 0, ScaleWidth, ScaleHeight - Command1.Height - 200
End Sub
