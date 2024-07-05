VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FPrintPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser PrintPrev 
      Height          =   8385
      Left            =   105
      TabIndex        =   1
      Top             =   585
      Width           =   9660
      ExtentX         =   17039
      ExtentY         =   14790
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
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   105
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CHARRANGE
  cpMin As Long
  cpMax As Long
End Type

Private Type FORMATRANGE
  hdc As Long
  hdcTarget As Long
  rc As RECT
  rcPage As RECT
  chrg As CHARRANGE
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, ByVal nIndex As Long) As Long
   
Private Declare Function SendMessage Lib "USER32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, _
  ByVal wp As Long, lp As Any) As Long

      
Dim LeftOffset, LeftMargin, RightMargin, LineWidth, CC As Single

Dim TopMargin, BottomMargin
Dim BalanceHeight As Single
Dim PrintInfo

'Public Sub SetUp(Info, PrintIT As Boolean)
Public Sub SetUp(Filename As String)
     Dim Info, printit As Boolean
     PrintPrev.Navigate2 Filename
     Exit Sub
End Sub


Private Function InchesToTwips(ByVal Inches As Single) As Single
    InchesToTwips = 1440 * Inches
End Function

Private Sub CPrint_Click()

End Sub
