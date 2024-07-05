VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin PanelFx32.PanelFx PanelFx8 
      Height          =   2055
      Left            =   6450
      TabIndex        =   19
      Top             =   2880
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3625
      TitleCaption    =   "Textured Panel"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackGroundStyle =   2
      TitleBitmap     =   "demo.frx":0000
      PanelBitmap     =   "demo.frx":0C54
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   90
         TabIndex        =   21
         Top             =   825
         Width           =   1845
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   20
         Top             =   480
         Width           =   1710
      End
   End
   Begin PanelFx32.PanelFx PanelFx7 
      Height          =   1095
      Left            =   2700
      TabIndex        =   18
      Top             =   5280
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1931
      TitleCaption    =   "Frame that can be used to move the Panels Parent"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowDraging    =   -1  'True
      BackGroundStyle =   1
      gCTitleStart    =   14737632
      gCTitleEnd      =   14588473
      gCPanelStart    =   16777215
      AllowParentDraging=   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Details Icon"
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   5790
      Width           =   1575
   End
   Begin PanelFx32.PanelFx PanelFx6 
      Height          =   2190
      Left            =   2685
      TabIndex        =   1
      Top             =   2850
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3863
      TileHeight      =   32
      TitleCaption    =   "...Moveable Panel..."
      TitleForeColor  =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      TileBackColor   =   12648447
      PanelBackColor  =   8454143
      RoundEdge       =   13
      AllowDraging    =   -1  'True
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Check1"
         Height          =   270
         Left            =   1620
         TabIndex        =   10
         Top             =   1725
         Width           =   1890
      End
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   1620
         TabIndex        =   9
         Top             =   570
         Width           =   1920
      End
      Begin VB.DirListBox Dir1 
         Height          =   765
         Left            =   105
         TabIndex        =   8
         Top             =   1005
         Width           =   1440
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   135
         TabIndex        =   7
         Top             =   570
         Width           =   1425
      End
   End
   Begin PanelFx32.PanelFx PanelFx5 
      Height          =   2730
      Left            =   2730
      TabIndex        =   6
      Top             =   75
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4815
      TitleCaption    =   "    Chat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundEdge       =   9
      BackGroundStyle =   1
      gCTitleEnd      =   16769973
      gCPanelStart    =   16769973
      TitleIcon       =   "demo.frx":18A8
      TitleIconXPos   =   5
      TitleIconYPos   =   2
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3615
         TabIndex        =   13
         Top             =   1980
         Width           =   870
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFCF7&
         Height          =   360
         Left            =   105
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1995
         Width           =   3465
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFEBCB&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "demo.frx":2582
         Top             =   690
         Width           =   3375
      End
   End
   Begin PanelFx32.PanelFx PanelFx3 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   6525
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1931
      TitleCaption    =   "Immediate"
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
   Begin PanelFx32.PanelFx PanelFx4 
      Height          =   1515
      Left            =   90
      TabIndex        =   16
      Top             =   4200
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   2672
      TileHeight      =   23
      TitleCaption    =   "Panel Collapsed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      RoundEdge       =   7
      CanCollapse     =   -1  'True
      BackGroundStyle =   1
      gCTitleStart    =   33023
      gCTitleEnd      =   12640511
      gCPanelStart    =   12640511
      gCPanelEnd      =   8438015
      gCPanelDir      =   0
   End
   Begin PanelFx32.PanelFx PanelFx2 
      Height          =   2430
      Left            =   105
      TabIndex        =   12
      Top             =   1710
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   4286
      TileHeight      =   23
      PanelBorderColor=   13160660
      TitleCaption    =   "Details"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TileBackColor   =   13160660
      TitleIcon       =   "demo.frx":25F3
      TitleIconWidth  =   16
      TitleIconHeight =   16
      TitleIconXPos   =   1
      TitleIconYPos   =   3
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensions 32 x 32"
         Height          =   210
         Left            =   270
         TabIndex        =   3
         Top             =   1860
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0023.ico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   1515
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   855
         Picture         =   "demo.frx":9AF5
         Top             =   675
         Width           =   480
      End
   End
   Begin PanelFx32.PanelFx PanelFx1 
      Height          =   1425
      Left            =   105
      TabIndex        =   15
      Top             =   120
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   2514
      TileHeight      =   23
      PanelBorderColor=   14588473
      TitleCaption    =   "Login"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundEdge       =   7
      BackGroundStyle =   1
      gCTitleStart    =   16761024
      gCPanelStart    =   16764828
      gCPanelEnd      =   14737632
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   195
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   810
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your login details below:"
         Height          =   195
         Left            =   165
         TabIndex        =   0
         Top             =   510
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    If Command2.Caption = "Hide Details Icon" Then
        Command2.Caption = "Show Details Icon"
    ElseIf Command2.Caption = "Show Details Icon" Then
        Command2.Caption = "Hide Details Icon"
    End If
    
   PanelFx2.HideTitleIcon = Not PanelFx2.HideTitleIcon
   
End Sub

Private Sub PanelFx4_TileClick()
Static bCollapse As Boolean
    bCollapse = Not bCollapse
    PanelFx4.Collapse bCollapse
End Sub


