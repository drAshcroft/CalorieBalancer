VERSION 5.00
Begin VB.Form FEasy 
   Caption         =   "Form3"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12045
   LinkTopic       =   "Form3"
   ScaleHeight     =   9135
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   12120
      Left            =   960
      ScaleHeight     =   12060
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   -3240
      Width           =   10095
      Begin CalorieBalance.AdvancedFlex AdvancedFlex1 
         Height          =   6495
         Left            =   360
         TabIndex        =   7
         Top             =   5280
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   11456
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   4080
         Width           =   2775
      End
      Begin CalorieBalance.MonthDayPicker MonthDayPicker1 
         Height          =   2655
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4683
      End
      Begin VB.Label Label4 
         Caption         =   "3. Now enter your days foods."
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4560
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "Morning Weight"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "2.  If you have weighted yourself, please enter your weight and and other information."
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   3360
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Select the day you wish to enter the foods."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   7695
      End
   End
End
Attribute VB_Name = "FEasy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
