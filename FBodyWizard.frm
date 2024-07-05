VERSION 5.00
Begin VB.Form FBodyWizard 
   Caption         =   "Body Wizard"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6975
   LinkTopic       =   "Form3"
   ScaleHeight     =   6675
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tfeet 
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox TInch 
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Text            =   " "
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TWeight 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Text            =   " "
      Top             =   960
      Width           =   2055
   End
   Begin VB.ListBox LSex 
      Height          =   840
      ItemData        =   "FBodyWizard.frx":0000
      Left            =   0
      List            =   "FBodyWizard.frx":000D
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox TBirthdate 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox tBFP 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ListBox LBodyType 
      Height          =   1230
      ItemData        =   "FBodyWizard.frx":002E
      Left            =   0
      List            =   "FBodyWizard.frx":003E
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox LWeightLossRate 
      Height          =   1230
      ItemData        =   "FBodyWizard.frx":006B
      Left            =   2160
      List            =   "FBodyWizard.frx":0081
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Low Fat"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Healthy Balance"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   6120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "High Protein"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Weight"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "ft"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "inches"
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "lbs."
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Please fill in the following information so the program can estimate your calorie needs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "Sex"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Birthdate (mm/dd/yyyy)"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Body Fat Percentage (Leave at 0 if you do not know.)"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "Body Type"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Desired Weight Change"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Plan Type"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   2655
   End
End
Attribute VB_Name = "FBodyWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sex As String
Dim BodyType As Long
Dim WLR As Single

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
LSex.Selected(0) = True
WLR = -99
BodyType = -99
End Sub

Private Sub LBodyType_Click()
Dim i As Long
For i = 1 To LBodyType.ListCount - 1
  If LBodyType.Selected(i) = True Then BodyType = i - 1
Next i
End Sub

Private Sub LSex_Click()
Dim i As Long
For i = 1 To LSex.ListCount - 1
  If LSex.Selected(i) = True Then Sex = LSex.List(i)
Next i
End Sub

Private Sub LWeightLossRate_Click()
Dim i As Long, j As Long
For i = 1 To LWeightLossRate.ListCount - 1
  If LWeightLossRate.Selected(i) = True Then j = i
Next i
WLR = j - 3

End Sub
