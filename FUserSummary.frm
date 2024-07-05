VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FUserSummary 
   Caption         =   "User Summary"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   ScaleHeight     =   9195
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10800
      TabIndex        =   76
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   10800
      TabIndex        =   75
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   8415
      Left            =   240
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox TFiber 
         Height          =   285
         Left            =   6600
         TabIndex        =   79
         Text            =   "25"
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox pProtein 
         Height          =   285
         Left            =   3120
         TabIndex        =   73
         Text            =   "10"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox PSugar 
         Height          =   285
         Left            =   3120
         TabIndex        =   72
         Text            =   "10"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox PCarbs 
         Height          =   285
         Left            =   3120
         TabIndex        =   71
         Text            =   "10"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox pFat 
         Height          =   285
         Left            =   3120
         TabIndex        =   70
         Text            =   "10"
         Top             =   1320
         Width           =   975
      End
      Begin CalorieBalance.PieChart Balance 
         Height          =   4080
         Left            =   120
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3840
         Width           =   4680
         _ExtentX        =   4022
         _ExtentY        =   3387
         MaskPicture     =   "FUserSummary.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Blend           =   5
      End
      Begin VB.Label Label13 
         Caption         =   "Fiber (g)"
         Height          =   375
         Left            =   5040
         TabIndex        =   78
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5040
         TabIndex        =   77
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   4920
         TabIndex        =   74
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label10 
         Caption         =   "% Calories From Sugar                        "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   69
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "% Calories From Carbs                                "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "% Calories From Protein                            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "% Calories From Fat                                 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label LTotCalsDisp 
         Caption         =   "2000"
         Height          =   255
         Left            =   3120
         TabIndex        =   65
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Total Calories"
         Height          =   375
         Left            =   1440
         TabIndex        =   64
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   34
         Left            =   7560
         TabIndex        =   62
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   33
         Left            =   7560
         TabIndex        =   61
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   32
         Left            =   7560
         TabIndex        =   60
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   31
         Left            =   7560
         TabIndex        =   59
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   30
         Left            =   7560
         TabIndex        =   58
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   29
         Left            =   7560
         TabIndex        =   57
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   28
         Left            =   7560
         TabIndex        =   56
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   27
         Left            =   5760
         TabIndex        =   55
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   26
         Left            =   5760
         TabIndex        =   54
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   25
         Left            =   5760
         TabIndex        =   53
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   24
         Left            =   5760
         TabIndex        =   52
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   23
         Left            =   5760
         TabIndex        =   51
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   22
         Left            =   5760
         TabIndex        =   50
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   21
         Left            =   5760
         TabIndex        =   49
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   20
         Left            =   4080
         TabIndex        =   48
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   19
         Left            =   4080
         TabIndex        =   47
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   18
         Left            =   4080
         TabIndex        =   46
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   17
         Left            =   4080
         TabIndex        =   45
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   16
         Left            =   4080
         TabIndex        =   44
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   15
         Left            =   4080
         TabIndex        =   43
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   42
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   41
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   13
         Left            =   2400
         TabIndex        =   40
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   39
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   11
         Left            =   2400
         TabIndex        =   38
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   10
         Left            =   2400
         TabIndex        =   37
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   36
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   35
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   33
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   31
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox CWatch 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   360
         TabIndex        =   18
         Top             =   3480
         Width           =   5055
         Begin VB.VScrollBar VScroll1 
            Height          =   4695
            Left            =   4080
            TabIndex        =   20
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   4095
            TabIndex        =   19
            Top             =   480
            Width           =   4095
            Begin VB.TextBox TVitValue 
               Height          =   285
               Index           =   0
               Left            =   2520
               TabIndex        =   24
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label LVitName 
               Caption         =   "LVname                                        "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   60
               Visible         =   0   'False
               Width           =   2895
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Target Value"
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Vitamin Name"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label RDA 
         Caption         =   "RDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Watch Nutrients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "These are the nutrients that are shown on the main page.  You can select up to 5 for this display."
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label Label3 
         Caption         =   $"FUserSummary.frx":29542
         Height          =   2655
         Left            =   5520
         TabIndex        =   25
         Top             =   4560
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8415
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.TextBox TRMR 
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Text            =   "1500"
         Top             =   3450
         Width           =   1215
      End
      Begin VB.TextBox TAMR 
         Height          =   285
         Left            =   3720
         TabIndex        =   2
         Text            =   "500"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TWCC 
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         Text            =   "-1000"
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "You must reduce your eating or increase your activity to this point to change weight (500 calories per pound per week)"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Label Label23 
         Caption         =   $"FUserSummary.frx":29652
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label22 
         Caption         =   "Resting Metabolism                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         Caption         =   "The number of calories that you would burn laying down all day"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label25 
         Caption         =   "Activity Metabolism                                                     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label26 
         Caption         =   "The number of calories that you burn in normal waking activities (not exercise and such)"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label Label27 
         Caption         =   "Weight Loss Calories                                                "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label Label29 
         Caption         =   "Total Daily Calories                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5760
         Width           =   3615
      End
      Begin VB.Label LTotal 
         Caption         =   "2000"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "This is the total  number of calories that you should get in a day in order to reach your weight change goals.  "
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Label LBMI 
         Caption         =   "23"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Body Mass Index                                                              "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   4335
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8895
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   15690
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Calories"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vitamins"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro Nutrient Balance"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FUserSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Private NutrientList As New Collection

Private Sub Command1_Click()
  SaveUser
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub SaveUser()
    Dim RS As Recordset
    Dim ideals As Recordset
    If DoDebug Then
        On Error GoTo 0
    Else
        On Error Resume Next
    End If
    
    Set RS = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    
    RS.Edit
    
    Dim BMR As Single, Cals As Single
    
    BMR = Val(TRMR.Text) + Val(TAMR.Text)
    Cals = BMR + Val(TWCC.Text)
    If Cals < 1000 Then Cals = 1000
    
    RS("weightlossrate") = Val(TWCC.Text) / 500
    RS("calories") = Cals
    RS("bmr") = BMR
    Dim junk As String
    junk = ""
    For i = 0 To CWatch.UBound
      If CWatch(i).Value = 1 Then
         If i = 0 Then
            junk = junk & CWatch(i).Caption
         Else
            junk = junk & "," & CWatch(i).Caption
         End If
      End If
    Next i
    RS("otherwatches") = junk
    
    RS("fat") = Val(pFat.Text) / 100 ' = Round(RS("fat") * 100, 1)
    RS("sugar") = Val(PSugar.Text) / 100
    RS("carbs") = Val(PCarbs.Text) / 100
    RS("protein") = Val(pProtein.Text) / 100
    RS("fiber") = Val(TFiber.Text)
    
    
    RS.Update
    RS.Close
    Set RS = Nothing
    
    Set ideals = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    ideals.Edit
    For i = 0 To LVitName.UBound
      ideals(LVitName(i).Caption) = Val(TVitValue(i).Text)
    
    Next i
    ideals("fat") = Val(pFat.Text) / 100 * Cals / 9 ' = Round(RS("fat") * 100, 1)
    ideals("sugar") = Val(PSugar.Text) / 100 * Cals / 4
    ideals("carbs") = Val(PCarbs.Text) / 100 * Cals / 4
    ideals("protein") = Val(pProtein.Text) / 100 * Cals / 4
    ideals("fiber") = Val(TFiber.Text)
    ideals.Update
    ideals.Close
    Set ideals = Nothing
End Sub

Private Sub PCarbs_Change()
  DoBalance
End Sub

Private Sub pFat_Change()
 DoBalance
End Sub
Private Sub DoBalance()
  Dim Calories As Single, f As Single, s As Single, c As Single, p As Single, fb As Single
  Calories = Val(LTotCalsDisp.Caption)
  f = Val(pFat.Text) * Calories / 9 / 100
  s = Val(PSugar.Text) * Calories / 4 / 100
  c = Val(PCarbs.Text) * Calories / 4 / 100
  p = Val(pProtein.Text) * Calories / 4 / 100
  fb = Val(TFiber.Text) * Calories / 4 / 100
  Call Module1.FigurePercentages(Balance, Calories, f, s, c, p, fb)
End Sub

Private Sub pProtein_Change()
  DoBalance
End Sub

Private Sub PSugar_Change()
  DoBalance
End Sub

Private Sub TabStrip1_Click()
  Dim i As Long
  i = TabStrip1.SelectedItem.Index
  If (i = 1) Then
           Frame1(0).Visible = True
           Frame2.Visible = False
           Frame4.Visible = False
  ElseIf (i = 2) Then
           Frame1(0).Visible = False
           Frame2.Visible = True
           Frame4.Visible = False
  Else
           Frame1(0).Visible = False
           Frame2.Visible = False
           Frame4.Visible = True

           
  End If
   
End Sub

Private Sub TAMR_Change()
   doCalorieSum
End Sub

Private Sub TRMR_Change()
   doCalorieSum
End Sub
Private Sub doCalorieSum()
On Error Resume Next
   Dim sum As Double
   sum = 0
   sum = Val(TRMR.Text) + Val(TAMR.Text) + Val(TWCC.Text)
   LTotal.Caption = sum

End Sub

Private Sub TWCC_Change()
   doCalorieSum
End Sub
Private Sub Form_Load()
    Dim RS As Recordset
    Dim ideals  As Recordset
    Set RS = DB.OpenRecordset("Select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    Set ideals = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    
    LBMI.Caption = Round(RS("bmi"), 1)
    Dim BMR As Double, bfp As Double
    Dim Weight As Single, Height As Single, age As Single, sex As String, dd As Date
    Height = RS("height")
    Weight = RS("weight")
    sex = RS("Sex")
    dd = RS("birthdate")
    age = Abs(DateDiff("d", Date, dd)) / 365
    bfp = RS("bfp")
    BMR = DoCalories(sex, age, Height, Weight, bfp)
    TRMR.Text = Round(BMR)
    TAMR.Text = Round(RS("bmr") - BMR)
    TWCC.Text = Round(RS("weightlossrate") * 500)
    
    
    LTotCalsDisp.Caption = Round(RS("calories"))
   
    pFat.Text = Round(RS("fat") * 100, 1)
    PSugar.Text = Round(RS("sugar") * 100, 1)
    PCarbs.Text = Round(RS("carbs") * 100, 1)
    pProtein.Text = Round(RS("protein") * 100, 1)
    TFiber.Text = RS("fiber")
    
    'NutrientList.Add "Fat"
    NutrientList.Add "Saturated Fat"
    NutrientList.Add "Monounsaturated Fat"
    NutrientList.Add "Polyunsaturated Fat"
    NutrientList.Add "Trans fat"
    NutrientList.Add "Cholesterol"
    NutrientList.Add "Sodium"
    NutrientList.Add "Carbohydrates"
    'NutrientList.Add "Fiber"
    'NutrientList.Add "Sugars"
    'NutrientList.Add "Protein"
    NutrientList.Add "Vitamin A"
    NutrientList.Add "Vitamin C"
    NutrientList.Add "Calcium"
    NutrientList.Add "Iron"
    
    NutrientList.Add "Magnesium"
    NutrientList.Add "Phosphorus"
    NutrientList.Add "Potassium"
    
    NutrientList.Add "Zinc"
    NutrientList.Add "Copper"
    NutrientList.Add "Manganese"
    NutrientList.Add "Selenium"
    
    NutrientList.Add "Thiamin"
    NutrientList.Add "Riboflavin"
    NutrientList.Add "Niacin"
    NutrientList.Add "Pantothenic acid"
    NutrientList.Add "Vitamin B6"
    NutrientList.Add "Folate"
    NutrientList.Add "Vitamin B12"
    
    NutrientList.Add "Retinol"
    NutrientList.Add "Vitamin D"
    NutrientList.Add "Vitamin E"
    NutrientList.Add "Vitamin K"
    NutrientList.Add "Alpha-carotene"
    For i = 0 To CWatch.UBound
       CWatch(i).Visible = False
    Next
    Dim OtherWatchs As String
    OtherWatchs = RS("otherwatches")
    For i = 1 To NutrientList.Count
       CWatch(i).Caption = NutrientList(i)
       CWatch(i).Visible = True
       If InStr(1, OtherWatchs, NutrientList(i), vbTextCompare) <> 0 Then
            CWatch(i).Value = 1
       End If
    Next i
    Dim Units As Recordset
    Set Units = DB.OpenRecordset("Select * from units;", dbOpenDynaset)
    Dim cc As Long, junk As String
    On Error Resume Next
    For i = 2 To ideals.Fields.Count - 1
       junk = LCase(ideals.Fields(i).Name)
       If (junk <> "fiber" And junk <> "carbs" And junk <> "sugar" And junk <> "fat" And junk <> "protein") Then
            Load LVitName(cc)
            LVitName(cc).Visible = True
            LVitName(cc).Caption = ideals.Fields(i).Name
            LVitName(cc).Caption = LVitName(i).Caption & " (" & Units(ideals.Fields(i).Name) & ")"
            LVitName(cc).Caption = LVitName(i).Caption & "                                                         ."
            LVitName(cc).Caption = Left$(LVitName(i).Caption, 55)
            Load TVitValue(cc)
            TVitValue(cc).Visible = True
            TVitValue(cc).Text = ideals(i)
            LVitName(cc).Top = (TVitValue(0).Height + 100) * cc + LVitName(0).Top
            TVitValue(cc).Top = (TVitValue(0).Height + 100) * cc
            cc = cc + 1
        End If
    Next i
    Picture1.Height = (TVitValue(0).Height + 100) * cc
End Sub


Private Sub VScroll1_Change()
Picture1.Top = (Frame3.Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = (Frame3.Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
