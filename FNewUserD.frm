VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FNewUserD 
   Caption         =   "Program Setup"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11520
   LinkTopic       =   "Form3"
   ScaleHeight     =   8715
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Calculated Personal Information"
      Height          =   6855
      Index           =   4
      Left            =   8400
      TabIndex        =   51
      Top             =   1560
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Advanced"
         Height          =   375
         Left            =   4320
         TabIndex        =   71
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox TWCC 
         Height          =   285
         Left            =   3720
         TabIndex        =   65
         Text            =   "-1000"
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox TAMR 
         Height          =   285
         Left            =   3720
         TabIndex        =   64
         Text            =   "500"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TRMR 
         Height          =   285
         Left            =   3720
         TabIndex        =   63
         Text            =   "1500"
         Top             =   3450
         Width           =   1215
      End
      Begin VB.Label LBMI 
         Caption         =   "23"
         Height          =   255
         Left            =   3720
         TabIndex        =   66
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "This is the total  number of calories that you should get in a day in order to reach your weight change goals.  "
         Height          =   615
         Left            =   360
         TabIndex        =   62
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Label LTotal 
         Caption         =   "2000"
         Height          =   255
         Left            =   3720
         TabIndex        =   61
         Top             =   5760
         Width           =   975
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
         TabIndex        =   60
         Top             =   5760
         Width           =   3615
      End
      Begin VB.Label Label28 
         Caption         =   "You must reduce your eating or increase your activity to this point to change weight (500 calories per pound per week)"
         Height          =   615
         Left            =   360
         TabIndex        =   59
         Top             =   5160
         Width           =   3015
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
         TabIndex        =   58
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label Label26 
         Caption         =   "The number of calories that you burn in normal waking activities (not exercise and such)"
         Height          =   735
         Left            =   360
         TabIndex        =   57
         Top             =   4440
         Width           =   3375
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
         TabIndex        =   56
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label24 
         Caption         =   "The number of calories that you would burn laying down all day"
         Height          =   1095
         Left            =   360
         TabIndex        =   55
         Top             =   3720
         Width           =   3135
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
         TabIndex        =   54
         Top             =   3480
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         Caption         =   $"FNewUserD.frx":0000
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
         TabIndex        =   53
         Top             =   360
         Width           =   5535
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
         TabIndex        =   52
         Top             =   2880
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Index           =   3
      Left            =   6600
      TabIndex        =   30
      Top             =   960
      Width           =   5775
      Begin MSComctlLib.Slider SActivityRating 
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   6240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.OptionButton Option15 
         Caption         =   "The Sonoma Diet"
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Perricone Prescription"
         Height          =   375
         Left            =   3120
         TabIndex        =   44
         Top             =   4080
         Width           =   2055
      End
      Begin VB.OptionButton Option14 
         Caption         =   "South Beach Diet"
         Height          =   375
         Left            =   1800
         TabIndex        =   49
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Volumetrics"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   5160
         Width           =   1335
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Macrobiotic"
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   4680
         Width           =   1575
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Raw Foods"
         Height          =   375
         Left            =   1800
         TabIndex        =   47
         Top             =   5160
         Width           =   1335
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Ornish"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   4680
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "The Zone"
         Height          =   375
         Left            =   3120
         TabIndex        =   43
         Top             =   5160
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Atkins"
         Height          =   255
         Left            =   1800
         TabIndex        =   42
         Top             =   4680
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "High Protein"
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   3720
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Healthy Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Low Fat"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   3720
         Width           =   1455
      End
      Begin VB.ListBox LWeightLossRate 
         Height          =   1230
         ItemData        =   "FNewUserD.frx":0152
         Left            =   1200
         List            =   "FNewUserD.frx":0168
         TabIndex        =   31
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label33 
         Caption         =   "Construction Work/ Athlete"
         Height          =   375
         Left            =   4080
         TabIndex        =   70
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "Office Work/ Sedentary"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "Daily Activity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "Last please enter your goals and how you would like to control your weight.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label11 
         Caption         =   "Desired Weight Change"
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Diet Type"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Index           =   2
      Left            =   6600
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox LBodyType 
         Height          =   1230
         ItemData        =   "FNewUserD.frx":01E0
         Left            =   2280
         List            =   "FNewUserD.frx":01F0
         TabIndex        =   38
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox tBFP 
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Text            =   "0"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox tfeet 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox TInch 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   " "
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox TWeight 
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Text            =   " "
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ListBox LSex 
         Height          =   840
         ItemData        =   "FNewUserD.frx":021D
         Left            =   120
         List            =   "FNewUserD.frx":022A
         TabIndex        =   15
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox TBirthdate 
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Body Type"
         Height          =   375
         Left            =   2280
         TabIndex        =   40
         Top             =   5040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Body Fat Percentage (Leave at 0 if you do not know.)"
         Height          =   375
         Left            =   2280
         TabIndex        =   39
         Top             =   4080
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label17 
         Caption         =   "Height"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Weight"
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "ft"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "inches"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "lbs."
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   $"FNewUserD.frx":024B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Birthdate (mm/dd/yyyy)"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   3120
         Width           =   2055
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Next"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Index           =   1
      Left            =   6480
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Website"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox TPass 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox TUser 
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "If you wish you can also become a member of our free website.  Click here to visit the page and sign up."
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   3960
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   $"FNewUserD.frx":0312
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
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
         Left            =   2760
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
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
         Left            =   600
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton FullControlOption 
         Caption         =   "I want to set up my own diet program."
         Height          =   495
         Left            =   480
         TabIndex        =   72
         Top             =   4200
         Width           =   4815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Please setup a default 2000 calorie day."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   6120
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Please help me figure out how many calories I need in a day "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   2880
         Value           =   -1  'True
         Width           =   5175
      End
      Begin VB.Label Label19 
         Caption         =   "Advanced Setup: (Only for those who really know their nutritional needs.)"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "Easy Setup: (Recommended)"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Calorie Tracker setup.  Please provide a little information so the program can help you meet your weight control goals."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   5535
      End
   End
End
Attribute VB_Name = "FNewUserD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public UserCreated As Boolean
Public ShowSummary As Boolean
Dim sex As String
Dim BodyType As Long
Dim WLR As Single

Dim FatP As Single
Dim SugarP As Single
Dim CarbsP As Single
Dim ProteinP As Single


Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call OpenURL("http://www.caloriebalancediet.com/addmember.asp")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Long
For i = 0 To Frame1.UBound
  If Frame1(i).Visible Then
    If i = 0 Then
       
    Else
       Frame1(i - 1).Visible = True
       Frame1(i).Visible = False
    End If
  End If
Next i

End Sub

Private Sub Command3_Click()
SaveCals
Unload Me
FUserSummary.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
ShowSummary = False
If Not DoDebug Then On Error Resume Next
LSex.Selected(0) = True
WLR = -99
BodyType = -99
Me.Width = Frame1(0).Width + Frame1(0).Left + 200
Me.Height = OKButton.Top + OKButton.Height + 1000
LBodyType.Selected(2) = True
Dim i As Long
For i = 1 To Frame1.UBound
   Frame1(i).Visible = False
Next i
 FatP = 0.25
 SugarP = 0.2
 CarbsP = 0.3
 ProteinP = 0.25
End Sub

Private Sub Form_Resize()
'Me.Frame1(0).Move 0, 0
Dim i As Long
For i = 0 To Frame1.UBound
   Frame1(i).Move 0, 0
Next i

End Sub


Private Sub LBodyType_Click()
If Not DoDebug Then On Error Resume Next
Dim i As Long
For i = 1 To LBodyType.ListCount - 1
  If LBodyType.Selected(i) = True Then BodyType = i - 1
Next i
End Sub

Private Sub LSex_Click()
If Not DoDebug Then On Error Resume Next
Dim i As Long
For i = 1 To LSex.ListCount - 1
  If LSex.Selected(i) = True Then sex = LSex.List(i)
Next i
End Sub

Private Sub LWeightLossRate_Click()
If Not DoDebug Then On Error Resume Next
Dim i As Long, j As Long
For i = 1 To LWeightLossRate.ListCount - 1
  If LWeightLossRate.Selected(i) = True Then j = i
Next i
WLR = j - 3

End Sub

Private Sub OKButton_Click()
If Not DoDebug Then On Error Resume Next
Dim i As Long, j As Long
For i = 0 To Frame1.UBound
  If Frame1(i).Visible Then
        If i = 3 Then
           If WLR = -99 Then
              Call MsgBox("Please enter the desired weight change", vbOKOnly, "")
              Exit Sub
           End If
           Call DoBodyWizard
           frmLogin.RefreshList
           'Exit Sub
        ElseIf i = 4 Then
           Call SaveCals
        ElseIf i = 2 Then
             If Val(tfeet.Text) = 0 Then
                Call MsgBox("Please enter your height", vbOKOnly, "")
                Exit Sub
             End If
             If Val(TWeight.Text) = 0 Then
                Call MsgBox("Please enter your weight", vbOKOnly, "")
                Exit Sub
             End If
             Dim k As Long
             k = 0
             For j = 0 To LSex.ListCount - 1
               If LSex.Selected(j) Then k = j
             Next j
             If k = 0 Then
                Call MsgBox("Please enter your sex", vbOKOnly, "")
                Exit Sub
             End If
             Dim dd As Date
             On Error Resume Next
             dd = TBirthdate.Text
             If Err.Number <> 0 Then
                Call MsgBox("Please enter your birthdate.  It is used to determine your calorie levels.", vbOKOnly, "")
                Exit Sub
             End If
             If FullControlOption.Value = True Then
               WLR = 0
               SActivityRating.Value = 0
               Call DoBodyWizard
               Call SaveCals
               frmLogin.RefreshList
               CurrentUser.Username = TUser.Text
               ShowSummary = True
               Me.Hide
               
               
             End If
        End If
        If (Frame1.UBound >= i + 1) Then
           Frame1(i + 1).Visible = True
        Else
           Unload Me
        End If
        Frame1(i).Visible = False
        Exit Sub
  End If
Next i

End Sub
Private Sub SaveCals()

    Dim RS As Recordset
    
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
    
    
    RS.Update
    RS.Close
    Set RS = Nothing


End Sub
Private Sub DoBodyWizard()
    If Not DoDebug Then On Error Resume Next
    If TUser.Text = "" Or TPass.Text = "" Then
       MsgBox "Please enter a username and password.", vbOKOnly, ""
       Exit Sub
    End If
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("select * from profiles where user='" & TUser.Text & "';", dbOpenDynaset)
    If Not RS.EOF Then
      If TUser.Text <> CurrentUser.Username Then
          MsgBox "This username has already been used.  Please select another.", vbOKOnly, ""
          Exit Sub
      Else
         RS.Edit
      End If
    Else
      RS.AddNew
    End If
    
      RS("user") = TUser.Text
      RS("password") = TPass.Text
      CurrentUser.Username = TUser.Text
      CurrentUser.Password = TPass.Text
    frmMain.LastUser = TUser.Text
    RS.Update
    RS.Close
    Set RS = Nothing
  
    
    
    
    UserCreated = False
    
    
    
    If Not DoDebug Then On Error Resume Next
    If tfeet.Text = "" Then MsgBox "Please enter your height.", vbOKOnly, ""
    If TWeight.Text = "" Then MsgBox "Please enter your weight.", vbOKOnly, ""
    If sex = "" Then MsgBox "Please enter your sex", vbOKOnly, ""
    If BodyType = -99 Then MsgBox "Please enter your body type", vbOKOnly, ""
    If WLR = -99 Then MsgBox "Please enter your desired weight change rate", vbOKOnly, ""
    Dim dd As Date
    On Error Resume Next
    If TBirthdate.Text = "" Then MsgBox "Please enter your birthdate.", vbOKOnly, ""
    Dim sParts() As String
    sParts = Split(TBirthdate.Text, "/")
    dd = DateHandler.IsoDateStringString(sParts(0), sParts(1), sParts(2)) ' TBirthdate.Text
    If Err.Number <> 0 Then
      MsgBox "Please make sure that your birthdate is in the mm/dd/yyyy format." & vbCrLf & "07/04/1975 for example.", vbOKOnly, ""
    End If
    
    If DoDebug Then On Error GoTo 0
    
    Dim Weight As Single, Height As Single, age As Single
    Height = Val(tfeet.Text) * 12 + Val(TInch.Text)
    Weight = Val(TWeight.Text)
    age = Abs(DateDiff("d", Date, dd)) / 365
    Set RS = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    RS.Edit
    RS("height") = Height
    RS("weight") = Weight
    RS("sex") = sex
    RS("birthdate") = dd
    RS("age") = age
    RS("bfpmeasure") = "Est"
    RS("weightlossrate") = WLR
    RS("bodybuild") = BodyType
    
    Dim BMI As Single, bfp As Single, BMR As Single, Cals As Single
    BMI = Weight * 703 / Height ^ 2
    If Val(tBFP.Text) = 0 Then
       bfp = DoBFP(BodyType, BMI, "Est", sex, 0, 0, Height, 0, Weight, 0, 0)
    Else
       bfp = 0
    End If
    BMR = DoCalories(sex, age, Height, Weight, bfp)
    Cals = BMR + WLR * 500 + SActivityRating.Value / 10 * 1500 + 300
    If Cals < 1000 Then Cals = 1000
    RS("bmi") = BMI
    RS("bfp") = bfp
    RS("calories") = Cals
    RS("bmr") = BMR
    
    LBMI.Caption = Round(BMI)
    TRMR.Text = Round(BMR)
    TAMR.Text = Round(SActivityRating.Value / 10 * 1500 + 300)
    TWCC.Text = WLR * 500
    
    
          RS("fat") = FatP
          RS("sugar") = SugarP
          RS("carbs") = CarbsP
          RS("protein") = ProteinP
          RS("fiber") = 25
    RS("calpound") = 3500
    RS("otherwatches") = "Calcium,Iron,Sodium,Vitamin C,Cholesterol"
    RS.Update
    RS.Close
    Set RS = Nothing
    
    Dim rs2 As Recordset
    Set RS = DB.OpenRecordset("select max(index) as maxit from ideals;", dbOpenDynaset)
    Dim maxi As Long
    maxi = RS("maxit") + 1
    Set RS = DB.OpenRecordset("select * from ideals where user='AnyUser';", dbOpenDynaset)
    Set rs2 = DB.OpenRecordset("select * from ideals where user='" & TUser.Text & "';", dbOpenDynaset)
    If rs2.EOF Then
       rs2.AddNew
       Dim i As Long
       For i = 0 To RS.Fields.Count - 1
         rs2(i) = RS(i)
       Next i
       rs2("user") = TUser.Text
       rs2("index") = maxi
       rs2("fat") = FatP * Cals / 9
       rs2("sugar") = SugarP * Cals / 4
       rs2("carbs") = CarbsP * Cals / 4
       rs2("protein") = ProteinP * Cals / 4
       rs2.Update
    End If

End Sub



Public Sub SetDefault()

If Not DoDebug Then On Error Resume Next
Dim RS As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim MaxIt As Long
Set RS = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
Set rs2 = DB.OpenRecordset("select * from ideals where user='anyuser';", dbOpenDynaset)
If RS.EOF Then
  Set rs3 = DB.OpenRecordset("select max(index) as maxit from ideals;", dbOpenDynaset)
  RS.AddNew
  MaxIt = rs3("maxit") + 1
  
  Set rs3 = Nothing
Else
  RS.Edit
  MaxIt = RS("index")
End If
Dim i As Long
For i = 1 To rs2.Fields.Count - 1
  RS(i) = rs2(i)
Next i
If RS("index") <> MaxIt Then RS("index") = MaxIt
RS("user") = CurrentUser.Username
RS.Update
RS.Close
Set RS = Nothing
rs2.Close
Set rs2 = Nothing



     Set RS = DB.OpenRecordset("Select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
     If RS.EOF Then
         ' Set rs3 = DB.OpenRecordset("select max(index) as maxit from profiles;", dbOpenDynaset)
        RS.AddNew
        'Maxit = rs3("maxit") + 1
        'rs3.Close
        'Set rs3 = Nothing
     Else
        RS.Edit
     End If
     RS("fat") = 0.25
     RS("protein") = 0.25
     RS("carbs") = 0.3
     RS("sugar") = 0.2
     RS("calpound") = 3500
     RS("calories") = 2000
     RS("age") = 25
     RS("birthdate") = DateAdd("yyyy", -25, Date)
     RS("sex") = "female"
     RS("bfpmeasure") = "Est"
     RS("weight") = 130
     RS("height") = 64
     RS("bmi") = 25
     RS("bfp") = 25
     RS("bmr") = 2000
     RS("calories") = 2000
     
     RS("otherwatches") = "Calcium,Iron,Sodium,Vitamin C,Cholesterol"
    ' rs("index") = Maxit
     RS.Update
     Set RS = Nothing
End Sub


Private Sub Option10_Click()
     FatP = 0.1
      SugarP = 0.2
      CarbsP = 0.5
      ProteinP = 0.2
End Sub

Private Sub Option11_Click()
      FatP = 0.15
      SugarP = 0.13
      CarbsP = 0.6
      ProteinP = 0.12
End Sub

Private Sub Option12_Click()
     FatP = 0.3
      SugarP = 0.2
      CarbsP = 0.2
      ProteinP = 0.3
End Sub

Private Sub Option13_Click()
     FatP = 0.3
      SugarP = 0.15
      CarbsP = 0.4
      ProteinP = 0.15
End Sub

Private Sub Option14_Click()
     FatP = 0.3
      SugarP = 0.2
      CarbsP = 0.2
      ProteinP = 0.3
End Sub

Private Sub Option15_Click()
      FatP = 0.3
      SugarP = 0.1
      CarbsP = 0.2
      ProteinP = 0.4
End Sub

Private Sub Option3_Click()

      FatP = 0.3
      SugarP = 0.2
      CarbsP = 0.2
      ProteinP = 0.3
End Sub

Private Sub Option5_Click()
 FatP = 0.25
 SugarP = 0.2
 CarbsP = 0.3
 ProteinP = 0.25


End Sub

Private Sub Option6_Click()
 FatP = 0.2
 SugarP = 0.2
 CarbsP = 0.35
 ProteinP = 0.25
End Sub

Private Sub Option7_Click()
FatP = 0.5
 SugarP = 0.05
 CarbsP = 0.15
 ProteinP = 0.3
End Sub

Private Sub Option8_Click()
 FatP = 0.3
 SugarP = 0.1
 CarbsP = 0.3
 ProteinP = 0.3
End Sub

Private Sub Option9_Click()
 FatP = 0.3
 SugarP = 0.2
 CarbsP = 0.2
 ProteinP = 0.3
End Sub

Private Sub Slider1_Click()

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
