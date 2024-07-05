VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl JournalPage 
   ClientHeight    =   13500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   ScaleHeight     =   13500
   ScaleWidth      =   11415
   Begin CalorieBalance.PanelFx PanelFx2 
      Height          =   13620
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   24024
      TileHeight      =   40
      TitleCaption    =   "Journals"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      BackGroundStyle =   1
      gCTitleStart    =   32768
      gCTitleEnd      =   8454016
      gCPanelStart    =   16777215
      Begin VB.CommandButton CCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   9240
         TabIndex        =   44
         Top             =   8160
         Width           =   1575
      End
      Begin VB.CommandButton CSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   7800
         TabIndex        =   43
         Top             =   8160
         Width           =   1335
      End
      Begin CalorieBalance.PanelFx PanelFx3 
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6376
         TitleCaption    =   "Daily Info"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundEdge       =   15
         BackGroundStyle =   1
         gCTitleStart    =   33023
         gCTitleEnd      =   128
         gCPanelStart    =   8438015
         gCPanelDir      =   0
         Begin VB.TextBox TWeight 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   2880
            Width           =   2295
         End
         Begin CalorieBalance.MonthDayPicker MonthDayPicker1 
            Height          =   1815
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3201
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Morning Weight:"
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
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Current Date:"
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
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2535
         End
      End
      Begin CalorieBalance.PanelFx PanelFx4 
         Height          =   3615
         Left            =   7560
         TabIndex        =   9
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6376
         TitleCaption    =   "Day Stat's"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundEdge       =   15
         BackGroundStyle =   1
         gCTitleStart    =   33023
         gCTitleEnd      =   128
         gCPanelStart    =   8438015
         gCPanelDir      =   0
         Begin VB.Label WeightChange 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label14"
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label LTotalCals 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label13"
            Height          =   255
            Left            =   2040
            TabIndex        =   18
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label BMR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label CalsBurned 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label11"
            Height          =   255
            Left            =   2040
            TabIndex        =   16
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Weight Change:"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   2040
            Width           =   1335
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Calories:"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   3120
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Metabolic Rate:"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Calories Burned:"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label CalsEaten 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   255
            Left            =   2040
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Calories Eaten:"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
      End
      Begin CalorieBalance.PanelFx PanelFx5 
         Height          =   4335
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7646
         TitleCaption    =   "Generals"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundEdge       =   15
         BackGroundStyle =   1
         gCTitleStart    =   33023
         gCTitleEnd      =   128
         gCPanelStart    =   8438015
         gCPanelDir      =   0
         Begin VB.ListBox LHunger 
            Height          =   3180
            Left            =   3720
            TabIndex        =   28
            Top             =   840
            Width           =   975
         End
         Begin VB.ListBox LSore 
            Height          =   3180
            Left            =   2520
            TabIndex        =   27
            Top             =   840
            Width           =   975
         End
         Begin VB.ListBox LEnergy 
            Height          =   3180
            Left            =   1320
            TabIndex        =   26
            Top             =   840
            Width           =   975
         End
         Begin VB.ListBox LStress 
            Height          =   3180
            ItemData        =   "JournalPage.ctx":0000
            Left            =   120
            List            =   "JournalPage.ctx":0007
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Hungry"
            Height          =   255
            Left            =   3720
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Soreness"
            Height          =   255
            Left            =   2520
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Energy"
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Stress"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
      End
      Begin CalorieBalance.PanelFx PanelFx6 
         Height          =   3615
         Left            =   3000
         TabIndex        =   29
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6376
         TitleCaption    =   "Optional Measurements"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundEdge       =   15
         BackGroundStyle =   1
         gCTitleStart    =   33023
         gCTitleEnd      =   128
         gCPanelStart    =   8438015
         gCPanelDir      =   0
         Begin VB.TextBox TBFP 
            Height          =   285
            Left            =   2520
            TabIndex        =   40
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox THips 
            Height          =   285
            Left            =   2520
            TabIndex        =   39
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TNeck 
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TWaist 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox TWrist 
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "All Measurements should be in inches"
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Hips"
            Height          =   255
            Left            =   2520
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Body Fat Percentage"
            Height          =   255
            Left            =   2520
            TabIndex        =   33
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Wrist"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Waist "
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Neck"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1455
         End
      End
      Begin CalorieBalance.PanelFx PanelFx1 
         Height          =   3975
         Left            =   120
         TabIndex        =   41
         Top             =   9480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7011
         TitleCaption    =   "Journal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundEdge       =   15
         BackGroundStyle =   1
         gCTitleStart    =   33023
         gCTitleEnd      =   128
         gCPanelStart    =   8438015
         gCPanelDir      =   0
         Begin RichTextLib.RichTextBox JournalText 
            Height          =   3375
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   5953
            _Version        =   393217
            ScrollBars      =   3
            TextRTF         =   $"JournalPage.ctx":0014
         End
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   $"JournalPage.ctx":009F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   5760
         TabIndex        =   45
         Top             =   5280
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Username 
         BackStyle       =   0  'Transparent
         Caption         =   "Journal For "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   5895
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "JournalPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Event CloseRequest()

Public CurrentDate As Date
Public Sub SetDay(DisplayDate As Date)

    CurrentDate = DisplayDate
    Username.Caption = "Journal For " & CurrentUser.Username
    MonthDayPicker1.SetDate (DisplayDate)
   
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("Select * from dailylog where user='" & CurrentUser.Username & "' and date=#" & DateHandler.IsoDate(DisplayDate) & "#;", dbOpenDynaset)
    If (Not RS.EOF) Then
       TWeight.Text = RS("weight") & ""
       TWrist.Text = RS("wrist") & ""
       TNeck.Text = RS("neck") & ""
       TWaist.Text = RS("waist") & ""
       THips.Text = RS("hips") & ""
       tBFP.Text = RS("bfp") & ""
       
       Dim Cals As Double, Exer As Double, BMRCals As Double, TotalCals As Double, EWeightChange As Double
       
       If Not IsNull(RS("calories")) Then
            Cals = RS("calories")
       Else
            Cals = 0
       End If
       
       If Not IsNull(RS("exercise_cal")) Then
            Exer = RS("exercise_cal")
       Else
            Exer = 0
       End If
       
       If Not IsNull(RS("bmr")) Then
            BMRCals = RS("bmr")
       Else
            BMRCals = 0
       End If
       
       TotalCals = Cals - BMRCals - Exer
       EWeightChange = TotalCals / 3500
       
       CalsEaten.Caption = Round(Cals, 1)
       CalsBurned.Caption = Round(Exer, 1)
       BMR.Caption = Round(BMRCals, 1)
       LTotalCals.Caption = Round(TotalCals, 1)
       WeightChange.Caption = Round(EWeightChange, 2)
       
       LStress.Clear
       LHunger.Clear
       LSore.Clear
       LEnergy.Clear
       Dim i As Long
        For i = 1 To 10
           LStress.AddItem (i)
           LHunger.AddItem (i)
           LSore.AddItem (i)
           LEnergy.AddItem i
        Next i
        If Not IsNull(RS("energy")) Then
            LEnergy.Selected(RS("energy")) = True
        End If
        If Not IsNull(RS("stress")) Then
            LStress.Selected(RS("stress")) = True
        End If
        If Not IsNull(RS("sore")) Then
            LSore.Selected(RS("sore")) = True
        End If
        If Not IsNull(RS("hunger")) Then
            LHunger.Selected(RS("hunger")) = True
        End If
        JournalText.Text = ""
        JournalText.TextRTF = RS("Comments") & " "
        
        RS.Close
    End If
    
End Sub

Public Sub SaveJournal()
     Dim DisplayDate As Date
    DisplayDate = MonthDayPicker1.GetDate()
   
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("Select * from dailylog where user='" & CurrentUser.Username & "' and date=#" & DateHandler.IsoDate(DisplayDate) & "#;", dbOpenDynaset)
    If (RS.EOF) Then
      RS.AddNew
      RS("user") = CurrentUser.Username
      RS("date") = DisplayDate
    Else
      RS.Edit
    End If
    
    If TWeight.Text <> "" And Val(TWeight.Text) <> 0 Then RS("weight") = Val(TWeight.Text)
    If TWrist.Text <> "" Then RS("wrist") = Val(TWrist.Text)
    If TNeck.Text <> "" Then RS("neck") = Val(TNeck.Text)
    If TWaist.Text <> "" Then RS("waist") = Val(TWaist.Text)
    If THips.Text <> "" Then RS("hips") = Val(THips.Text)
    If tBFP.Text <> "" Then RS("bfp") = Val(tBFP.Text)
  
  
    Dim i As Integer
    i = GetSelected(LStress)
    If (i <> -1) Then RS("stress") = i
    
    i = GetSelected(LEnergy)
    If (i <> -1) Then RS("energy") = i
    
    i = GetSelected(LSore)
    If (i <> -1) Then RS("sore") = i
    
    i = GetSelected(LHunger)
    If (i <> -1) Then RS("hunger") = i
    
    
    If JournalText.Text <> "" Then
        RS("Comments") = JournalText.Text
    End If
    RS.Update
    RS.Close
    
End Sub

Private Sub CCancel_Click()
RaiseEvent CloseRequest
End Sub

Private Sub CSave_Click()
   SaveJournal
   RaiseEvent CloseRequest
   
End Sub

Private Function GetSelected(lb As ListBox) As Integer
   GetSelected = -1
   Dim i As Long
   For i = 0 To lb.ListCount - 1
     If lb.Selected(i) = True Then
        GetSelected = i
     End If
   Next i
End Function


Private Sub MonthDayPicker1_DateSelected(NewDate As Date)
    SetDay (NewDate)
End Sub

