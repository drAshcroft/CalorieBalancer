VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmJournal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Journal"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Daily Journal"
   Begin SHDocVwCtl.WebBrowser Report 
      Height          =   6375
      Left            =   6720
      TabIndex        =   69
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   11245
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
   Begin VB.CommandButton CReport 
      Caption         =   "Run Report"
      Height          =   495
      Left            =   9840
      TabIndex        =   68
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame FNutrient 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      TabIndex        =   61
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ListBox LNutrients 
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Nutrient"
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame FToFrom 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   495
      Left            =   6720
      TabIndex        =   56
      Top             =   6720
      Width           =   3015
      Begin VB.ListBox LFrom 
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox LTo 
         Height          =   255
         Left            =   1320
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "From:"
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "To:"
         Height          =   255
         Left            =   1320
         TabIndex        =   59
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Show Weight loss with Nutrient"
      Height          =   495
      Left            =   13680
      TabIndex        =   55
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6615
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   3720
         TabIndex        =   49
         Top             =   360
         Width           =   2415
         Begin VB.Label LBFPInfo 
            Height          =   975
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label LBfp 
            Caption         =   " "
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Your Current Body Fat Percentage is:"
            Height          =   495
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame FClothTape 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1695
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   3015
         Begin VB.Frame FMale 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   975
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
            Begin VB.TextBox TNeckM 
               Height          =   285
               Left            =   1200
               TabIndex        =   46
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TWaistM 
               Height          =   285
               Left            =   1200
               TabIndex        =   45
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "Neck (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Waist (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.Frame FWoman 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1575
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   2775
            Begin VB.TextBox TForearm 
               Height          =   285
               Left            =   1200
               TabIndex        =   39
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox TWrist 
               Height          =   285
               Left            =   1200
               TabIndex        =   38
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TNeckW 
               Height          =   285
               Left            =   1200
               TabIndex        =   37
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox THips 
               Height          =   285
               Left            =   1200
               TabIndex        =   36
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label15 
               Caption         =   "Forearm (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1200
               Width           =   1050
            End
            Begin VB.Label Label14 
               Caption         =   "Wrist (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label13 
               Caption         =   "Neck (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Hips (in)"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   120
               Width           =   855
            End
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Calculate"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Frame FOTher 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox TDBFP 
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Body Fat Percentage"
            Height          =   495
            Left            =   0
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Body Fat Measurement Method"
         Height          =   975
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton Option2 
            Caption         =   "Machine Measured"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cloth Tape"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Journal View"
      Height          =   495
      Left            =   840
      TabIndex        =   25
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame2 
         Caption         =   "Results"
         Height          =   3375
         Left            =   3720
         TabIndex        =   15
         Top             =   0
         Width           =   2895
         Begin VB.Label LBMR 
            Caption         =   "xxxxxx"
            Height          =   255
            Left            =   2160
            TabIndex        =   54
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Resting Calories"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   2760
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Total 
            Caption         =   "xxxxxx"
            Height          =   255
            Left            =   2160
            TabIndex        =   21
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Sum Total"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Burned 
            Caption         =   "xxxxxx"
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Calories Burned"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Consumed 
            Caption         =   "xxxxxx"
            Height          =   255
            Left            =   2160
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Calories Consumed"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Feedback (1 = Lots / 10 = None)"
         Height          =   2655
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   3375
         Begin VB.ListBox LSore 
            Height          =   2010
            ItemData        =   "frmAbout.frx":0000
            Left            =   2640
            List            =   "frmAbout.frx":0002
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.ListBox LHunger 
            Height          =   2010
            ItemData        =   "frmAbout.frx":0004
            Left            =   1800
            List            =   "frmAbout.frx":0006
            TabIndex        =   66
            Top             =   480
            Width           =   615
         End
         Begin VB.ListBox LStress 
            Height          =   2010
            ItemData        =   "frmAbout.frx":0008
            Left            =   960
            List            =   "frmAbout.frx":000A
            TabIndex        =   65
            Top             =   480
            Width           =   615
         End
         Begin VB.ListBox LEnergy 
            Height          =   2010
            ItemData        =   "frmAbout.frx":000C
            Left            =   120
            List            =   "frmAbout.frx":000E
            TabIndex        =   64
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Soreness"
            Height          =   255
            Left            =   2520
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Hunger"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Stress"
            Height          =   255
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Energy"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TWeight 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox Journal 
         Height          =   3135
         Left            =   0
         TabIndex        =   22
         Top             =   3480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmAbout.frx":0010
      End
      Begin VB.Label Label5 
         Caption         =   "Morning Weight"
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton CRBPF 
      Caption         =   "Record Body Fat"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Weight Loss"
      Height          =   495
      Left            =   12480
      TabIndex        =   5
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Weight"
      Height          =   495
      Left            =   11160
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin CalorieTracker.uGraphSurface GS 
      Height          =   6375
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      _extentx        =   14843
      _extenty        =   11456
      font            =   "frmAbout.frx":009B
      borderstyle     =   3
      drawstyle       =   101
      drawwidth       =   121
      scaleheight     =   -10
      scaletop        =   10
      scalewidth      =   11.001
      scalemode       =   0
      axisfont        =   "frmAbout.frx":00CB
      captionfont     =   "frmAbout.frx":00FB
      numberfont      =   "frmAbout.frx":0129
      yaxisname       =   "Calories  "
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Progress"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Property Variables:
Dim m_Sore As Long
Dim m_Energy As Long
Dim m_Stress As Long
Dim m_Hunger As Long
Dim m_CaloriesGoal As Single
Dim m_CaloriesSum As Single
Dim Fields()
Dim BFP As Single
Dim sex As Boolean
Dim m_BMR As Single
Dim m_CPP As Single
Dim Slopes(8) As Single, Bs(8) As Single

Dim UserChange As Boolean

Public Sub ResetMe(today As Date)
   Frame3.Visible = True
   Frame4.Visible = False
   Command6.Visible = False
 


   Dim temp As Recordset, Junk As String, junk2 As String
   Dim junkL As Long
   Dim Yesterday As Date
   Set temp = DB.OpenRecordset("Select * from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
   If LCase$(temp.Fields("Sex")) = "male" Then sex = True Else sex = False
   If sex Then
        FMale.Visible = sex
        LBFPInfo.Caption = "The acceptable range is from 15% to 25%"
   Else
        
        FWoman.Visible = Not sex
        LBFPInfo.Caption = "The acceptable range is from 20% to 30%"
   End If
   junk2 = temp.Fields("BMR")
   
   LBMR.Caption = junk2

   
   
   TDate.Text = today
   Yesterday = DateAdd("d", -1, TDate.Text)
   Set temp = DB.OpenRecordset("Select * from dailylog where ((User = '" & CurrentUser.Username & "') and (date = #" & today & "#));", dbOpenDynaset)
   If temp.RecordCount <> 0 Then
     Junk = temp.Fields("Exercise_cal")
     Burned.Caption = Val(Junk)
   Else
     Burned.Caption = 0
   End If
   
   On Error Resume Next
   temp.Close
   Set temp = Nothing
   Set temp = DB.OpenRecordset("Select * from dailylog where ((User = """ & CurrentUser.Username & """) and (date = #" & TDate.Text & "#));", dbOpenDynaset)
   If temp.RecordCount <> 0 Then
     Junk = temp.Fields("calories")
     Consumed.Caption = Round(Val(Junk), 0)
     
     Junk = temp.Fields("Weight")
     TWeight.Text = Junk
     RecordWeight = Val(Junk)
     
     junkL = Val(temp.Fields("stress"))
     If junkL = 0 Then junkL = 4
     LStress.Selected(junkL) = True
     
     junkL = Val(temp.Fields("sore"))
     If junkL = 0 Then junkL = 4
     LSore.Selected(junkL) = True
     
     junkL = temp.Fields("Hunger")
     If junkL = 0 Then junkL = 4
     LHunger.Selected(junkL) = True
     
     junkL = temp.Fields("Energy")
     If junkL = 0 Then junkL = 4
     LEnergy.Selected(junkL) = True
     
     Junk = temp.Fields("Comments")
     Journal.TextRTF = Junk
     
   End If
   temp.Close
   Set temp = Nothing
   Junk = Nutmaxes("Calories")
   'GoalCalories.Caption = junk
   
   On Error GoTo 0
Dim MinWeek As Long, MaxWeek As Long, I As Long
Set temp = DB.OpenRecordset("Select * from WeekLog Where User = '" & CurrentUser.Username & "' order by week asc;", dbOpenDynaset)
Erase Fields
ReDim Fields(8, 1)
If Not temp Is Nothing And Not temp.EOF Then
  Dim minDate As Date, D As Date
  Dim firstSunday As Date
  minDate = temp.Fields("week")
  I = Weekday(minDate, vbSunday)
  firstSunday = DateAdd("d", -1 * I, minDate)



While Not temp.EOF
   D = temp.Fields("week")
   I = DateDiff("d", minDate, D)
   ReDim Preserve Fields(8, I)
   Fields(0, I) = temp("Week")
   Fields(1, I) = temp.Fields("Calories")
   Fields(2, I) = temp.Fields("Exercise")
   Fields(3, I) = temp.Fields("RWeightloss")
   Fields(4, I) = temp("TWeightloss")
   Fields(5, I) = temp.Fields("Weight")
   Fields(6, I) = temp.Fields("BMR")
   Fields(7, I) = temp.Fields("BFP")
   Fields(8, I) = 1
   
   temp.MoveNext
Wend
End If
temp.Close
Set temp = Nothing

If today = Date Then
   Dim BMR As Single, CPP As Single
   Call ShowCaloriesToWeightLoss(0, True, BMR, CPP)

   Set temp = DB.OpenRecordset("Select * from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
   If BMR = 0 Then
      Call ShowCaloriesToWeightLoss(0, False, BMR, CPP)
   End If
   temp.Edit
   temp("bmr") = BMR
   temp("calPound") = CPP
   temp.Update
   temp.Close

   Set temp = Nothing
End If


End Sub


Private Sub ShowCaloriesToWeightLoss(Display As Integer, Optional GetBMR As Boolean = False, Optional BMR As Single, Optional CPP As Single)
If Not GetBMR Then
  GS.Reset
  GS.CaptionFont.Size = 13
  GS.ShowLegend = False
If Display = 0 Then
  GS.CaptionName = "Weight Loss with Exercise and Calories"
  GS.XAxisName = "Calories - Exercise Calories Daily"
ElseIf Display = 1 Then
  GS.CaptionName = "Weight Loss with Calories Consumed"
  GS.XAxisName = "Calories Daily"
ElseIf Display = 2 Then
  GS.CaptionName = "Weight Loss with Exercise"
  GS.XAxisName = "Exercise Calories Burned Daily"
ElseIf Display = 3 Then
  GS.CaptionName = "Effect of Fat on Diet"
  GS.XAxisName = "Percent of Fat Calories Daily"
ElseIf Display = 4 Then
  GS.CaptionName = "Effect of Sugar on Diet"
  GS.XAxisName = "Percent of Sugar Calories Daily"
ElseIf Display = 5 Then
  GS.CaptionName = "Effect of Protein on Diet"
  GS.XAxisName = "Percent of Protein Calories Daily"
ElseIf Display = 6 Then
  GS.CaptionName = "Effect of Fiber on Diet"
  GS.XAxisName = "Grams of Fiber Daily"
ElseIf Display = 7 Then
  GS.CaptionName = "Effect of Carbs on Diet"
  GS.XAxisName = "Percent of Carbs Calories Daily"
ElseIf Display = 8 Then
  GS.CaptionName = "Effect of All Nutrients on Diet"
  GS.XAxisName = "Percent Daily"
  GS.ShowLegend = True
End If

If Display > 2 Then
    GS.AutoScaleY = False
    GS.SetYMax 0, -5
    GS.SetYMax 1, 5
Else
    GS.AutoScaleY = False
    GS.SetYMax 0, -10
    GS.SetYMax 1, 10

End If
    
    GS.MonthAxis = False


End If

Dim T() As Single, L
Dim I As Long, CC As Long
If Display <= 2 And Not GetBMR Then
    CC = 0
    For I = 0 To UBound(Fields, 2)
        If Fields(8, I) = 1 Then
            ReDim Preserve T(1, CC)
            If Display = 0 Then
               T(0, CC) = (Fields(1, I) - Fields(2, I)) / 7
            ElseIf Display = 1 Then
                T(0, CC) = (Fields(1, I)) / 7
            ElseIf Display = 2 Then
                T(0, CC) = (Fields(2, I)) / 7
            End If
            T(1, CC) = Fields(3, I)
            CC = CC + 1
        End If
    Next I
End If


Dim temp As Recordset
Dim T2() As Single, j As Long
If GetBMR Then
    Dim WeekAgo As Date
    Dim today As Date
    today = TDate.Text
    WeekAgo = DateAdd("d", -14, today)
    Set temp = DB.OpenRecordset("Select * from dailylog where ((user ='" & CurrentUser.Username & "') and (date >= #" & WeekAgo & "#) and (date< #" & today & "#));", dbOpenDynaset)
Else
   Set temp = DB.OpenRecordset("Select * from dailylog where (user ='" & CurrentUser.Username & "') order by date;", dbOpenDynaset)
End If
CC = 0
On Error Resume Next
While Not temp.EOF
   ReDim Preserve T2(8, CC)
   T2(2, CC) = temp.Fields("Exercise_Cal")
   T2(1, CC) = temp.Fields("Weight")
   T2(0, CC) = temp.Fields("Calories")
   If Display > 2 Then
     T2(3, CC) = temp.Fields("Fat") * 9 / T2(0, CC) * 100
     T2(4, CC) = temp.Fields("Sugar") * 4 / T2(0, CC) * 100
     T2(5, CC) = temp.Fields("Protein") * 4 / T2(0, CC) * 100
     T2(6, CC) = temp.Fields("Fiber")
     T2(7, CC) = temp.Fields("Carbs") * 4 / T2(0, CC) * 100
     T2(8, CC) = temp("Bmr")
   End If
   temp.MoveNext
  CC = CC + 1
Wend
temp.Close
Set temp = Nothing

Dim MultiLine As Boolean
If Display = 8 Then
  MultiLine = True
  Display = 3
Else
  MultiLine = False
End If

Do

If Display = 6 And MultiLine Then Display = 7
If Display = 4 And MultiLine Then Display = 5

Dim T3() As Single, Con As Single, Ex As Single
Dim TotCal As Single
CC = 0
For I = 2 To UBound(T2, 2)
  If T2(1, I - 2) * T2(1, I - 1) * T2(1, I) * T2(1, I + 1) <> 0 Then
    ReDim Preserve T3(1, CC)
    
    T3(1, CC) = (T2(1, I + 1) - T2(1, I - 2)) / 3 * 7
    
    Con = (T2(0, I - 2) + T2(0, I - 1) + T2(0, I)) / 3
    Ex = ((T2(2, I - 2) + T2(2, I - 1) + T2(2, I))) / 3
    If Display = 0 Then
        T3(0, CC) = Con - Ex
    ElseIf Display = 1 Then
        T3(0, CC) = Con
    ElseIf Display = 2 Then
        TotCal = 0
        For j = -2 To 0
            TotCal = TotCal + T2(0, I + j)
        Next j
        TotCal = TotCal / 3
        
        T3(0, CC) = Ex
        T3(1, CC) = T3(1, CC) - (TotCal - T2(8, I)) / CurrentUser.CalPound * 7
    Else
        TotCal = 0
        For j = -2 To 0
            TotCal = TotCal + T2(0, I + j) - T2(2, I + j)
        Next j
        TotCal = TotCal / 3

        T3(0, CC) = (T2(Display, I - 2) + T2(Display, I - 1) + T2(Display, I)) / 3
        T3(1, CC) = T3(1, CC) - (TotCal - T2(8, I)) / CurrentUser.CalPound * 7
    End If
    CC = CC + 1
  End If
Next I
Dim T4() As Single
CC = 0
ReDim T4(1, UBound(T3, 2))

For I = 0 To UBound(T3, 2) - 1
  If Not (T3(0, I) = 0 And T3(1, I) = 0) Then
     ReDim Preserve T4(1, CC)
     
     T4(0, CC) = T3(0, I)
     T4(1, CC) = T3(1, I)
     
     CC = CC + 1
  End If
Next I
Dim slope As Double, B As Double
If Display <= 2 And Not GetBMR Then
    For I = 0 To UBound(T, 2)
         ReDim Preserve T4(1, CC)
        T4(0, CC) = T(0, I)
        T4(1, CC) = T(1, I)
        CC = CC + 1
    Next I
    Call Module1.LinearFit(T4, slope, B)
    For I = 0 To UBound(T, 2)
        T(1, I) = slope * T(0, I) + B
    Next I
Else
    ReDim T(1, UBound(T4, 2))
    Call Module1.LinearFit(T4, slope, B)
    For I = 0 To UBound(T4, 2)
        
        T(0, I) = T4(0, I)
        T(1, I) = slope * T4(0, I) + B
    Next I
    
End If
    GS.AddLine LNutrients.List(Display), gLine, QBColor(Display), T
    GS.AddLine LNutrients.List(Display), gScatter, QBColor(Display), T4
    
    Slopes(Display) = slope
    Bs(Display) = B
    If MultiLine Then
     
       Display = Display + 1
       Erase T
       Erase T4
       Erase T3
    End If
Loop Until Display > 7 Or MultiLine = False
    
    GS.DrawGraph
    GS.AutoScaleY = False
    
    
    If Display = 0 Then
        BMR = -1 * B / slope
        CPP = 1 / slope * 7
        If Display = 0 Then
            m_BMR = BMR
            m_CPP = CPP
        End If
    End If
End Sub


Public Sub SaveJournal()
  Dim temp As Recordset
  Set temp = DB.OpenRecordset("Select * from dailylog where ((User = """ & CurrentUser.Username & """) and (date = #" & TDate.Text & "#));", dbOpenDynaset)
  If temp.RecordCount = 0 Then
    temp.AddNew
  Else
    temp.Edit
  End If
  temp.Fields("User") = CurrentUser.Username
  temp.Fields("Date") = TDate.Text
  temp.Fields("Stress") = m_Stress
  temp.Fields("Sore") = m_Sore
  temp.Fields("Hunger") = m_Hunger
  temp.Fields("Energy") = m_Energy
  temp.Fields("Weight") = Val(TWeight.Text)
  RecordWeight = Val(TWeight.Text)
  If BFP <> 0 Then temp.Fields("Bfp") = BFP
  temp.Fields("Comments") = Journal.Text
  temp.Fields("BMR") = Val(LBMR.Caption)
  temp.Update
  temp.Close
  Set temp = Nothing
  
  If Val(TWeight.Text) <> 0 Then
      Set temp = DB.OpenRecordset("Select weight from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
      temp.Edit
      temp.Fields("Weight") = Val(TWeight.Text)
      temp.Update
      temp.Close
      Set temp = Nothing
  End If
End Sub


Private Sub Bonus_Click()
Call DoTotals
End Sub
Private Sub DoTotals()
   Dim I, O, S, B
   I = Val(Consumed.Caption)
   O = Val(Burned.Caption)
   B = Val(LBMR.Caption)
   S = I - O - B
   Total.Caption = Round(S)
   
End Sub
Private Sub Burned_Change()
Call DoTotals
End Sub

Private Sub Command1_Click()
Call SaveJournal
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()
On Error Resume Next
Me.Width = GS.Left + GS.Width + 100
Command4_Click
Me.Left = 0
End Sub

Private Sub Command4_Click()
Report.Visible = False
FToFrom.Visible = True
FNutrient.Visible = False
UserChange = False
  LFrom.Selected(0) = True
  LTo.Selected(0) = True
UserChange = True

GS.YAxisName = "Weight"
GS.XAxisName = ""
GS.CaptionName = "Weight over Time"
Dim T() As Single
Dim I As Long, CC As Long, LWeight As Single
Dim D As Date, M As Long, W As Long, Ws As Long, y As Long
Dim O As Date, O2 As Date, NM As Date
Dim dYear As Long, Y2 As Long
CC = 0
y = Year(Fields(0, 0))
O = "1/1/" & y
For I = 0 To UBound(Fields, 2)
  If Fields(8, I) = 1 Then
     ReDim Preserve T(1, CC)
     M = Month(Fields(0, I))
     Y2 = Year(Fields(0, I))
     O2 = M & "/1/" & Y2
     dYear = (Y2 - y) * 12
     NM = DateAdd("m", 1, O2)
     Ws = Int(DateDiff("d", O2, NM))
     W = Int(DateDiff("d", O2, Fields(0, I)))
     T(0, CC) = M + W / Ws + dYear
     
     If Fields(5, I) <> 0 Then
        T(1, CC) = Fields(5, I)
        LWeight = Fields(5, I)
     Else
        T(1, CC) = LWeight
     End If
     CC = CC + 1
  End If
Next I
CC = CC - 1
GS.AutoScaleY = True
GS.Reset
GS.AddLine "", gLine, vbRed, T
GS.DrawGraph

End Sub

Private Sub Command5_Click()
Report.Visible = False
FToFrom.Visible = True
FNutrient.Visible = False
UserChange = False
  LFrom.Selected(0) = True
  LTo.Selected(0) = True
UserChange = True

GS.YAxisName = "Pounds per week"
GS.XAxisName = ""
GS.CaptionName = "Predicted Weight Loss"

Dim T() As Single, T2() As Single, L
Dim I As Long, CC As Long
Dim D As Date, M As Long, W As Long, Ws As Long, y As Long
Dim O As Date, O2 As Date
Dim NM As Date, Y2 As Long, dYear As Long

CC = 0
y = Year(Fields(0, 0))
O = "1/1/" & y
For I = 0 To UBound(Fields, 2)
  If Fields(8, I) = 1 Then
     ReDim Preserve T(1, CC)
     ReDim Preserve T2(1, CC)
     M = Month(Fields(0, I))
     Y2 = Year(Fields(0, I))
     O2 = M & "/1/" & Y2
     dYear = (Y2 - y) * 12
     NM = DateAdd("m", 1, O2)
     Ws = Int(DateDiff("d", O2, NM))
     W = Int(DateDiff("d", O2, Fields(0, I)))
     T(0, CC) = M + W / Ws + dYear
     T2(0, CC) = T(0, CC)
     
     T(1, CC) = Fields(3, I)
     T2(1, CC) = Fields(4, I)
     CC = CC + 1
  End If
Next I

GS.AutoScaleY = True
GS.Reset
GS.AddLine "", gLine, vbRed, T
GS.AddLine "", gLine, vbBlue, T2
GS.DrawGraph
GS.AutoScaleY = False

End Sub

Private Sub Command6_Click()
Frame3.Visible = True
Frame4.Visible = False
Command6.Visible = False
End Sub

Private Sub Command7_Click()
  If Option1.Value Then
      If LCase$(sex) = "male" Then
         BFP = Round(86.01 * Log10(TWaistM - TNeckM) - 70.041 * Log10(CurrentUser.Height) + 36.76, 1)
      Else
         BFP = Round(-71.938 + 105.42 * Log10(CurrentUser.Weight) + 0.4396 * THips - 0.5086 * TWrist - 3.997 * TForearm - 1.3085 * CurrentUser.Height - 1.354 * TNeckW, 1)
      End If
  ElseIf Option2.Value Then
      BFP = Val(TDBFP.Text)
  End If
  LBfp.Caption = Round(BFP, 1) & "%"
End Sub


Private Sub Command8_Click()
Report.Visible = False
FToFrom.Visible = False
FNutrient.Visible = True

UserChange = False
  LFrom.Selected(0) = True
  LTo.Selected(0) = True
  LNutrients.Selected(0) = True
UserChange = True

GS.YAxisName = "Pounds lost"
GS.XAxisName = "Calories Consumed Daily"
GS.CaptionName = "Pounds lost With Calories Eaten"


Call ShowCaloriesToWeightLoss(0)
End Sub


Private Sub Consumed_Change()
Call DoTotals
End Sub

Private Sub CRBPF_Click()
Frame3.Visible = False
Frame4.Visible = True
Command6.Visible = True
End Sub

Private Sub CReport_Click()
Call ShowCaloriesToWeightLoss(0)
Call GS.SaveBitmap(App.Path & "\resources\temp\temp_CalMinusEx.bmp", 4, 3)
  
Call ShowCaloriesToWeightLoss(8)
Call GS.SaveBitmap(App.Path & "\resources\temp\temp_ALL.bmp", 4, 3)
    
Call ShowCaloriesToWeightLoss(6)
Call GS.SaveBitmap(App.Path & "\resources\temp\temp_Fiber.bmp", 3, 3)
  
Call ShowCaloriesToWeightLoss(4)
Call GS.SaveBitmap(App.Path & "\resources\temp\temp_Sugar.bmp", 3, 3)

Dim ff As Long

ff = FreeFile
Open App.Path & "\resources\temp\report.html" For Output As #ff
   Print #ff, "<html><body>"
  
   Print #ff, "Welcome to your body.  We have run reports using the data that <br>"
   Print #ff, "you have entered.  These will tell you what you need to do to loose weight and<br>"
   Print #ff, "to keep it off.  <br>"
   Print #ff, "<h2>Calories and Exercise</h2>"
   Print #ff, "<img src='temp_calMinusex.bmp'><br>"
   Print #ff, "We have places an imaginary line through your data to help us figure out<br>"
   Print #ff, "Your Base Metabolic Rate(BMR) and your Calories Per Pound(CPP)<br>"
   Print #ff, "<ul><li> Your BMR is " & Round(m_BMR) & "</li>"
   Print #ff, "<li>Your CPP is " & Round(m_CPP) & "</li></ul>"
   Print #ff, "<h2>Effects of Nutrients</h2>"
   Print #ff, "<img src='temp_ALL.bmp' ><br> "
   Print #ff, "From the information on the graph, we recomend that you follow<br>"
      
   Dim FP As Single, CP As Single, PP As Single
   If Abs(Slopes(3)) > Abs(Slopes(5)) And Abs(Slopes(3)) > Abs(Slopes(7)) Then
      FP = (-1.5 - Bs(3)) / Slopes(3)
      If Abs(Slopes(5)) > Abs(Slopes(7)) Then
        PP = (-0.5 - Bs(5)) / Slopes(5)
        CP = 100 - (FP + PP)
      Else
        CP = (-0.5 - Bs(7)) / Slopes(7)
        PP = 100 - (CP + FP)
      End If
   ElseIf Abs(Slopes(5)) > Abs(Slopes(3)) And Abs(Slopes(5)) > Abs(Slopes(7)) Then
      PP = (-1.5 - Bs(5)) / Slopes(5)
      If Abs(Slopes(3)) > Abs(Slopes(7)) Then
        FP = (-0.5 - Bs(3)) / Slopes(3)
        CP = 100 - (FP + PP)
      Else
        CP = (-0.5 - Bs(7)) / Slopes(7)
        FP = 100 - (CP + PP)
      End If
   ElseIf Abs(Slopes(7)) > Abs(Slopes(5)) And Abs(Slopes(7)) > Abs(Slopes(3)) Then
      CP = (-1.5 - Bs(7)) / Slopes(7)
      If Abs(Slopes(5)) > Abs(Slopes(3)) Then
        PP = (-0.5 - Bs(5)) / Slopes(5)
        FP = 100 - (CP + PP)
      Else
        FP = (-0.5 - Bs(3)) / Slopes(3)
        PP = 100 - (CP + FP)
      End If
   End If
   Print #ff, "<table>"
   Print #ff, "<tr><td>% Protein</td><td>" & Round(PP) & "</td></tr>"
   Print #ff, "<tr><td>% Carbs</td><td>" & Round(CP) & "</td></tr>"
   Print #ff, "<tr><td>% Fat</td><td>" & Round(FP) & "</td></tr>"
   Print #ff, "</table>"
   Print #ff, "<h2>Types of carbs</h2>"
   Print #ff, "<table><tr><td>"
   Print #ff, "<img src='temp_Fiber.bmp' ></td><td> "
   Print #ff, "<img src='temp_sugar.bmp' ></td></tr></table> "

   Print #ff, "<FORM>"
   Print #ff, "<input type=""button"" value=""Print"" Onclick=""print()"";>"
   Print #ff, "</FORM>"

   Print #ff, "</body>"
   Print #ff, "</html>"
 

Close #ff
Report.Visible = True
Report.Navigate2 App.Path & "\resources\temp\report.html"
End Sub

Private Sub Form_Load()

'Set GS.Font = Me.Font
GS.OverSizeY = True

Me.Width = Command2.Left + Command2.Width + 100
  TDate.Text = DisplayDate

Dim I As Long
For I = 1 To 10
   LEnergy.AddItem I
   LSore.AddItem I
   LHunger.AddItem I
   LStress.AddItem I
Next I
  ResetMe DisplayDate
  
  
  LNutrients.AddItem "Calories - Exercise"
  LNutrients.AddItem "Calories"
  LNutrients.AddItem "Exercise Calories"
  LNutrients.AddItem "Fat"
  LNutrients.AddItem "Sugar"
  LNutrients.AddItem "Protein"
  LNutrients.AddItem "Fiber"
  LNutrients.AddItem "Carbs"
  LNutrients.AddItem "All Nutrients"

  
  
Dim Months(12) As String
Months(1) = "Jan"
Months(2) = "Feb"
Months(3) = "Mar"
Months(4) = "Apr"
Months(5) = "May"
Months(6) = "Jun"
Months(7) = "Jul"
Months(8) = "Aug"
Months(9) = "Sep"
Months(10) = "Oct"
Months(11) = "Nov"
Months(12) = "Dec"
UserChange = False
LTo.AddItem "AutoSelect"
LFrom.AddItem "AutoSelect"
For I = 1 To 12
  LTo.AddItem Months(I), I
  LFrom.AddItem Months(I), I
Next I
LTo.Selected(0) = True
LFrom.Selected(0) = True
UserChange = True
End Sub

Private Sub GoalCalories_Click()
Call DoTotals
End Sub

Private Sub LEnergy_Click()

  m_Energy = SelectedList(LEnergy)
End Sub

Private Sub LFrom_Click()
If UserChange Then
   Dim I
   For I = 0 To LFrom.ListCount - 1
     If LFrom.Selected(I) Then
       GoTo DrawIt
     End If
   Next I
   Exit Sub
DrawIt:
   If I = 0 Then
     GS.AutoScaleXMin = True
     GS.DrawGraph
     Exit Sub
   End If

 
   GS.AutoScaleXMin = False
   Call GS.SetXMax(0, I)
   GS.DrawGraph
   
End If
End Sub

Private Sub LHunger_Click()
m_Hunger = SelectedList(LHunger)
End Sub

Private Sub LNutrients_Click()
If UserChange Then
    Dim I As Long, j As Integer
    j = -1
    For I = 0 To LNutrients.ListCount - 1
        If LNutrients.Selected(I) Then
            j = I
            Exit For
        End If
    Next I
    If j = -1 Then Exit Sub
    Call ShowCaloriesToWeightLoss(j)
End If
End Sub

Private Sub LSore_Click()
m_Sore = SelectedList(LSore)
End Sub

Private Sub LStress_Click()
m_Stress = SelectedList(LStress)
End Sub




Private Sub LTo_Click()
If UserChange Then
   Dim I
   For I = 0 To LTo.ListCount - 1
     If LTo.Selected(I) Then
       GoTo DrawIt
     End If
   Next I
   Exit Sub
DrawIt:
   If I = 0 Then
     GS.AutoScaleXMax = True
     GS.DrawGraph
     Exit Sub
   End If
  
   GS.AutoScaleXMax = False
   Call GS.SetXMax(1, I)
   GS.DrawGraph
End If
End Sub

Private Sub Option1_Click()
Command7.Visible = True
FClothTape.Visible = True
FOTher.Visible = False
End Sub

Private Sub Option2_Click()
Command7.Visible = False
FClothTape.Visible = False
FOTher.Visible = True
End Sub

Private Sub TDBFP_Change()
Command7_Click
End Sub
