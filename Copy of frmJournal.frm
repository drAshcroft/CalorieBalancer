VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJournal 
   Caption         =   "Daily Journal"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11550
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Daily Journal"
   Begin VB.CommandButton CCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Journal"
            Key             =   "Journal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Progress Report"
            Key             =   "Progress"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Report 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      ExtentX         =   11456
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

Dim CaloriesVsX() As Single

Dim today As Date


Public Sub ResetMe(mToday As Date)
   today = mToday


   Dim temp As Recordset, Junk As String, junk2 As String, ProfileBMR As Single
   Dim junkL As Long
   Dim Yesterday As Date
   Set temp = DB.OpenRecordset("Select * from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
   If LCase$(temp.Fields("Sex")) = "male" Then sex = True Else sex = False

   ProfileBMR = Val(temp.Fields("BMR"))
   
   Yesterday = DateAdd("d", -1, today)
   Set temp = DB.OpenRecordset("Select * from dailylog where ((User = '" & CurrentUser.Username & "') and (date = #" & today & "#));", dbOpenDynaset)
   If temp.RecordCount <> 0 Then
     Junk = temp.Fields("Exercise_cal")
     M_Burned = Val(Junk)
   Else
     M_Burned = 0
   End If
   
   On Error Resume Next
   temp.Close
   Set temp = Nothing
   Set temp = DB.OpenRecordset("Select * from dailylog where ((User = """ & CurrentUser.Username & """) and (date = #" & TDate.Text & "#));", dbOpenDynaset)
   If temp.RecordCount <> 0 Then
     Junk = temp.Fields("calories")
     M_eaten = Round(Val(Junk), 0)
     
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
   
   On Error GoTo 0
Dim MinWeek As Long, MaxWeek As Long, I As Long
Set temp = DB.OpenRecordset("Select * from WeekLog Where User = '" & CurrentUser.Username & "' order by week asc;", dbOpenDynaset)
Erase Fields

If Not temp Is Nothing And Not temp.EOF Then
  Dim Mindate As Date, D As Date
  Dim firstSunday As Date
  Mindate = temp.Fields("week")
  I = Weekday(Mindate, vbSunday)
  firstSunday = DateAdd("d", -1 * I, Mindate)



While Not temp.EOF
   D = temp.Fields("week")
   I = DateDiff("d", Mindate, D)
   ReDim Preserve Fields(12, I)
   Fields(0, I) = temp("Week")
   Fields(1, I) = temp.Fields("Weight")
   Fields(2, I) = temp.Fields("Weightloss")
   Fields(3, I) = temp.Fields("Calories")
   Fields(4, I) = temp("bmr")
   Fields(5, I) = temp.Fields("carbs")
   Fields(6, I) = temp.Fields("exercise")
   Fields(7, I) = temp.Fields("fat")
   Fields(8, I) = temp.Fields("fiber")
   Fields(9, I) = temp.Fields("protein")
   Fields(10, I) = temp.Fields("sugar")
   Fields(11, I) = 1
   Fields(12, I) = temp.Fields("TWeightloss")
   temp.MoveNext
Wend
End If
temp.Close
Set temp = Nothing

If today = Date Then
   Dim BMR As Single, CPP As Single
 '  Call ShowCaloriesToWeightLoss(0, True, BMR, CPP)

   Set temp = DB.OpenRecordset("Select * from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
   If BMR = 0 Then
 '     Call ShowCaloriesToWeightLoss(0, False, BMR, CPP)
   End If
   If BMR < 1000 Then
      BMR = DoCalories(temp("sex"), temp("age"), temp("Height"), temp("weight"), temp("Bfp"))
   End If
   temp.Edit
 '  temp("bmr") = BMR
 '  temp("calPound") = CPP
   temp.Update
   temp.Close

   Set temp = Nothing
   ProfileBMR = BMR
End If

   
   
   LBMR.Caption = Round(ProfileBMR)


'get all the weekly information
Dim L
Dim CC As Long, j As Long
 CC = 0
 On Error GoTo JumpOut
 For I = 0 To UBound(Fields, 2)
     If Fields(11, I) = 1 Then
         ReDim Preserve CaloriesVsX(10, CC)
         CaloriesVsX(0, CC) = Fields(1, I)
         For j = 1 To 9
           CaloriesVsX(j, CC) = Fields(j + 1, I) '/ 7 ' = temp.Fields("Weight")
         Next j
         CaloriesVsX(2, CC) = CaloriesVsX(2, CC) / 7
         CaloriesVsX(6, CC) = CaloriesVsX(6, CC) / 7
         CaloriesVsX(8, CC) = CaloriesVsX(8, CC) / 7
         CaloriesVsX(9, CC) = CaloriesVsX(9, CC) / 7
         CaloriesVsX(4, CC) = CaloriesVsX(4, CC) / 7
         
         CC = CC + 1
     End If
 Next I
JumpOut:
 'now get all the daily information
 Dim TT, T2
 TT = Module1.DoDays

For I = 0 To UBound(TT, 2)
         ReDim Preserve CaloriesVsX(10, CC)
         For j = 0 To 9
           CaloriesVsX(j, CC) = TT(j, I)   ' = temp.Fields("Weight")
         Next j
         CC = CC + 1
Next I


For I = 0 To UBound(CaloriesVsX, 2)
  CaloriesVsX(6, I) = CaloriesVsX(6, I) * 9 / CaloriesVsX(2, I) * 100
  CaloriesVsX(8, I) = CaloriesVsX(8, I) * 4 / CaloriesVsX(2, I) * 100
  CaloriesVsX(9, I) = CaloriesVsX(9, I) * 4 / CaloriesVsX(2, I) * 100
  CaloriesVsX(4, I) = CaloriesVsX(4, I) * 4 / CaloriesVsX(2, I) * 100
  CaloriesVsX(10, I) = CaloriesVsX(2, I) - CaloriesVsX(5, I)
Next I



End Sub


Private Sub ShowCaloriesToWeightLoss(Display As Integer)

Dim I As Long
Dim DisplayData() As Single
Dim DisplayLine As Long

  GS.Reset
  GS.CaptionFont.Size = 13
  GS.ShowLegend = False
  GS.AutoScaleY = False
  Call GS.SetYMax(0, -1)
  Call GS.SetYMax(1, 1)
  If Display = 0 Then
    GS.CaptionName = "Weight Loss with (Calories - Exercise)"
    GS.XAxisName = "(Calories - Exercise) Daily"
    DisplayLine = 10
  ElseIf Display = 1 Then
    GS.CaptionName = "Weight Loss with Calories Consumed"
    GS.XAxisName = "Calories Daily"
    DisplayLine = 2
  ElseIf Display = 2 Then
    GS.CaptionName = "Weight Loss with Exercise"
    GS.XAxisName = "Exercise Calories Burned Daily"
    DisplayLine = 5
  ElseIf Display = 3 Then
    GS.CaptionName = "Effect of Fat on Diet"
    GS.XAxisName = "Percent of Fat Calories Daily"
    DisplayLine = 6
  ElseIf Display = 4 Then
    GS.CaptionName = "Effect of Sugar on Diet"
    GS.XAxisName = "Percent of Sugar Calories Daily"
    DisplayLine = 9
  ElseIf Display = 5 Then
    GS.CaptionName = "Effect of Protein on Diet"
    GS.XAxisName = "Percent of Protein Calories Daily"
    DisplayLine = 8
  ElseIf Display = 6 Then
    GS.CaptionName = "Effect of Fiber on Diet"
    GS.XAxisName = "Grams of Fiber Daily"
    DisplayLine = 7
  ElseIf Display = 7 Then
    GS.CaptionName = "Effect of Carbs on Diet"
    GS.XAxisName = "Percent of Carbs Calories Daily"
    DisplayLine = 4
  ElseIf Display = 8 Then
    GS.CaptionName = "Effect of All Nutrients on Diet"
    GS.XAxisName = "Percent Daily"
    GS.ShowLegend = True
  End If



    ReDim DisplayData(1, UBound(CaloriesVsX, 2))
    For I = 0 To UBound(CaloriesVsX, 2)
       DisplayData(0, I) = CaloriesVsX(DisplayLine, I)
       DisplayData(1, I) = CaloriesVsX(1, I)
    Next I



Dim slope As Double, B As Double
Dim T() As Single
Call Module1.LinearFit(DisplayData, slope, B)

ReDim T(1, UBound(DisplayData, 2))
For I = 0 To UBound(DisplayData, 2)
   T(1, I) = slope * DisplayData(0, I) + B
   T(0, I) = DisplayData(0, I)
Next I
    
GS.MonthAxis = False
GS.AddLine LNutrients.List(Display), gLine, vbRed, T
GS.AddLine LNutrients.List(Display), gScatter, 0, DisplayData
    
GS.DrawGraph
    
    
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
  If BFP <> 0 Then
     temp.Fields("Bodyfat") = BFP
  End If
  temp.Fields("Comments") = Journal.Text
  temp.Fields("BMR") = Val(LBMR.Caption)
  temp.Update
  temp.Close
  Set temp = Nothing
  
  If Val(TWeight.Text) <> 0 Then
      Set temp = DB.OpenRecordset("Select * from profiles where (user = '" & CurrentUser.Username & "');", dbOpenDynaset)
      temp.Edit
      temp.Fields("Weight") = Val(TWeight.Text)
      If TDate.Text = Date Then
        If BFP <> 0 Then temp("BFP") = BFP
        temp("Weight") = RecordWeight
        temp("age") = Abs(DateDiff("d", Date, temp("birthdate")) / 365)
      End If
      temp.Update
      temp.Close
      Set temp = Nothing
  End If
End Sub

Private Sub WeightGraph()
'get the right stuff visible
Report.Visible = False
FToFrom.Visible = True
FNutrient.Visible = False


UserChange = False
  LFrom.Selected(0) = True
  LTo.Selected(0) = True
UserChange = True

'set up the graph
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
  If Fields(11, I) = 1 Then
     ReDim Preserve T(1, CC)
     M = Month(Fields(0, I))
     Y2 = Year(Fields(0, I))
     O2 = M & "/1/" & Y2
     dYear = (Y2 - y) * 12
     NM = DateAdd("m", 1, O2)
     Ws = Int(DateDiff("d", O2, NM))
     W = Int(DateDiff("d", O2, Fields(0, I)))
     T(0, CC) = M + W / Ws + dYear
     
     If Fields(1, I) <> 0 Then
        T(1, CC) = Fields(1, I)
        LWeight = Fields(1, I)
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

Private Sub WeightLossGraph()
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
  If Fields(11, I) = 1 Then
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
     
     T(1, CC) = Fields(2, I)
     T2(1, CC) = Fields(12, I)
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


Private Sub MakeReport()
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

Private Sub CSave_Click()
  Call SaveJournal
  Me.Hide
End Sub

Private Sub Form_Load()
Report.Navigate2 App.Path & "\Resources\Journal\boxed.htm"

'Set GS.Font = Me.Font
GS.OverSizeY = True

TDate.Text = DisplayDate
End Sub

Private Sub Form_Resize()
Report.Move 0, TabStrip1.Top + TabStrip1.Height, Me.ScaleWidth, Me.ScaleHeight - (TabStrip1.Top + TabStrip1.Height)
End Sub


Private Sub Report_DownloadComplete()

End Sub


