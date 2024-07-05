VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMenuPlanner 
   BackColor       =   &H00C00000&
   Caption         =   "Menu Planner"
   ClientHeight    =   9975
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   13950
   Icon            =   "frmMenuPlanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   675
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser PP 
      Height          =   30
      Left            =   120
      TabIndex        =   10
      Top             =   9720
      Visible         =   0   'False
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
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
   Begin RichTextLib.RichTextBox MealView 
      Height          =   1575
      Left            =   2160
      TabIndex        =   8
      Top             =   8040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMenuPlanner.frx":57E2
   End
   Begin CalorieTracker.MonthDayPicker Calendar 
      Height          =   1935
      Left            =   10920
      TabIndex        =   5
      Top             =   8040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3413
   End
   Begin VB.PictureBox WeekHold 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   7935
      Left            =   2160
      ScaleHeight     =   7875
      ScaleWidth      =   11715
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      Begin CalorieTracker.PieChart Balance 
         Height          =   2130
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   5160
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   3757
         BackColor       =   16777215
         MaskPicture     =   "frmMenuPlanner.frx":5864
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Blend           =   5
      End
      Begin VB.Label LCalories 
         BackStyle       =   0  'Transparent
         Caption         =   "Calories = 0"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Shape HiLite 
         Height          =   1695
         Left            =   7185
         Top             =   1365
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label MealItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   735
         Index           =   0
         Left            =   3480
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Line GridLines 
         BorderColor     =   &H00C0C000&
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   240
         X2              =   5520
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line TitleLine 
         X1              =   120
         X2              =   5880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line DaySep 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   2400
         X2              =   2400
         Y1              =   360
         Y2              =   6360
      End
      Begin VB.Label TDays 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
   End
   Begin MSComctlLib.TreeView TPlans 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1296
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1296
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label DropLabel 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveMonth 
         Caption         =   "Save Meal Plan to File"
      End
      Begin VB.Menu mnuOpenPlan 
         Caption         =   "Open Meal Plan from File"
      End
      Begin VB.Menu MnuUpload 
         Caption         =   "Upload Meal Plan to Server"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep1243 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintInstructions 
         Caption         =   "Print Meal Instructions"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Meal Plan"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuMealTools 
      Caption         =   "Meal Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuNewMeal 
         Caption         =   "New Meal"
      End
      Begin VB.Menu mnuViewMeal 
         Caption         =   "View Meal"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeleteMeal 
         Caption         =   "Delete Meal"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuEditMeal 
         Caption         =   "Edit Meal"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAddtoFavorites 
         Caption         =   "Add to favorites"
      End
   End
   Begin VB.Menu mnuNewMealOnly 
      Caption         =   "New Meal"
      Begin VB.Menu mnuAddNewMeal 
         Caption         =   "Add New Meal"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuplans 
      Caption         =   "mnuPlans"
      Visible         =   0   'False
      Begin VB.Menu mnuNewPlan 
         Caption         =   "New Plan"
      End
      Begin VB.Menu mnuRunPlan 
         Caption         =   "Run Plan"
      End
      Begin VB.Menu mnuDeletePlan 
         Caption         =   "Delete Plan"
      End
   End
   Begin VB.Menu mnuMealOptions 
      Caption         =   "mnuMealOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteWeekMeal 
         Caption         =   "Delete Meal"
      End
   End
End
Attribute VB_Name = "frmMenuPlanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'plan information
Dim MealNodes As MSComctlLib.Nodes
Dim PlanNodes As MSComctlLib.Nodes
Dim MealDesc As Collection
Dim AbbrevIDS(5) As Long

Public CurrentPlanID As Long
Private Type Cals
   ID As Long
   'nDay As Long
   'nMeal As Long
   Calories As Single
   carbs As Single
   fiber As Single
   fat As Single
   Protein As Single
   sugar As Single
   Changed As Boolean
   MealItemID As Long
End Type

Dim PrintMode As String

Dim MealCals(7, 6) As Cals

'currentdate
Dim m_Month As Long
Dim m_Year As Long
Dim m_Day As Long

'used for saving
Dim Changed(7) As Boolean

'used for the drag information
Dim DragMealName As String
Dim DragID As Long, DragMeal As String, DragCals As Single

'geometery: width and height of meal boxes
Dim sW As Single, sH As Single

'for the seperator
Dim MouseDWN As Boolean, UpDown As Boolean

'for dragdrop
Dim StartX As Single, StartY As Single, MouseMode As Boolean
Dim MoveMode As Boolean, MoveIndex(1) As Long
Dim Hilited As Long

'weekdates
Dim firstSunday As Date, Lastday As Date


Dim NM_F As Single, NM_c As Single, NM_FB As Single, NM_P As Single, NM_S As Single, NM_Calories As Single

Private Sub SaveWeek()
On Error GoTo errhandl
Dim temp As Recordset
Dim i As Long
Dim Lastday As Date
Dim nDay As Long
Lastday = DateAdd("d", 6, firstSunday)
'Debug.Print Lastday
Set temp = DB.OpenRecordset("Select * from meals where entrydate >=#" & firstSunday & "# " _
              & "and entrydate<=#" & Lastday & "# and user = '" & CurrentUser.Username & "';", dbOpenDynaset)
While Not temp.EOF
   temp.Delete
   temp.MoveNext
Wend
Dim j As Long
For i = 0 To 7 'UBound(MealCals)
  For j = 0 To 6
  With MealCals(i, j)
   If .MealItemID <> -1 Then
       temp.AddNew
       temp("MealId") = .ID
       temp("MealNumber") = j
       temp("User") = CurrentUser.Username
       temp("EntryDate") = DateAdd("d", i, firstSunday)
       temp.Update
   End If
  End With
  Next
Next
For i = 0 To 6
  For j = 0 To 5
    If MealCals(i, j).Changed Then
       Call SaveDaytoList(j, DateAdd("d", i, firstSunday))
       MealCals(i, j).Changed = False
    End If
  Next j
  Changed(i) = False
Next i
temp.Close
Set temp = Nothing

Call frmMain.FlexDiet.OpenDay(DisplayDate)
Exit Sub
errhandl:
MsgBox "Unable to save week.  Please check all entries." & vbCrLf & Err.Description, vbOKOnly, ""
'Resume
End Sub

Public Sub SaveDaytoList(nMeal As Long, Curdate As Date, Optional Checkdays As Boolean = True)
On Error GoTo errhandl
  Dim temp As Recordset, MealPlans As Recordset
  
  
  Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.Date)=#" & Curdate & "#) AND " _
  & "(DaysInfo.user='" & CurrentUser.Username & "') and " _
  & " meal=" & nMeal & ") " _
  & " ORDER BY daysinfo.order;", dbOpenDynaset)
  Set MealPlans = DB.OpenRecordset("SELECT * From Meals " _
   & "WHERE (((Meals.User)='" & CurrentUser.Username & "') AND " _
   & "((Meals.EntryDate)=#" & Curdate & "#) AND" _
   & "((Meals.mealnumber)=" & nMeal & "));", dbOpenDynaset)
                 
  If Not (temp.EOF And temp.BOF) Then
'     temp.MoveFirst
     While Not temp.EOF
       temp.Delete
       temp.MoveNext
     Wend
     temp.Close
  End If
     
  If Not (MealPlans.EOF And MealPlans.BOF) Then
     MealPlans.Close
     Call UpLoadMealPlan(nMeal, Curdate, CurrentUser.Username)
  End If

  Set temp = Nothing
  Set MealPlans = Nothing
  Exit Sub
errhandl:
 
  Resume Next
End Sub

Public Sub UpLoadMealPlan(nMeal As Long, D As Date, USER As String)
Dim i As Long
Dim temp As Recordset, temp2 As Recordset
Dim Daily As Recordset
Dim junk As String, junk2 As Long
Dim dy As Long, k As Long
   On Error GoTo errhandl
AbbrevIDS(0) = -200
AbbrevIDS(1) = -201
AbbrevIDS(2) = -202
AbbrevIDS(3) = -201
AbbrevIDS(4) = -203
AbbrevIDS(5) = -204
        
        'open the daily log
        Set Daily = DB.OpenRecordset("Select * from daysinfo where " _
        & "meal = " & nMeal & " and date = #" & D & "# and user = '" & USER & "';", dbOpenDynaset)
        'now get the meal plan
        Set temp = DB.OpenRecordset("SELECT * From Meals " _
           & "WHERE (((Meals.User)='" & USER & "') AND " _
           & "((Meals.EntryDate)=#" & D & "#) AND " _
           & "mealnumber = " & nMeal & " ) " _
           & "order by mealnumber;", dbOpenDynaset)
        k = 0
        While Not temp.EOF
           junk2 = temp.Fields("MealID")
           Set temp2 = DB.OpenRecordset("Select * from MealDefinition where (MealID = " & junk2 & ");", dbOpenDynaset)
           If Not temp2.EOF Then
                Daily.AddNew
                Daily.Fields("user") = USER
                Daily.Fields("date") = D
                Daily.Fields("itemid") = AbbrevIDS(temp("mealnumber"))
               ' Debug.Print D, temp("mealnumber"), AbbrevIDS(temp("mealnumber"))
                Daily.Fields("meal") = nMeal
                Daily.Fields("order") = 0
                Daily.Update
                k = k + 1
                While Not temp2.EOF
                    Daily.AddNew
                    Daily.Fields("user") = USER
                    Daily.Fields("Date") = D
                    Daily.Fields("ItemID") = temp2.Fields("AbbrevID")
                    Daily.Fields("Servings") = temp2.Fields("Serving")
                    Daily.Fields("Unit") = temp2.Fields("Unit")
                    Daily.Fields("meal") = nMeal
                    Daily.Fields("order") = 0
                    k = k + 1
                    Daily.Update
                    temp2.MoveNext
                Wend
                temp2.Close
                Set temp2 = Nothing
           End If
           temp.MoveNext
        Wend
        
        On Error Resume Next
        temp.Close
        Set temp = Nothing

Exit Sub
errhandl:
 
  Resume Next
End Sub


Private Sub OpenWeek(WeekDate As Date)
On Error GoTo errhandl
Dim temp As Recordset
Dim i As Long, j As Long, k As Long


On Error Resume Next
For i = 1 To MealItem.Count
   Unload MealItem(i)
Next i
On Error GoTo errhandl
For i = 0 To 7
  For j = 0 To 6
     MealCals(i, j).MealItemID = -1
  Next j
Next i

firstSunday = FindFirstDay(WeekDate)
Lastday = DateAdd("d", 6, firstSunday)
Set temp = DB.OpenRecordset("Select * from meals where entrydate >=#" & firstSunday & "# " _
              & "and entrydate<=#" & Lastday & "# and user = '" & CurrentUser.Username & "';", dbOpenDynaset)
              
Dim T As Single
T = TDays(1).Height
While Not temp.EOF
  j = Abs(DateDiff("d", temp("entrydate"), firstSunday))
  k = temp("Mealnumber")
  Call NewItem(j * sW, k * sH + T, temp("MealId"), j, k)
  MealCals(j, k).Changed = False
  temp.MoveNext
Wend

temp.Close
Set temp = Nothing
Dim DD As Date

For i = 1 To 7
   DD = DateAdd("d", i - 1, firstSunday)
   TDays(i).Caption = Month(DD) & "/" & Day(DD) & " " & WeekdayName(i, True, vbSunday)
Next i
Call DoTotals
Exit Sub
errhandl:
'MsgBox Err.Description, vbOKOnly, ""
'Resume
End Sub
Public Sub LoadPlan(PlanID As Long)
  On Error Resume Next
  Dim temp As Recordset
  Dim junk As String, junk2 As String, junk3 As Long
  Dim MealCals As Single
  If MealNodes Is Nothing Then Exit Sub
  MealNodes.Clear
  Call SetHeads
  If PlanID = -2 Then
     Set temp = DB.OpenRecordset("select * from mealplanner where (mealid>-1);", dbOpenDynaset)
  Else
     Set temp = DB.OpenRecordset("Select * from mealplanner where (planid = " & PlanID & ");", dbOpenDynaset)
  End If
  Dim i As Long
  For i = 1 To MealDesc.Count
     MealDesc.Remove 1
  Next i
  
  While Not temp.EOF
    junk3 = temp.Fields("mealid")
    If Val(junk3) <> -1 Then
        junk = temp.Fields("Meal")
        junk2 = temp.Fields("mealname")
        MealCals = temp("calories")
        PlanID = temp.Fields("PlanID")
        Call MealNodes.Add(LCase$(junk), 4, "I" & junk3 & "~" & MealCals, junk2)
        junk = temp("description")
        MealDesc.Add junk, STR(junk3)
    End If
    temp.MoveNext
  Wend
  temp.Close
  Set temp = Nothing
  Exit Sub
errhandl:
  MsgBox "Unable to load all meals" & vbCrLf & Err.Description, vbOKOnly, ""
'Resume
End Sub

Public Sub Init()
  Dim temp As Recordset
  Dim junk As String, junk2 As String, junk3 As Long, PlanID As Long
  Dim n As Node
  On Error Resume Next
  
  PlanNodes.Clear
  PlanNodes.Add , , "i1", "Favorites"
  Set temp = DB.OpenRecordset("Select * from mealplanner where (mealid = -1) and (calories = 0);", dbOpenDynaset)
  
  While Not temp.EOF
    junk3 = temp.Fields("mealid")
    junk = temp.Fields("Mealname")
    junk2 = temp.Fields("PlanID")
    n = PlanNodes.Add(, , "I" & junk2, junk) ', 3)
    n.Expanded = True
    temp.MoveNext
  Wend

  temp.Close
  Set temp = Nothing
'  MonthAndYear.Caption = Months(m_Month) & " " & m_Year
  Call LoadPlan(1)
'  Call OpenMonth
  CurrentPlanID = 1
End Sub



Private Sub Calendar_DateSelected(NewDate As Date)
  On Error Resume Next
  Call SaveWeek
  Call OpenWeek(NewDate)
End Sub

Private Sub CEdit_Click()
On Error Resume Next
mnuEditMeal_Click
End Sub

Private Sub cView_Click()
On Error Resume Next
mnuViewMeal_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
   Call DeleteSomething
End If
End Sub

Private Sub Form_Load()

On Error Resume Next
Set MealDesc = New Collection
Dim today As Date
today = Date
m_Month = Month(today)
m_Day = Day(today)
m_Year = Year(today)



Dim i As Long, j As Long
For i = 0 To 7
   For j = 0 To 6
     MealCals(i, j).MealItemID = -1
   Next j
Next i



AbbrevIDS(0) = -200
AbbrevIDS(1) = -201
AbbrevIDS(2) = -202
AbbrevIDS(3) = -201
AbbrevIDS(4) = -203
AbbrevIDS(5) = -204


    ReDim Meals(7)
    Meals(0) = "B"
    Meals(1) = "S"
    Meals(2) = "L"
    Meals(3) = "S"
    Meals(4) = "D"
    Meals(5) = "T"
    
    Set MealNodes = TreeView1.Nodes
    Set PlanNodes = TPlans.Nodes
    PlanNodes.Add , , "i1", "Favorites"
    
   Call SetHeads
   
   
  For i = 2 To 7
      Load TDays(i)
      Load Balance(i)
      TDays(i).Visible = True
      Balance(i).Visible = True
     
      Load DaySep(i)
      Load LCalories(i)
      DaySep(i).Visible = True
      LCalories(i).Visible = True
  Next i
        Dim PT As Single
        'Dim F As Single, c As Single, FB As Single, P     As Single, Calories As Single, S As Single
        'Dim Calories As Single
        NM_Calories = Nutmaxes("Calories")
        NM_F = Nutmaxes("Fat")
        NM_c = Nutmaxes("Carbs")
        NM_FB = Nutmaxes("Fiber")
        NM_P = Nutmaxes("Protein")
        NM_S = Nutmaxes("Sugar")
  
  For i = 1 To 7
     TDays(i).Caption = WeekdayName(i, True, vbSunday)
     Call Module1.FigurePercentages(Balance(i), NM_Calories, NM_F, NM_S, NM_c, NM_P, NM_FB)
  Next i
   For i = 2 To 6
      Load GridLines(i)
      GridLines(i).Visible = True
   Next i
  ' Me.WindowState = 2
   Call Form_Resize
'   Call ClearMonth
   Call OpenWeek(DisplayDate)
'   Call FillMonth
   Call Init

   Unload Dialog
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If X - TreeView1.Width < 50 Then
 MouseDWN = True
 Y = Y - TPlans.Height
End If
If Y < 50 And Y > 0 And X < TreeView1.Width Then
   MouseDWN = False
   UpDown = True
End If

errhandl:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo errhandl
  If MouseDWN Then
     TreeView1.Width = X - 25
     Form_Resize
  End If
  If UpDown Then
     TPlans.Height = Y - 25
     TreeView1.Top = Y + 25
     TreeView1.Height = Me.ScaleHeight - (Y + 25)
  End If
  DoEvents
errhandl:
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo errhandl
  If MouseDWN Then
     MouseDWN = False
     TreeView1.Width = X - 25
     Form_Resize
  End If
  If UpDown Then
     TPlans.Height = Y - 25
     TreeView1.Top = Y + 25
     TreeView1.Height = Me.ScaleHeight - (Y + 25)
     UpDown = False
  End If
errhandl:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Not NoQuestions Then
Dim CHD As Boolean, i As Long
 For i = 1 To 7
    CHD = CHD Or Changed(i)
 Next i
If CHD Then
   Dim ret As VbMsgBoxResult
   ret = MsgBox("Do you wish to save your changes?", vbYesNoCancel, "")
   If ret = vbYes Then
     Call SaveWeek
   ElseIf ret = vbCancel Then
     Cancel = 1
   End If
End If
frmMain.FlexDiet.OpenDay DisplayDate
End If
End Sub



Private Sub MealItem_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
 On Error Resume Next
 Call weekhold_DragDrop(Source, X + MealItem(Index).Left, Y + MealItem(Index).Top)
End Sub

Private Sub MealItem_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
HiLite.Move MealItem(Index).Left, MealItem(Index).Top, sW, sH
End Sub

Private Sub MealItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
'If Button = 1 Then
Dim j As Long, k As Long
Dim T As Single
X = MealItem(Index).Left + MealItem(Index).Width / 2
Y = MealItem(Index).Top + MealItem(Index).Height / 2

T = TDays(1).Height
j = Int(X / sW)
k = Int((Y - T) / sH)


   StartX = X
   StartY = Y
   DropLabel.Caption = MealItem(Index).Caption
   MouseMode = True
'      DragMeal = Left$(Junk, InStr(1, Junk, "\", vbBinaryCompare) - 1)
   DragMealName = MealItem(Index).Caption
'      Junk = Node1.Key
'      parts = Split(Junk, "~")
'      DragID = Val(Right$(parts(0), Len(parts(0)) - 1))
'      DragCals = Val(parts(1))
'   End If
   
   MoveIndex(0) = j
   MoveIndex(1) = k
   Hilited = Index
'Else
If Button = 2 Then
   MouseMode = False
End If
errhandl:
End Sub

Private Sub MealItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If MouseMode Then
    If Abs(StartX - X) > 100 Or Abs(StartY - Y) > 100 And Index = Hilited Then
       MoveMode = True
       'MealItem(Index).ForeColor = RGB(200, 200, 200)
       HiLite.Move -sW, -sH
       HiLite.Visible = True
       DropLabel.Move X + WeekHold.Left + MealItem(Index).Left, Y + WeekHold.Top + MealItem(Index).Top
       'DropLabel.Move X, Y
       DropLabel.Drag vbBeginDrag
       
     
       Changed(Int(X / sW)) = True
       
       MouseMode = False
    End If
End If
errhandl:
End Sub

Private Sub MealItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If Button = 1 Then
    MouseMode = False
    MoveMode = False
    Dim T As Single
    HiLite.Move MealItem(Index).Left, MealItem(Index).Top, sW, sH
    HiLite.Visible = True
    Hilited = Index
Else
    PopUpMenu mnuMealOptions, , WeekHold.Left + MealItem(Index).Left + X, MealItem(Index).Top + Y

End If
errhandl:
End Sub

Private Sub mnuAddNewMeal_Click()
On Error GoTo errhandl
  Call SaveWeek
  Call mnuNewMeal_Click
  Exit Sub
errhandl:
MsgBox "Unable to add meal." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub mnuAddtoFavorites_Click()
'On Error Resume Next
  Dim temp As Recordset
  Dim junk As String, Parts() As String
  Dim parts2(), i As Long
  'Dim tNode As MSComctlLib.Node
  'Set tNode = TPlans.HitTest(X, Y)
  'If tNode <> "" Then
    Set temp = DB.OpenRecordset("Select * from mealplanner where (mealid = " & DragID & ");", dbOpenDynaset)
    If temp("planid") <> 1 Then
'        ReDim parts2(temp.Fields.Count)
'        For i = 0 To temp.Fields.Count - 1
'            parts2(i) = temp.Fields(i)
'        Next i

'        temp.AddNew
'        For i = 1 To temp.Fields.Count - 1
'           temp.Fields(i) = parts2(i)
'        Next i
        temp.Edit
        temp.Fields("PlanID") = 1
        temp.Update
        
    End If
    On Error Resume Next
    temp.Close
    Set temp = Nothing
    Call LoadPlan(CurrentPlanID)
  'End If

End Sub

Private Sub mnuDeleteMeal_Click()
On Error GoTo errhandl
Dim temp As Recordset, temp2 As Recordset
Call SaveWeek
Set temp = DB.OpenRecordset("Select * from mealplanner where mealid = " & DragID & ";", dbOpenDynaset)
If Not temp.EOF Then
     Set temp2 = DB.OpenRecordset("Select * from mealdefinition where mealid=" & temp("mealid") & ";", dbOpenDynaset)
     While Not temp2.EOF
        temp2.Delete
        temp2.MoveNext
     Wend
     temp.Delete
End If
temp.Close
Set temp = Nothing
Call LoadPlan(CurrentPlanID)
Exit Sub
errhandl:
MsgBox "Unable to delete meal." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub mnuDeletePlan_Click()
On Error GoTo errhandl
Dim temp As Recordset
Dim temp2 As Recordset
Call SaveWeek
Dim ret As VbMsgBoxResult
ret = MsgBox("Do you wish to delete all the meals in this plan?", vbYesNoCancel, "")
If ret = vbCancel Then
   Exit Sub
ElseIf ret = vbYes Then
    Set temp = DB.OpenRecordset("Select * from mealplanner where planid = " & CurrentPlanID & ";", dbOpenDynaset)
ElseIf ret = vbNo Then
    Set temp = DB.OpenRecordset("Select * from mealplanner where mealid=-1 and planid = " & CurrentPlanID & ";", dbOpenDynaset)
End If
While Not temp.EOF
  If ret = vbYes Then
     Set temp2 = DB.OpenRecordset("Select * from mealdefinition where mealid=" & temp("mealid") & ";", dbOpenDynaset)
     While Not temp2.EOF
        temp2.Delete
        temp2.MoveNext
     Wend
  End If
  temp.Delete
  temp.MoveNext
Wend
temp.Close
Set temp = Nothing

Call Init
Call OpenWeek(DisplayDate)
Exit Sub
errhandl:
MsgBox "Unable to delete plan." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub mnuDeleteWeekMeal_Click()
On Error Resume Next
Call DeleteSomething
End Sub

Private Sub mnuExit_Click()
On Error Resume Next

frmMain.FlexDiet.OpenDay DisplayDate
Unload Me
End Sub




Private Sub mnuHelp_Click()
On Error GoTo errhandl
Call SaveWeek
HelpWindowHandle = htmlHelpTopic(Me.hWnd, HelpPath, _
         0, "HTML/meal_planner.htm")
Exit Sub
errhandl:
MsgBox "Unable to open help file." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub


Public Sub RefreshMenu()
 
  Call Init
  Call OpenWeek(DisplayDate)

End Sub
Private Sub mnuOpenPlan_Click()
On Error GoTo errhandl2
'  Dim vars As Collection
'  Set vars = New Collection
  CD.CancelError = True
  CD.Filter = "Meal Plan (*.mdb)|*.mdb"
  CD.InitDir = App.path & "\resources\plans"
  CD.ShowOpen
  On Error GoTo errhandl
  WeekHold.MousePointer = 11
  DoEvents
  Call REadScriptMod.ReadScript(CD.Filename, CurrentUser.Username, True)
  DoEvents
  Call RefreshMenu
   WeekHold.MousePointer = 0
   Exit Sub
errhandl:
  MsgBox "Unable to load plan" & vbCrLf & Err.Description, vbOKOnly, ""
  'Set vars = Nothing
errhandl2:
End Sub

Private Sub mnuPrint_Click()
On Error GoTo errhandl
    Call SaveWeek
    PP.Visible = True
    PrintMode = "Print"
   
        
        PP.Navigate2 App.path & "\Resources\temp\PrintMealPlan.htm"
    Exit Sub
errhandl:
MsgBox "Unable to print." & vbCrLf & Err.Description, vbOKOnly, ""

    
End Sub

Private Sub mnuPrintInstructions_Click()
On Error GoTo errhandl
    Call SaveWeek
    PP.Visible = True
    PrintMode = "Print"
   
        PP.Navigate2 App.path & "\Resources\temp\Instructions Page.htm"
Exit Sub
errhandl:
MsgBox "Unable to print." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub mnuRunPlan_Click()
On Error GoTo errhandl

WeekHold.MousePointer = 11
Dim temp As Recordset
Dim PlanFile As String
Set temp = DB.OpenRecordset("Select * from mealplanner where planid = " & CurrentPlanID & " and mealid=-1;", dbOpenDynaset)
PlanFile = temp("Planfile") & " "
PlanFile = Replace(PlanFile, "~FilePath~", App.path, , , vbTextCompare)
Set temp = Nothing
If Trim$(PlanFile) = "" Then Exit Sub
  
Call REadScriptMod.ReadScript(PlanFile, CurrentUser.Username, True)
DoEvents
Call RefreshMenu
WeekHold.MousePointer = 0
Exit Sub
errhandl:
MsgBox "Unable to run plan." & vbCrLf & Err.Description, vbOKOnly, ""
'Resume
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
Call SaveWeek
Call DisplayDay(DisplayDate)
Call frmMain.MakeMealList
End Sub

Private Sub MnuUpload_Click()

On Error GoTo errhandl
CD.CancelError = True
CD.InitDir = App.path & "\resources\plans"
CD.ShowOpen


Exit Sub
errhandl:

End Sub

Private Sub PP_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error GoTo errhandl
If URL = "" Then Exit Sub
Call SaveWeek
If InStr(1, URL, "PrintMealPlan", vbTextCompare) <> 0 Then

 Dim i As Long
  On Error GoTo errhandl


   Dim Doc As HTMLDocument
   Set Doc = PP.document
   Dim TT As HTMLTableCell
   Dim j, junk As String, k As Long, L As Long
   Dim junk2 As String
   If Doc Is Nothing Then Exit Sub
   Set j = Doc.getElementsByTagName("td")
   i = 0
   For Each TT In j
      junk = TT.innerText
      
      'Debug.Print junk
      k = InStr(1, junk, "#", vbBinaryCompare)
      If k <> 0 Then
         junk2 = MiD$(junk, k + 1, 1)
         TT.innerHTML = Replace(junk, "#" & junk2 & "#", DateAdd("d", Val(junk2) - 1, firstSunday))
      End If
      If Val(junk) > 0 And Val(junk) < 8 Then
         k = Val(junk) - 1
         L = Val(Right$(junk, Len(junk) - InStr(1, junk, ",")))
       '  Debug.Print k, L
         junk2 = ""
         With MealCals(k, L)
           If .MealItemID > -1 Then
              junk2 = MealItem(.MealItemID).Caption
           End If
         End With
            
         If junk2 = "Label1" Then junk2 = ""
         TT.innerHTML = junk2
      End If
   Next
   MsgBox "Please set print properties to landscape", vbOKOnly, ""
   'If PrintMode = "Print" Then
   '   PP.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, Nothing, Null
   'ElseIf PrintMode = "Preview" Then
      PP.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
   'End If
   PP.Visible = False
ElseIf InStr(1, URL, "Instructions", vbTextCompare) <> 0 Then
   Dim Body As HTMLBody, temp As Recordset
   Dim Instruc As String, lastMealID As Long
   Set Body = PP.document.Body
   'body.innerHTML
   Set temp = DB.OpenRecordset("SELECT Meals.EntryDate, MealPlanner.Instructions, " _
    & "Meals.MealId, MealPlanner.MealName FROM Meals INNER JOIN MealPlanner " _
    & "ON Meals.MealId = MealPlanner.MealID Where " _
    & "(((Meals.EntryDate) >= #" & firstSunday & "#) and " _
    & "(meals.entrydate<#" & Lastday & "#)) ORDER BY Meals.MealId;", dbOpenDynaset)
   Instruc = ""
   lastMealID = 0
   While Not temp.EOF
     If temp("mealid") <> lastMealID Then
      Instruc = Instruc & "<b><bigger>" & temp("mealname") & "</bigger></b><br>"
      Instruc = Instruc & temp("instructions") & "<br>"
      lastMealID = temp("mealid")
     End If
     temp.MoveNext
   Wend
   Body.innerHTML = Instruc
   Set Body = Nothing
   PP.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, Nothing, Null
  ' PP.Navigate2 App.path & "\Resources\temp\PrintMealPlan.htm"
End If
Exit Sub
errhandl:
   PP.Visible = False
   MsgBox "Unable to print plan." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub



Private Sub TPlans_Click()
'Call TPlans_DblClick
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
   Call DeleteSomething
End If
End Sub

Private Sub DeleteSomething()
On Error Resume Next
  If Hilited > 0 Then
     MealItem(Hilited).Visible = False
     Unload MealItem(Hilited)
     HiLite.Move -sW, -sH
     Dim i As Long, j As Long
     For i = 0 To 7
       For j = 0 To 6
         If MealCals(i, j).MealItemID = Hilited Then
              MealCals(i, j).MealItemID = -1
              MealCals(i, j).ID = -1
              MealCals(i, j).Changed = True
         End If
       Next j
     Next i
nextit:
     Call DoTotals
  End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errhandl
   Dim junk As String, Parts() As String
   junk = Node.FullPath
   If InStr(1, junk, "\", vbBinaryCompare) <> 0 Then
      junk = Node.Key
      Parts = Split(junk, "~")
      DragID = Val(Right$(Parts(0), Len(Parts(0)) - 1))
      DragCals = Val(Parts(1))
      junk = "Calories = " & Round(DragCals) & vbCrLf & MealDesc(STR(DragID))
      MealView.Text = junk
   End If
errhandl:
End Sub


Private Sub weekhold_DragDrop(Source As Control, X As Single, Y As Single)
On Error GoTo errhandl
Dim j As Long, k As Long, L As Long
Dim i As Long, junk As String
Dim T As Single, MiD As Long
T = TDays(1).Height
j = Int(X / sW)
X = j * sW
k = Int((Y - T) / sH)
Y = k * sH + T
'HiLite.Move -sW, -sH
'HiLite.Visible = False
Changed(j) = True

With MealCals(j, k)
   MiD = .MealItemID
   If MiD <> -1 Then
      Unload MealItem(MiD)
      .MealItemID = -1
      .Changed = True
      .Calories = 0
      .ID = -1
   End If
End With
If Not MoveMode Then
    Call NewItem(X, Y, DragID, j, k)
Else
    'MiD = MealCals(j, k).MealItemID
    
    MiD = MealCals(MoveIndex(0), MoveIndex(1)).MealItemID
    MealItem(MiD).Move X, Y, sW, sH
    MealItem(MiD).ForeColor = 0
    Changed(MoveIndex(0)) = True
    
    MealCals(j, k) = MealCals(MoveIndex(0), MoveIndex(1))
    MealCals(j, k).Changed = True
    MealCals(MoveIndex(0), MoveIndex(1)).Changed = True
    MealCals(MoveIndex(0), MoveIndex(1)).MealItemID = -1
    MealCals(MoveIndex(0), MoveIndex(1)).ID = -1
    i = MealCals(j, k).MealItemID
    MoveMode = False
End If

Call DoTotals

Hilited = i


errhandl:
End Sub
Private Sub NewItem(X As Single, Y As Single, DragID As Long, nDay As Long, nMeal As Long)
On Error GoTo errhandl
    If DragID = 0 Then Exit Sub
    Dim i As Long, L As Long
    
    L = MealCals(nDay, nMeal).MealItemID
    If L = -1 Then
       i = MealItem.UBound + 1
       Load MealItem(i)
    Else
       i = L
    End If
    MealItem(i).Move X, Y, sW, sH
    
    MealItem(i).Visible = True
       
    Dim temp As Recordset
    Set temp = DB.OpenRecordset("Select * from mealplanner where (mealid = " & DragID & ");", dbOpenDynaset)
    MealItem(i).Caption = temp("MealName")
    With MealCals(nDay, nMeal)
       .ID = DragID
       .MealItemID = i
       On Error Resume Next
       .Calories = temp("calories")
       .fat = temp("fat")
       .fiber = temp("Fiber")
       .Protein = temp("protein")
       .sugar = temp("Sugar")
       .carbs = temp("carbs")
       .Changed = True
    End With
errhandl:

End Sub
Private Sub DoTotals()
Dim DaysCals(7) As Cals
Dim i As Long, j As Long
On Error GoTo errhandl
For i = 0 To 7
  For j = 0 To 6
   With MealCals(i, j)
     If .MealItemID > -1 Then
        DaysCals(i).Calories = DaysCals(i).Calories + .Calories
        DaysCals(i).carbs = DaysCals(i).carbs + .carbs
        DaysCals(i).fat = DaysCals(i).fat + .fat
        DaysCals(i).fiber = DaysCals(i).fiber + .fiber
        DaysCals(i).Protein = DaysCals(i).Protein + .Protein
        DaysCals(i).sugar = DaysCals(i).sugar + .sugar
     End If
   End With
  Next j
Next i
For i = 0 To 6
  With DaysCals(i)
   If .Calories = 0 Then
    Call Module1.FigurePercentages(Balance(i + 1), NM_Calories, NM_F, NM_S, NM_c, NM_P, NM_FB)
   Else
    Call Module1.FigurePercentages(Balance(i + 1), .Calories, .fat, .sugar, .carbs, .Protein, .fiber)
   End If
   LCalories(i + 1).Caption = "Calories = " & Round(.Calories)
  End With
Next i
errhandl:
End Sub

Private Sub weekhold_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Dim i As Long, junk As String
Dim T As Single
T = TDays(1).Height
X = Int(X / sW) * sW
Y = Int((Y - T) / sH) * sH + T
HiLite.Move X, Y, sW, sH

End Sub



Private Sub mnuNewMeal_Click()
On Error Resume Next
    frmMeals.PlanID = CurrentPlanID
    frmMeals.Show
End Sub


Private Sub mnuNewPlan_Click()
On Error GoTo errhandl
Dim ret As String
ret = InputBox("Please enter new plan name", "")
If Trim$(ret) = "" Then
   MsgBox "You must enter a name.", vbOKOnly, ""
   Exit Sub
End If
  Dim temp As Recordset
  Dim junk As String, junk2 As String, junk3 As Long
  Dim temp2 As Recordset
  Dim ID As Long
  
  Set temp = DB.OpenRecordset("Select * from mealplanner;", dbOpenDynaset)
  Set temp2 = DB.OpenRecordset("Select max(planid) as MAXID from mealplanner;", dbOpenDynaset)
  
  temp.AddNew
  temp("mealid") = -1
  temp("mealname") = ret
  ID = temp2("maxid") + 1
  temp("planid") = ID
  temp.Update
  
  On Error Resume Next
  temp.Close
  temp2.Close
  Set temp = Nothing
  Set temp2 = Nothing

  TPlans.Nodes.Clear
  PlanNodes.Add , , "i1", "Favorites", 2
  On Error GoTo errhandl
  Call Init
  Call LoadPlan(ID)
Exit Sub
errhandl:
MsgBox "Unable to make new plan." & vbCrLf & Err.Description, vbOKOnly, ""
  
End Sub

Private Sub mnuPrintPreview_Click()
On Error GoTo errhandl
  PP.Visible = True
  PrintMode = "Preview"

  PP.Navigate2 App.path & "\Resources\temp\PrintMealPlan.htm"
 
Exit Sub
errhandl:
MsgBox "Unable to print." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub


Private Sub mnuSaveMonth_Click()
On Error Resume Next
SaveWeek
Form1.PlanMode = 0
Form1.Show vbModal, Me

End Sub

Private Sub mnuEditMeal_Click()
On Error GoTo errhandl
Call SaveWeek
frmMeals.PlanID = CurrentPlanID
Call frmMeals.ViewMeal(DragID, CurrentPlanID)
frmMeals.Show
'Call LoadPlan(CurrentPlanID)
Exit Sub
errhandl:
MsgBox "Unable to open meal." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub
Private Sub mnuViewMeal_Click()
On Error GoTo errhandl

   'MsgBox MealDesc(STR(DragID)), vbOKOnly, DragMealName
   Dim temp As Recordset
   Set temp = DB.OpenRecordset("SELECT [MealPlanner].[MealName], [MealPlanner].[Instructions], [ABBREV].[Foodname], [MealDefinition].[Serving], [MealDefinition].[Unit], [MealPlanner].[MealID] " _
    & "FROM (MealPlanner INNER JOIN MealDefinition ON [MealPlanner].[MealID]=[MealDefinition].[MealID]) INNER JOIN ABBREV ON [MealDefinition].[AbbrevID]=[ABBREV].[Index] " _
    & "WHERE ((([MealPlanner].[MealID])=" & DragID & ")); ", dbOpenDynaset)
   Dim junk As String, junk2 As String
   junk2 = temp("Instructions")
   
   junk = temp("Mealname") & vbCrLf
   
   While Not temp.EOF
      junk = junk & "   " & Module1.ConvertDecimalToFraction(temp("Serving")) & " " & temp("Unit") & " " & temp("Foodname") & vbCrLf
      temp.MoveNext
   Wend
   
   junk = junk & vbCrLf & junk2
   
   MsgBox junk, vbOKOnly, ""
   Set temp = Nothing
   Exit Sub
errhandl:
MsgBox "Unable to view meal." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub SetHeads()
On Error Resume Next
    Dim n As Node
    Set n = MealNodes.Add(, , "breakfast", "Breakfast")
    n.Expanded = True
    Set n = MealNodes.Add(, , "snack", "Snacks")
    n.Expanded = True
    Set n = MealNodes.Add(, , "lunch", "Lunch")
    n.Expanded = True
    Set n = MealNodes.Add(, , "dinner", "Dinner")
    n.Expanded = True
    Set n = MealNodes.Add(, , "treat", "Treat")
    n.Expanded = True
    

End Sub




Private Sub TPlans_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
  Dim temp As Recordset
  Dim junk As String, Parts() As String
  Dim tNode As MSComctlLib.Node
  Set tNode = TPlans.HitTest(X, Y)
  If tNode <> "" Then
    Set temp = DB.OpenRecordset("Select * from mealplanner where (mealid = " & DragID & ");", dbOpenDynaset)
    temp.Edit
  
    junk = tNode.Key
    Parts = Split(junk, "~")
    junk = Parts(0)
    temp.Fields("PlanID") = Val(Right$(junk, Len(junk) - 1))
    temp.Update
    On Error Resume Next
    temp.Close
    Set temp = Nothing
    Call LoadPlan(CurrentPlanID)
  End If
End Sub

Private Sub TPlans_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
   Dim node1 As MSComctlLib.Node, junk As String
   Set node1 = TPlans.HitTest(X, Y)
   junk = node1.Key
   CurrentPlanID = Val(Right$(junk, Len(junk) - 1))
   Call LoadPlan(CurrentPlanID)

   PopUpMenu mnuPlans, , TPlans.Left + X, TPlans.Top + Y
End If
End Sub

Private Sub TPlans_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
    Dim junk As String
    junk = Node.Key
    CurrentPlanID = Val(Right$(junk, Len(junk) - 1))
    Call LoadPlan(CurrentPlanID)
End Sub


Private Sub Form_Resize()
Dim swO As Single, shO As Single
swO = sW
shO = sH

Dim i As Long, j As Long
Dim L As Single, W As Single
Dim T As Single, Index As Long
On Error GoTo errhandl
T = Me.ScaleHeight * 0.25
TPlans.Move 0, 0, TreeView1.Width, T
'cView.Move 0, T, TreeView1.Width / 2 - 50
'CEdit.Move TreeView1.Width / 2 + 50, T, TreeView1.Width / 2 - 50, cView.Height
T = T + 50 ' cView.Height
TreeView1.Move 0, T + 50, TreeView1.Width, Me.ScaleHeight - T - 100

'Seperator.Move TreeView1.Width, 0, 50, Me.ScaleHeight

L = TreeView1.Width + 100
WeekHold.Move L, 0, Me.ScaleWidth - L - 100, Me.ScaleHeight - Calendar.Height
sW = WeekHold.ScaleWidth / 7
For i = 1 To 7
  Balance(i).Height = ScaleY(ScaleX(Balance(i).Width, WeekHold.ScaleMode, vbInches), vbInches, WeekHold.ScaleMode)
Next i
Calendar.Move Me.ScaleWidth - Calendar.Width, Me.ScaleHeight - Calendar.Height

MealView.Move L, Me.ScaleHeight - Calendar.Height + 100, Me.ScaleWidth - Calendar.Width - L - 100, Calendar.Height - 200

Dim M As Single
T = WeekHold.ScaleHeight - Balance(1).Height - LCalories(1).Height


For i = 1 To 7
  M = i - 1
  TDays(i).Move sW * M, 0, sW - 10
  LCalories(i).Move sW * M, T, sW - 10
  Balance(i).Move sW * M, T + LCalories(1).Height, sW - 10
  
  DaySep(i).x1 = sW * i
  DaySep(i).X2 = sW * i
  DaySep(i).y1 = 0
  DaySep(i).Y2 = T - 5
  
Next i
T = TDays(1).Height
W = WeekHold.ScaleWidth
TitleLine.y1 = T
TitleLine.Y2 = TitleLine.y1
TitleLine.x1 = 0
TitleLine.X2 = W
sH = (WeekHold.ScaleHeight - T - Balance(1).Height - 5 - LCalories(1).Height) / 6
Dim MealsN(6) As String
MealsN(1) = "Breakfast"
MealsN(2) = "Snack"
MealsN(3) = "Lunch"
MealsN(4) = "Snack"
MealsN(5) = "Dinner"
MealsN(6) = "Treat"
WeekHold.Cls
For i = 1 To 6
    WeekHold.CurrentX = 0
    WeekHold.CurrentY = T + sH * (i) - WeekHold.TextHeight("_~!^")
    WeekHold.Print MealsN(i)
    
    GridLines(i).y1 = T + sH * i
    GridLines(i).Y2 = GridLines(i).y1
    GridLines(i).x1 = 0
    GridLines(i).X2 = W
Next i

DropLabel.Move 0, 0, sW, sH

Dim X As Single, Y As Single
T = TDays(1).Height
On Error GoTo errhandl
For i = 0 To 6
  For j = 0 To 5
    If MealCals(i, j).MealItemID <> -1 Then
     X = i * sW
     Y = j * sH + T
     MealItem(MealCals(i, j).MealItemID).Move X, Y, sW, sH
     DoEvents
    End If
  Next j
Next i
For i = 1 To 7
  DaySep(i).Refresh
Next i
Exit Sub
errhandl:

Err.Clear
End Sub


Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
   StartX = X
   StartY = Y
   Dim node1 As MSComctlLib.Node, junk As String, Parts() As String
   
   Set node1 = TreeView1.HitTest(X, Y)
   If node1 Is Nothing Then Exit Sub
   junk = node1.FullPath
   If InStr(1, junk, "\", vbBinaryCompare) <> 0 Then
      DropLabel.Caption = node1.Text
      MouseMode = True
      DragMeal = Left$(junk, InStr(1, junk, "\", vbBinaryCompare) - 1)
      DragMealName = node1.Text
      junk = node1.Key
      Parts = Split(junk, "~")
      DragID = Val(Right$(Parts(0), Len(Parts(0)) - 1))
      DragCals = Val(Parts(1))
      junk = "Calories = " & Round(DragCals) & vbCrLf & MealDesc(STR(DragID))
      MealView.Text = junk
      mnuEditMeal.Enabled = True
      mnuViewMeal.Enabled = True
      mnuDeleteMeal.Enabled = True
   Else
      mnuEditMeal.Enabled = False
      mnuViewMeal.Enabled = False
      mnuDeleteMeal.Enabled = False
      DragID = -1
   End If
   Set node1 = Nothing
If Button = 2 Then
   MouseMode = False
End If
Exit Sub
errhandl:

End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If MouseMode Then
    If Abs(StartX - X) > 100 Or Abs(StartY - Y) > 100 Then
       HiLite.Move -sW, -sH
       HiLite.Visible = True
       DropLabel.Move X + TreeView1.Left, Y + TreeView1.Top
       DropLabel.Drag vbBeginDrag
       MouseMode = False
    End If
End If
    
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MouseMode = False
If Button = 1 Then
   HiLite.Visible = False
End If
If Button = 2 Then
  If DragID <> -1 Then
    PopUpMenu MnuMealTools, , TreeView1.Left + X, TreeView1.Top + Y
  Else
    PopUpMenu mnuNewMealOnly, , TreeView1.Left + X, TreeView1.Top + Y
  End If
End If
End Sub




Private Sub WeekHold_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim T As Single, j As Single, k As Single
T = TDays(1).Height
j = Int(X / sW)
X = j * sW
k = Int((Y - T) / sH)
Y = k * sH + T

HiLite.Move X, Y, sW, sH
HiLite.Visible = True
End Sub
