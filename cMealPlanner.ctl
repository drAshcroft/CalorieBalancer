VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl cMealPlanner 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
   ScaleHeight     =   5550
   ScaleWidth      =   9735
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   4080
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grab and move this meal"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy This Meal"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut this Meal"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove Meal from planner"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit this Meal"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View this meal"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cMealPlanner.ctx":109A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   4215
      Left            =   5760
      TabIndex        =   0
      Top             =   -720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   9
      Cols            =   8
      ScrollBars      =   0
   End
   Begin CalorieBalance.PieChart PC 
      Height          =   1920
      Index           =   1
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3387
      MaskPicture     =   "cMealPlanner.ctx":13EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Droplabel 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "cMealPlanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit

Private Type DayInfo
   Label As String
   mealID As Long
   Calories As Single
   fat As Single
   carbs As Single
   sugar As Single
   Protein As Single
   fiber As Single
End Type
Dim DSM As Boolean, dsX As Single, dsY As Single, TSM As Boolean
Dim dsMealInfo As String, dsCol As Long, dsRow As Long

Event StartDrag(MealName As String, X As Single, Y As Single)
Event DragMove(X As Single, Y As Single)

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DropMeal(Caption As String, DropDate As Date, MealNumber As Long)
Event ShowPopUp(Meal As String, X As Single, Y As Single)

Dim TitleHeight As Single, RegHeight As Single, GraphHeight As Single

Dim MealDown As String
Dim NullInfo As DayInfo
Dim MealArray() As DayInfo
Dim Totals(7) As DayInfo
Dim sR As Long, SC As Long
Dim MD As Boolean, dsR As Single
Dim firstSunday As Date

Private Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer



Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub PrintMeal(Filename As String)
On Error Resume Next
   Dim i As Long, j As Long, junk As String
   Dim html As String
   html = "<html><body><h1>Meal Plan for " & CurrentUser.Username & "</h1>"
   html = html & "<table width = ""100%"" height=""100%"" BORDER=1 BORDERCOLOR='#000000' " _
   & "CELLPADDING=4 CELLSPACING=0 FRAME=VOID RULES=GROUPS " _
   & "STYLE='page-break-before: always'>"
   
   html = html & "<tr>"
   For i = 0 To FG.Cols - 1
      html = html & "<td  BGCOLOR='#e6e6ff'>" _
      & "<P ALIGN=LEFT STYLE='font-style: normal; text-decoration: none'>" _
      & "<FONT COLOR='#000000'><FONT SIZE=3><B>" & FG.TextMatrix(0, i)
      html = html & "</b></font></font></td>"
   Next i
   html = html & "</tr>"
   
   For j = 1 To FG.Rows - 3
      html = html & "<tr>"
      html = html & "<td  BGCOLOR='#e6e6ff'>" _
      & "<P ALIGN=LEFT STYLE='font-style: normal; text-decoration: none'>" _
      & "<FONT COLOR='#000000'><FONT SIZE=3><B>" & FG.TextMatrix(j, 0)
      html = html & "</b></font></font></td>"
      For i = 1 To FG.Cols - 1
         html = html & "<td>" & FG.TextMatrix(j, i) & "</td>"
      Next i
      html = html & "</tr>"
   Next j
   
   j = FG.Rows - 2
      html = html & "<tr>"
      html = html & "<td  BGCOLOR='#cccccc'>" _
      & "<P ALIGN=LEFT STYLE='font-style: normal; text-decoration: none'>" _
      & "<FONT COLOR='#000000'><FONT SIZE=3><B>" & FG.TextMatrix(j, 0)
      html = html & "</b></font></font></td>"
      For i = 1 To FG.Cols - 1
         html = html & "<td BGCOLOR='#cccccc'>" & FG.TextMatrix(j, i) & "</td>"
      Next i
      html = html & "</tr>"
   
   
   j = FG.Rows - 1
      html = html & "<tr>"
      html = html & "<td  BGCOLOR='#cccccc'>" _
      & "<P ALIGN=LEFT STYLE='font-style: normal; text-decoration: none'>" _
      & "<FONT COLOR='#000000'><FONT SIZE=3><B>" & FG.TextMatrix(j, 0)
      html = html & "</b></font></font></td>"
      For i = 1 To FG.Cols - 1
        With Totals(i)
         html = html & "<td valign='top' BGCOLOR='#cccccc'>"
         html = html & "Protein = " & Round(.Protein) & "<br>"
         html = html & "Fat = " & Round(.fat) & "<br>"
         html = html & "Sugar = " & Round(.sugar) & "<br>"
         html = html & "Carbs = " & Round(.carbs) & "<br>"
         html = html & "Fiber = " & Round(.fiber) & "<br>"
         html = html & "</td>"
        End With
      Next i
      html = html & "</tr>"
   
   
     
   
   html = html & "</table></body></html>"
   Dim ff As Long
   ff = FreeFile
   Open Filename For Output As #ff
   Print #ff, html
   Close #ff
End Sub


Public Sub Delete()
On Error Resume Next
   Dim i As Long, j As Long, junk As String
   Dim sR As Long, eR As Long, SC As Long, eC As Long
   sR = FG.RowSel
   eR = FG.Row
   If sR > eR Then
     i = eR
     eR = sR
     sR = i
   End If
   SC = FG.ColSel
   eC = FG.Col
   If SC > eC Then
     i = eC
     eC = SC
     SC = i
   End If
   junk = ""
   For i = SC To eC
     For j = sR To eR
       Call DeleteMeal(j, i)
     Next j
   Next i
   Call DoTotals
End Sub
Public Sub Paste()
On Error Resume Next
   Dim i As Long, j As Long, junk As String
   Dim sR As Long, eR As Long, SC As Long, eC As Long
   sR = FG.RowSel
   eR = FG.Row
   If sR > eR Then
     i = eR
     eR = sR
     sR = i
   End If
   SC = FG.ColSel
   eC = FG.Col
   If SC > eC Then
     i = eC
     eC = SC
     SC = i
   End If
   junk = Clipboard.GetText
   Dim lines() As String, Parts() As String, junk2 As String, jj As String
   Dim k As Long
   lines = Split(junk, vbCrLf)
   For i = 0 To UBound(lines)
     Parts = Split(lines(i), vbTab)
     For j = 0 To UBound(Parts)
       junk = Trim$(Parts(j)) 'Trim$(ReverseString(Parts(j)))
       If Right$(junk, 1) = ")" Then
          junk2 = ""
          For k = Len(junk) - 1 To 0 Step -1
             jj = Mid$(junk, k, 1)
             If jj = "(" Then Exit For
             junk2 = jj & junk2
          Next k
          If Val(junk2) <> 0 Then
             FG.TextMatrix(i + sR, j + SC) = LoadMealInfo("~~~ ~~~" & Val(junk2), i + sR, j + SC)
          Else
             Call DeleteMeal(i + sR, j + SC)
          End If
       End If
     Next j
   Next i
   Call DoTotals
End Sub
Public Sub Cut()
On Error Resume Next
  Call Copy
  Call Delete
End Sub
Public Sub Copy()

On Error Resume Next
   Dim i As Long, j As Long, junk As String
   Dim sR As Long, eR As Long, SC As Long, eC As Long
   sR = FG.RowSel
   eR = FG.Row
   If sR > eR Then
     i = eR
     eR = sR
     sR = i
   End If
   SC = FG.ColSel
   eC = FG.Col
   If SC > eC Then
     i = eC
     eC = SC
     SC = i
   End If
   junk = ""
   For j = sR To eR
     For i = SC To eC
          junk = junk & FG.TextMatrix(j, i) & "(" & MealArray(j, i).mealID & ")" & vbTab
     Next
     junk = junk & vbCrLf
   Next
   Clipboard.Clear
   Clipboard.SetText junk
End Sub
Private Sub DoTotals()
On Error Resume Next
  Dim i As Long, j As Long
  For i = 1 To 7
    Totals(i) = NullInfo
    With Totals(i)
      For j = 1 To FG.Rows - 3
         .Calories = .Calories + MealArray(j, i).Calories
         .carbs = .carbs + MealArray(j, i).carbs
         .fat = .fat + MealArray(j, i).fat
         .Protein = .Protein + MealArray(j, i).Protein
         .sugar = .sugar + MealArray(j, i).sugar
         .fiber = .fiber + MealArray(j, i).fiber
      Next j
      Call Module1.FigurePercentages(PC(i), .Calories, .fat, .sugar, .carbs, .Protein, .fiber)
      FG.TextMatrix(FG.Rows - 2, i) = Round(.Calories, 1)
    End With
  Next i
End Sub
Public Function Clear()
On Error Resume Next
    Erase MealArray
    ReDim MealArray(6, 7)
    FG.Clear
    FG.Rows = 9
    FG.Cols = 8
  Dim RS As Recordset, rs2 As Recordset
  Dim i As Long, j As Long
  Set RS = DB.OpenRecordset("SELECT * from meals " _
    & "where meals.user='" & CurrentUser.Username & _
    "' and entrydate>=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
  While Not RS.EOF
    Set rs2 = DB.OpenRecordset("select * from daysinfo where user ='" & CurrentUser.Username _
       & "' and date>=#" & FixDate(DisplayDate) & "# and mealid=" & RS("id") & ";", dbOpenDynaset)
    If rs2.EOF Then
      RS.Delete
    End If
    RS.MoveNext
  Wend
    
End Function
Public Function OpenWeek()

  
    
On Error Resume Next
    Erase MealArray
    ReDim MealArray(6, 7)

  FG.Clear
    FG.Rows = 9
    FG.Cols = 8
  Dim RS As Recordset, LastSat As Date
  Dim i As Long, j As Long
  firstSunday = Module1.FindFirstDay(DisplayDate)
  LastSat = DateAdd("d", 6, firstSunday)
  Set RS = DB.OpenRecordset("SELECT Meals.*, MealPlanner.* " _
    & "FROM Meals INNER JOIN MealPlanner ON Meals.MealId = MealPlanner.MealID " _
    & "where meals.user='" & CurrentUser.Username & _
    "' and entrydate>=#" & FixDate(firstSunday) & "# and entrydate<=#" & FixDate(LastSat) & "#;", dbOpenDynaset)
  While Not RS.EOF
     i = Abs(DateDiff("d", firstSunday, RS("entrydate"))) + 1
     j = RS("mealnumber") + 1
     If j + 2 > FG.Rows Then
        FG.Rows = j + 3
        Dim tMA() As DayInfo, k As Long, L As Long
        ReDim tMA(j, 7)
        For k = 0 To UBound(MealArray, 1)
          For L = 0 To UBound(MealArray, 2)
            tMA(k, L) = MealArray(k, L)
          Next L
        Next k
        Erase MealArray
        ReDim MealArray(j, 7)
        For k = 0 To UBound(MealArray, 1)
          For L = 0 To UBound(MealArray, 2)
            MealArray(k, L) = tMA(k, L)
          Next L
        Next k
        Erase tMA
     End If
     FG.TextMatrix(j, i) = LoadMealInfo("~~~ ~~~" & RS("meals.mealid"), j, i, False)
     RS.MoveNext
  Wend
  
   FG.TextArray(1) = "Sun " & DateAdd("d", 0, firstSunday)
   FG.TextArray(2) = "Mon " & DateAdd("d", 1, firstSunday)
   FG.TextArray(3) = "Tues " & DateAdd("d", 2, firstSunday)
   FG.TextArray(4) = "Wed " & DateAdd("d", 3, firstSunday)
   FG.TextArray(5) = "Thurs " & DateAdd("d", 4, firstSunday)
   FG.TextArray(6) = "Fri " & DateAdd("d", 5, firstSunday)
   FG.TextArray(7) = "Sat " & DateAdd("d", 6, firstSunday)
  Call DoTotals
  Call UserControl_Resize
End Function
Public Function DeleteMeal(sR As Long, SC As Long) As String
On Error GoTo errhandl
  DeleteMeal = MealArray(sR, SC).Label
  MealArray(sR, SC) = NullInfo
  FG.TextMatrix(sR, SC) = ""
  Call SaveMealInfo(0, SC, sR) ' sR, sC)
  'Call DoTotals
errhandl:
End Function
Private Function SaveMealInfo(Func As Long, nDay As Long, ByVal nMeal As Long)
On Error Resume Next
 Dim RS As Recordset, d As Date, rs2 As Recordset
 nMeal = nMeal - 1
 d = DateAdd("d", nDay - 1, firstSunday)
 
 If Func = 1 Then
     RaiseEvent DropMeal(MealArray(nMeal + 1, nDay).Label, d, nMeal) ' (nmeal,nday)
 Else
    Set RS = DB.OpenRecordset("select * from meals where user='" & CurrentUser.Username & "' " _
         & " and entrydate=#" & FixDate(d) & "# " _
         & " and mealnumber=" & nMeal & ";", dbOpenDynaset)
    Set rs2 = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username & "' " _
         & " and date=#" & d & "# " _
         & " and mealid=" & RS("id") & ";", dbOpenDynaset)
     If Not RS.EOF Then RS.Delete
     If Not rs2 Is Nothing Then
       While Not rs2.EOF
          rs2.Delete
          rs2.MoveNext
          
       Wend
     End If
     rs2.Close
     RS.Close
     Set RS = Nothing
     Set rs2 = Nothing
 End If
   
End Function
Private Function LoadMealInfo(Caption As String, r As Long, c As Long, Optional SaveInfo As Boolean = True) As String
On Error GoTo errhandl

  Dim RS As Recordset
  Dim Parts() As String
   
   Parts = Split(Caption, "~~~")
   'FG.TextMatrix(mR, mC) = parts(1)
   'Call LoadMealInfo(Val(parts(2)), mR, mC)
   
  Set RS = DB.OpenRecordset("Select * from mealplanner where mealid=" & Parts(2) & ";", dbOpenDynaset)
  With MealArray(r, c)
    .Label = "~~~" & RS("mealname") & "~~~" & RS("mealID")
    .mealID = Val(Parts(2))
    .Calories = RS("calories")
    .carbs = RS("carbs")
    .fat = RS("fat")
    .fiber = RS("fiber")
    .Protein = RS("protein")
    .sugar = RS("sugar")
   
    LoadMealInfo = RS("mealname")   ' & vbCrLf & "Cals: " & .Calories
  End With
  RS.Close
  Set RS = Nothing
  If SaveInfo Then Call SaveMealInfo(1, c, r)
  Exit Function
errhandl:
 
End Function
Public Sub DragDrop(Source, X As Single, Y As Single)
Dim mC As Long, mR As Long
On Error GoTo errhandl

DSM = False
mC = FG.MouseCol
mR = FG.MouseRow
'On Error GoTo errhandl
If mC = 0 Then Exit Sub
If mR > FG.Rows - 3 Then Exit Sub
FG.HighLight = flexHighlightAlways
If Left$(Source.Caption, 3) = "~~~" Then
   FG.ColSel = mC
   FG.RowSel = mR
   FG.TextMatrix(mR, mC) = LoadMealInfo(Source.Caption, mR, mC)
End If
FG.Col = FG.MouseCol
FG.Row = FG.MouseRow
Call DoTotals
Exit Sub
errhandl:
If DoDebug Then MsgBox Err.Description
End Sub

Private Sub FG_Click()
Call FG_EnterCell
End Sub

Private Sub FG_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
   Call DragDrop(Source, X, Y)
End Sub


Private Sub FG_EnterCell()


    On Error GoTo Err_Proc
If FG.Col > 0 And FG.Row > 0 And FG.Row < (FG.Rows - 2) And FG.Text <> "" And FG.Row = FG.RowSel And FG.Col = FG.ColSel Then
  If FG.Col = FG.Cols - 1 Then
    Toolbar1.Move FG.Width - Toolbar1.Width, FG.RowPos(FG.Row + 1)
  Else
    Toolbar1.Move FG.ColPos(FG.Col), FG.RowPos(FG.Row + 1)
  End If
  Toolbar1.Visible = True
  Toolbar1.ZOrder
Else
  Toolbar1.Visible = False
End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "cMealPlanner", "FG_EnterCell", Err.Description
    Resume Exit_Proc


End Sub

Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then Call Delete
If KeyCode = 71 Then
Dim X As Single, Y As Single
  dsRow = FG.MouseRow
  dsCol = FG.MouseCol
  X = FG.ColPos(dsCol)
  Y = FG.RowPos(dsRow)
         FG.HighLight = flexHighlightNever
       Dim RS As Recordset
       Set RS = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username _
           & "' and date=#" & FixDate(DateAdd("d", dsCol - 1, firstSunday)) & "# and meal=" & (dsCol - 1) & ";", dbOpenDynaset)
       While Not RS.EOF
         RS.Delete
         RS.MoveNext
       Wend
       Set RS = Nothing
       With MealArray(dsRow, dsCol)
          .Calories = 0
          .carbs = 0
          .fat = 0
          .fiber = 0
          .Label = ""
          .mealID = 0
          .Protein = 0
          .sugar = 0
       End With
       FG.TextMatrix(dsRow, dsCol) = ""
       Call DoTotals
       RaiseEvent StartDrag(dsMealInfo, X, Y)
End If
End Sub

Private Sub FG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  dsX = X
  dsY = Y

If Button = 1 And (GetKeyState(vbKeyShift) And &H1000) Then
  DSM = True
  dsRow = FG.MouseRow
  dsCol = FG.MouseCol
ElseIf Button = 1 Then
  TSM = True
  
End If
Toolbar2.Visible = False
  dsMealInfo = MealArray(FG.MouseRow, FG.MouseCol).Label
  
End Sub

Private Sub FG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  Dim dd As Single
  
  If TSM Then
    dd = ((X - dsX) ^ 2 + (Y - dsY) ^ 2) ^ 0.5
    If dd > FG.RowHeight(1) Then
       Toolbar1.Visible = False
       Toolbar2.Visible = False
       TSM = False
    ElseIf dd < FG.RowHeight(1) And dsX <> 0 Then
       
    End If
    
  ElseIf DSM Then
    dd = ((X - dsX) ^ 2 + (Y - dsY) ^ 2) ^ 0.5
    
    If dd > 250 Then
       FG.HighLight = flexHighlightNever
       Dim RS As Recordset
       Set RS = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username _
           & "' and date=#" & FixDate(DateAdd("d", dsCol - 1, firstSunday)) & "# and meal=" & (dsCol - 1) & ";", dbOpenDynaset)
       While Not RS.EOF
         RS.Delete
         RS.MoveNext
       Wend
       Set RS = Nothing
       With MealArray(dsRow, dsCol)
          .Calories = 0
          .carbs = 0
          .fat = 0
          .fiber = 0
          .Label = ""
          .mealID = 0
          .Protein = 0
          .sugar = 0
       End With
       FG.TextMatrix(dsRow, dsCol) = ""
       Call DoTotals
       RaiseEvent StartDrag(dsMealInfo, X, Y)
       
    End If
  
   
  End If
End Sub

Private Sub FG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
  Dim junk As String
  'junk = MealArray(FG.MouseRow, FG.MouseCol).Label '= "~~~" & rs("mealname") & "~~~" & rs("mealID")
  RaiseEvent ShowPopUp(dsMealInfo, X, Y)
End If


If DSM Then
FG.HighLight = flexHighlightAlways
DSM = False
End If
End Sub



Private Sub FG_SelChange()


    On Error GoTo Err_Proc
Dim X As Single, Y As Single

If FG.Col <> FG.ColSel And FG.Row <> FG.RowSel Then
  Toolbar1.Visible = False
  Toolbar2.Visible = True
  Toolbar2.ZOrder
  If FG.Col < FG.ColSel Then
     X = FG.ColSel
  Else
     X = FG.Col
  End If
  If FG.Row < FG.RowSel Then
     Y = FG.RowSel
  Else
     Y = FG.Row
  End If
  If FG.Col = FG.Cols - 1 Then
    Toolbar2.Move FG.Width - Toolbar2.Width, FG.RowPos(Y + 1)
  Else
    Toolbar2.Move FG.ColPos(X), FG.RowPos(Y + 1)
  End If

End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "cMealPlanner", "FG_SelChange", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
  Dim junk As String
  Dim RS As Recordset
  Select Case Button.Index
     Case 1
       DSM = True
       FG.HighLight = flexHighlightNever
       Set RS = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username _
           & "' and date=#" & FixDate(DateAdd("d", FG.Col - 1, firstSunday)) & "# and meal=" & (FG.Col - 1) & ";", dbOpenDynaset)
       While Not RS.EOF
         RS.Delete
         RS.MoveNext
       Wend
       Set RS = Nothing
       junk = MealArray(FG.Row, FG.Col).Label
       With MealArray(FG.Row, FG.Col)
          .Calories = 0
          .carbs = 0
          .fat = 0
          .fiber = 0
          .Label = ""
          .mealID = 0
          .Protein = 0
          .sugar = 0
       End With
       FG.TextMatrix(FG.Row, FG.Col) = ""
       Call DoTotals
       RaiseEvent StartDrag(junk, FG.ColPos(FG.Col), FG.RowPos(FG.Row))
     Case 2
       Call Copy
     Case 3
       Call Cut
     Case 4
       Call Delete
     Case 5
        Dim mealID As Long, PlanID As Long
        
        junk = MealArray(FG.Row, FG.Col).mealID
        If junk = "0" Or junk = "" Then Exit Sub
        Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & junk & ";", dbOpenDynaset)
        mealID = RS("mealid")
        PlanID = RS("planid")
        If Not RS.EOF Then
           RS.Close
           Call frmMeals.ViewMeal(mealID, PlanID)
           Call frmMeals.Show
        End If
        Set RS = Nothing
     Case 6
        Call ViewMeal(MealArray(FG.Row, FG.Col).mealID & "")
  End Select
  Call FG_EnterCell
End Sub
Public Sub ViewMeal(mealID As String)

On Error Resume Next
Dim RS As Recordset, TT As String
Dim rs2 As Recordset
Dim ret As VbMsgBoxResult
   mealID = Trim$(mealID)
   If mealID = "" Or mealID = "0" Then Exit Sub

   Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & mealID & ";", dbOpenDynaset)
   mealID = RS("mealid")
   
   
   Set rs2 = DB.OpenRecordset("SELECT MealDefinition.*, Abbrev.Foodname " _
    & "FROM MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID = Abbrev.Index " _
    & "where mealid=" & mealID & ";", dbOpenDynaset)
   TT = ""
   While Not rs2.EOF
     TT = TT & rs2("serving") & " " & rs2("unit") & " " & rs2("Foodname") & vbCrLf
     rs2.MoveNext
   Wend
   TT = TT & vbCrLf & RS("instructions")
   TT = TT & vbCrLf & "Calories = " & Round(RS("Calories"), 1)
   TT = TT & vbCrLf & "Fat(g) = " & Round(RS("fat"), 1)
   TT = TT & vbCrLf & "Carbs(g) = " & Round(RS("Carbs"), 1)
   TT = TT & vbCrLf & "Protein(g) = " & Round(RS("protein"), 1)
   TT = TT & vbCrLf & "fiber(g) = " & Round(RS("fiber"), 1)
   RS.Close
   MsgBox TT, vbOKOnly, ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
  Case 1
    Call Copy
  Case 2
    Call Cut
  Case 3
    Call Delete
End Select
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
   Dim i As Long
   FG.Row = 0
   For i = 0 To FG.Cols - 1
      FG.Col = i
      FG.CellFontSize = 10
      FG.CellFontBold = True
   Next i
   RegHeight = FG.RowHeight(0)
   TitleHeight = UserControl.TextHeight("~_I") * 1.2
   FG.RowHeight(0) = TitleHeight
   FG.TextArray(1) = "Sun"
   FG.TextArray(2) = "Mon"
   FG.TextArray(3) = "Tues"
   FG.TextArray(4) = "Wed"
   FG.TextArray(5) = "Thurs"
   FG.TextArray(6) = "Fri"
   FG.TextArray(7) = "Sat"
   
   ReDim MealArray(6, 7)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim i As Long, H As Single, W As Single, T As Single
Dim SWidth As Single
If PC.UBound = 1 Then
  For i = 2 To 7
    Load PC(i)
    PC(i).Visible = True
    PC(i).ZOrder
  Next i
End If
   FG.TextMatrix(1, 0) = "Breakfast"
   FG.TextMatrix(2, 0) = "Brunch"
   FG.TextMatrix(3, 0) = "Lunch"
   FG.TextMatrix(4, 0) = "Snack"
   FG.TextMatrix(5, 0) = "Dinner"
   FG.TextMatrix(6, 0) = "Treat"
 '  If FG.Rows > 8 Then
    For i = 7 To FG.Rows - 3
     FG.TextMatrix(i, 0) = "Extra"
    Next i
  ' End If
   FG.TextMatrix(i, 0) = "Total Cals"
   FG.TextMatrix(i + 1, 0) = "Balance"
SWidth = 0
FG.Row = 1
FG.Col = 0
UserControl.FontName = FG.CellFontName
UserControl.FontBold = FG.CellFontBold
UserControl.FontSize = FG.CellFontSize
For i = 0 To FG.Rows - 1
  W = TextWidth(FG.TextMatrix(i, 0)) * 1.2
  If W > SWidth Then SWidth = W
Next i

FG.Move 0, 0, ScaleWidth - 50, ScaleHeight
FG.RowHeight(0) = TitleHeight
FG.RowHeight(FG.Rows - 2) = TitleHeight

FG.ColWidth(0) = SWidth


W = (ScaleWidth - SWidth - 100) / 7
FG.RowHeight(FG.Rows - 1) = W

H = (ScaleHeight - 2 * TitleHeight - W - 150) / (FG.Rows - 3)

For i = 1 To FG.Cols - 1
  FG.ColWidth(i) = W
Next i

For i = 1 To FG.Rows - 3
  FG.RowHeight(i) = H
Next i

Dim L As Single
T = FG.RowPos(FG.Rows - 1) + 50
L = FG.ColPos(1) + 50
For i = 0 To 6
   PC(i + 1).Move L + W * i, T, W, W
   PC(i + 1).ZOrder
Next i


End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
