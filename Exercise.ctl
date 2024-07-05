VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Exercise 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10905
   ScaleHeight     =   6315
   ScaleWidth      =   10905
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   2400
      TabIndex        =   6
      Text            =   "Exercise Calories will vary with your current weight."
      Top             =   4440
      Width           =   5655
   End
   Begin CalorieBalance.ExerciseBox TEnter 
      Height          =   2175
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Text1"
   End
   Begin VB.TextBox Totals 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Total Calories Burned"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Totals 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1800
      Width           =   1335
   End
   Begin CalorieBalance.AutoCompleteEX EList 
      Height          =   3255
      Left            =   4920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5741
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
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   8280
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox TEnter2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Daily 
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   12632064
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
End
Attribute VB_Name = "Exercise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit

Dim IDs() As Long, Formula() As String
Dim Calories() As Single
Dim Disabled() As Boolean

Dim CurGrid As Long
Dim mDB As Database

Dim LRow As Long, Lcol As Long
Dim EWeights As Recordset
Dim PopUpMenu As Object
Dim CaloriesMode As Boolean
Dim TempText() As String
Dim Override As Boolean
Dim Curdate As Date
Dim Opening As Boolean

Dim DownX As Long, DownY As Long ' records location of mousedown for blocking data
'Default Property Values:
Const m_def_AddLastWeeks = False
Const m_def_UpdateValues = True
'Property Variables:
Dim m_AddLastWeeks As Boolean
Dim m_UpdateValues As Boolean
Event TodaysCalories(NewCals As Single)
Dim TodaysCol As Long

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub SetBackGround(back)
On Error Resume Next
   UserControl.BackColor = back
   
End Sub

Public Function PrintExercise(Filename As String)
On Error GoTo errhandl
  Dim ff As Long, i As Long, j As Long
  ff = FreeFile
  Open Filename For Output As #ff
    Print #ff, "<html><body>"
    Print #ff, "<h1>Exercises for " & CurrentUser.Username & "</h1>"
    Print #ff, "<table width = ""100%"" height=""100%"" border=""1"" cellspacing=0 cellpading=0>"
    Print #ff, "<tr>"
    For i = 0 To Daily.Cols - 1
       Print #ff, "<td borderColor=white><b>&nbsp;" & Daily.TextMatrix(0, i) & "</b></td>"
    Next i
    Print #ff, "</td>"
    Dim clr As Long, clr2 As Long
    For i = 1 To Daily.Rows - 1
       Print #ff, "<tr>"
       Daily.Row = i
       For j = 0 To Daily.Cols - 1
         Daily.Col = j
         clr = Daily.CellBackColor
         clr2 = Daily.CellForeColor
         If clr = 0 Then clr = vbWhite
         Print #ff, "<td borderColor=black bgcolor =""#" & WebHexColor(clr) & """><font color =""#" & WebHexColor(clr2) & """>&nbsp;" & Daily.TextMatrix(i, j) & "</font></td>"
         
       Next j
       Print #ff, "</tr>"
    Next i
    For i = 1 To 5
       Print #ff, "<tr>"
       For j = 0 To Daily.Cols - 1
         Print #ff, "<td borderColor=black >&nbsp;</td>"
       Next j
       Print #ff, "</tr>"
    Next i
    
    Print #ff, "<tr><td borderColor=white><B>Totals</b> </td>"
    For i = 1 To 7
       Print #ff, "<td borderColor=black>&nbsp;" & Round(Val(Totals(i))) & "</td>"
    Next i
  Close #ff
  Exit Function
errhandl:
  MsgBox "Unable to print exercise." & vbCrLf & Err.Description, vbOKOnly, ""
End Function
Public Sub ShowUnits()
On Error Resume Next
Dim i As Long, j As Long
For j = 1 To Daily.Rows - 2
  For i = 1 To 7
    Daily.TextMatrix(j, i) = TempText(i, j)
  Next i
Next j
CaloriesMode = False
  
End Sub
Public Sub ShowCalories()
  On Error Resume Next
  Dim i As Long, j As Long
  ReDim TempText(7, Daily.Rows - 2)
  CaloriesMode = True
  For j = 1 To Daily.Rows - 2
      For i = 1 To 7
        TempText(i, j) = Daily.TextMatrix(j, i)
        If Disabled(i, j) Then
           Daily.TextMatrix(j, i) = Round(Calories(i, j), 1)
        Else
           Daily.TextMatrix(j, i) = 0
        End If
      Next i
  Next j
End Sub
Private Function TranslateText(T As String) As Long
On Error Resume Next
  Dim temp As Recordset, first As Boolean
  Set temp = DB.OpenRecordset("Select index from abbrevexercise " & _
          " where exercisename = '" & T & "';", dbOpenDynaset)
  While (Not temp.EOF) And (Not first)
     TranslateText = temp.Fields("index")
     first = True
  Wend
  
  temp.Close
  Set temp = Nothing
  
End Function
Public Sub PasteRow()
On Error GoTo errhandl
 Dim lines() As String
  Dim junk As String, Parts() As String, re As Long, RS As Long
  Dim i As Long
  
Call Daily_LeaveCell
EList.Visible = False
TEnter.Visible = False

junk = Clipboard.GetText


  re = Daily.RowSel
  RS = Daily.Row
  
  Call SaveWeek(DisplayDate)
  Dim Week As Date
  Week = FindFirstDay(DisplayDate)
  
 If InStr(1, junk, vbCrLf, vbBinaryCompare) <> 0 Then
    lines = Split(junk, vbCrLf, , vbBinaryCompare)
 Else
    ReDim lines(0)
    lines(0) = junk
 End If
 'move all the data up
 Dim r As Recordset, A As Recordset
 Set r = DB.OpenRecordset("select * from exerciselog where week=#" & FixDate(Week) & "# and user='" & CurrentUser.Username & "' and [order]>=" & RS & ";", dbOpenDynaset)
 While Not r.EOF
    r.Edit
    r("order") = r("order") + UBound(lines) + 1
    r.Update
    r.MoveNext
 Wend
 
 'now enter the new data
 Dim j As Long
 For i = 0 To UBound(lines)
    Parts = Split(lines(i), vbTab)
    If LCase$(Parts(0)) <> "exercise" Then
       Set A = DB.OpenRecordset("select * from abbrevexercise where exercisename='" & Parts(0) & "';", dbOpenDynaset)
       If A.EOF = False Then
          r.AddNew
          r("week") = Week 'FixDate(Week)
          r("user") = CurrentUser.Username
          r("exerciseid") = A("index")
          junk = ""
          For j = 1 To UBound(Parts)
            junk = junk & Parts(j) & "~"
          Next j
          r("weekinfo") = junk
          r("order") = i + RS
          r.Update
       End If
       A.Close
       Set A = Nothing
    End If
 Next i

  
  Call OpenWeek(DisplayDate)
errhandl:
End Sub
Public Sub CutRow()
On Error Resume Next
  Call Copy
  Call DeleteRow
End Sub
Public Sub Copy()
On Error GoTo errhandl
Call Daily_LeaveCell
EList.Visible = False
TEnter.Visible = False


Dim junk As String, i As Long, j As Long
Dim RS As Long, re As Long
RS = Daily.Row
re = Daily.RowSel
  For j = 0 To Daily.Cols - 1
     junk = junk & Daily.TextArray(j) & vbTab
  Next j
  junk = junk & vbCrLf
  For i = RS To re
    For j = 0 To Daily.Cols - 1
       Daily.Row = i
       Daily.Col = j
       If Daily.CellForeColor = RGB(200, 200, 200) Then
          junk = junk & "*" & Daily.TextMatrix(i, j) & vbTab
       ElseIf Daily.CellBackColor = RGB(50, 50, 50) Then
          junk = junk & "-" & Daily.TextMatrix(i, j) & vbTab
       Else
          junk = junk & Daily.TextMatrix(i, j) & vbTab
       End If
      
    Next j
    If i <> re Then junk = junk & vbCrLf
  Next i
  'put it all on the clipboard
  Clipboard.Clear
  Clipboard.SetText junk
  Daily.Row = RS
  Daily.RowSel = re
 ' Daily.RowSel = Daily.Row
 ' Daily.ColSel = Daily.Col
  Exit Sub
errhandl:
  MsgBox "Unable to copy." & vbCrLf & Err.Description, vbOKOnly, ""
End Sub
Public Sub SetPopUp(PopUp As Object)
On Error Resume Next
  Set PopUpMenu = PopUp
End Sub
Public Sub InsertRow(Optional nRows As Long = 1)
On Error GoTo errhandl
Call Daily_LeaveCell
EList.Visible = False
TEnter.Visible = False

  
  Dim i As Long, j As Long
  Daily.Rows = Daily.Rows + nRows
  ReDim Preserve IDs(UBound(IDs) + nRows)
  ReDim Preserve Formula(UBound(Formula) + nRows)
  ReDim Preserve Calories(8, UBound(Calories, 2) + nRows)
  ReDim Preserve Disabled(8, UBound(Calories, 2) + nRows)
  
  For i = Daily.Rows - 1 To LRow + 1 Step -1
     IDs(i) = IDs(i - nRows)
     Formula(i) = Formula(i - nRows)
     For j = 0 To 7
       Calories(j, i) = Calories(j, i - nRows)
       Disabled(j, i) = Disabled(j, i - nRows)
       Daily.TextMatrix(i, j) = Daily.TextMatrix(i - nRows, j)
     Next j
  Next i
  On Error GoTo errhandl
  For j = 0 To nRows - 1
      IDs(LRow + j) = 0
      Formula(LRow + j) = ""
      For i = 0 To 7
         Calories(i, LRow + j) = 0
         Daily.TextMatrix(LRow + j, i) = ""
         Disabled(i, LRow + j) = False
      Next i
  Next j
  For i = 1 To 7
     Call DoTotal(i)
  Next i
Exit Sub
errhandl:
MsgBox "Unable to insert row." & vbCrLf & Err.Description, vbOKOnly, ""
End Sub
Public Sub DeleteRow()
On Error GoTo errhandl
  Dim i As Long, j As Long, RS As Long, re As Long
  re = Daily.RowSel
  RS = Daily.Row
  If re < RS Then
     i = RS
     RS = re
     re = i
  End If
  
  Call Daily_LeaveCell
  EList.Visible = False
  TEnter.Visible = False
  
  're = Daily.RowSel
  'rs = Daily.Row
  If Daily.ColSel <> 0 And Daily.Col <> 0 Then
    Dim Ce As Long, Cs As Long
    Ce = Daily.ColSel
    Cs = Daily.Col
    If Ce < Cs Then
       i = Cs
       Cs = Ce
       Ce = i
    End If
    
    For j = Cs To Cs
        For i = RS To re
      
         Daily.TextMatrix(i, j) = ""
         Calories(j, i) = 0
         
      Next
      Call DoTotal(j)
    Next
    Call SaveWeek(DisplayDate)
    
    Exit Sub
  End If
  
  Call SaveWeek(DisplayDate)
  
  Dim r As Recordset, Week As Date
  Week = FindFirstDay(DisplayDate)
  Set r = DB.OpenRecordset("select * from exerciselog where week=#" & FixDate(Week) & "# and user='" _
    & CurrentUser.Username & "';", dbOpenDynaset)
  While Not r.EOF
    If r("order") >= RS And r("order") <= re Then
       r.Delete
    End If
    r.MoveNext
  Wend
    
  r.Close
  Set r = Nothing
errhandl:
  On Error Resume Next
  Call OpenWeek(DisplayDate)
End Sub
Public Sub DeleteRowO()
On Error GoTo errhandl
  Call Daily_LeaveCell
  EList.Visible = False
  TEnter.Visible = False
  
  Dim i As Long, j As Long
  Dim RS As Long
  Dim re As Long, nRows As Long
  
  RS = Daily.Row
  re = Daily.RowSel
  If RS > re Then
    i = RS
    RS = re
    re = i
  End If
  
  nRows = re - RS + 1
  
  For i = RS To Daily.Rows - nRows - 1
     IDs(i) = IDs(i + nRows)
     Formula(i) = Formula(i + nRows)
     For j = 0 To 7
       Calories(j, i) = Calories(j, i + nRows)
       Disabled(j, i) = Disabled(j, i + nRows)
       Daily.TextMatrix(i, j) = Daily.TextMatrix(i + nRows, j)
     Next j
  Next i
    
  Daily.Rows = Daily.Rows - nRows
  ReDim Preserve IDs(UBound(IDs) - nRows)
  ReDim Preserve Formula(UBound(Formula) - nRows)
  ReDim Preserve Calories(8, UBound(Calories, 2) - nRows)
  ReDim Preserve Disabled(8, UBound(Calories, 2) - nRows)
  For i = 1 To 7
    Call DoTotal(i)
  Next i
  Exit Sub
errhandl:
MsgBox "Unable to insert row." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub
Public Sub SaveWeek(SaveDate As Date)
On Error GoTo errhandl
Call Daily_LeaveCell
Opening = True
EList.Visible = False
TEnter.Visible = False

If CurrentUser.Username = "" Then Exit Sub
  Curdate = SaveDate
  Dim temp As Recordset, Week As Date, i As Long, k As Long
  Dim junk As String
  Week = FindFirstDay(SaveDate)
  Set temp = DB.OpenRecordset("Select * from ExerciseLog where ((Week = #" & FixDate(Week) & "#) and (User = """ & CurrentUser.Username & """)) ;", dbOpenDynaset)
  If temp.RecordCount <> 0 Then
    While Not temp.EOF
     temp.Delete
     temp.MoveNext
    Wend
  End If
  Dim junk2 As String
  For i = 1 To Daily.Rows - 1
    Daily.Row = i
    Daily.Col = 0
    If IDs(i) <> 0 And Trim$(Daily.TextMatrix(i, 0)) <> "" Then
      temp.AddNew
      temp.Fields("week") = Week 'FixDate(Week)
      temp.Fields("user") = CurrentUser.Username
      temp.Fields("Order") = i
      If IDs(i) = -1111 Then
        If Trim$(Daily.TextMatrix(i, 0)) <> "" Then
            Dim temp3 As Recordset, MaxIndex As Long
            Set temp3 = DB.OpenRecordset("Select max(index) as MaxIT from abbrevexercise;", dbOpenDynaset)
            MaxIndex = temp3("Maxit") + 1
            Set temp3 = Nothing
            Set temp3 = DB.OpenRecordset("Select * from abbrevExercise;", dbOpenDynaset)
            temp3.AddNew
            temp3("Index") = MaxIndex
            temp3("Exercisename") = Daily.TextMatrix(i, 0)
            temp3("Formula") = 0
            temp3("Increase") = 1
            temp3("Usage") = 0
            temp3.Update
            temp3.Close
            Set temp3 = Nothing
        
            temp.Fields("ExerciseID") = MaxIndex
        End If
      Else
        temp.Fields("exerciseID") = IDs(i)
      End If
      junk = ""
      For k = 1 To 7
        Daily.Col = k
        If Daily.CellForeColor = vbBlack Then
           junk2 = Daily.Text
        Else
           junk2 = "*" & Daily.Text
        End If
        If Daily.CellBackColor = RGB(50, 50, 50) Then
           junk2 = "-" & junk2
        End If
        'junk = junk & Daily.TextMatrix(i, k) & " ~ "
        junk = junk & junk2 & " ~ "
      Next k
      temp.Fields("weekinfo") = junk
      temp.Update
    End If
  Next i
  On Error Resume Next
  temp.Close
  Set temp = Nothing
  On Error GoTo errhandl
  Dim Day As Long
  Dim firstday As Date
  Dim SDay As Date
  Day = Weekday(SaveDate, vbSunday)
  firstday = DateAdd("d", -1 * Day, SaveDate)
  For i = 1 To 7
      SDay = DateAdd("d", i, firstday)
      Set temp = DB.OpenRecordset("Select * from dailylog where ((date = #" & FixDate(SDay) & "#) and (user = """ & CurrentUser.Username & """));", dbOpenDynaset)
      If temp.RecordCount = 0 Then
        temp.AddNew
        temp.Fields("user") = CurrentUser.Username
      Else
        temp.Edit
      End If
      temp.Fields("date") = SDay ''FixDate(SDay)
      temp.Fields("Exercise_Cal") = Val(Totals(i).Text)
      temp.Update
 Next i
 On Error Resume Next
 temp.Close
 Set temp = Nothing
 Opening = False
 Exit Sub
errhandl:
MsgBox "Unable to save exercises." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub
Private Function GetWeekMaxes(temp As Recordset, AA As Single)
On Error GoTo errhandl
  Dim i As Long, k As Long
  Dim T() As String
  Dim junk As String, Parts() As String
  Dim parts2() As String
  Dim tMax As Single, tMax2 As Single, tMaxSTR As String
  Dim junks As Single, junk2 As String, JunkSTR As String
  Dim temp2 As Recordset, AA2, jj
  i = 0
  While Not temp.EOF
    i = i + 1
    ReDim Preserve T(1, i)
    T(0, i) = temp.Fields("ExerciseID")
    If AA <> 0 Then
       If Left$(Trim$(T(0, i)), 2) = "-1" Then
          jj = 1 '.05
       Else
          Set temp2 = DB.OpenRecordset("Select * from AbbrevExercise where (Index = " & T(0, i) & ");", dbOpenDynaset)
          jj = 1 'temp2.Fields("increase")
       End If
       junk2 = ""
       On Error Resume Next
       junk2 = jj
       AA2 = Val(junk2)
    End If
    junk = temp.Fields("WeekInfo")
      
    Parts = Split(junk, "~", , vbBinaryCompare)
    tMax = 0
    tMax2 = 0
    tMaxSTR = ""
    junk2 = ""
    ReDim parts2(UBound(Parts) - 1)
    For k = 0 To UBound(Parts) - 1
       junk = Trim$(Parts(k))
       If Left$(junk, 1) = "*" Then
          junk = Right$(junk, Len(junk) - 1)
          junks = Val(junk)
          If junks > tMax2 Then
                 tMax2 = junks
             If InStr(1, junk, "/") = 0 Then
                 tMaxSTR = junks
             Else
                 tMaxSTR = junk
             End If
          End If
          parts2(k) = "*"
       ElseIf Left$(junk, 1) = "-" Then
          junk = Right$(junk, Len(junk) - 1)
          junks = Val(junk)
          If junks > tMax2 Then
                 tMax2 = junks
             If InStr(1, junk, "/") = 0 Then
                 tMaxSTR = junks
             Else
                 tMaxSTR = junk
             End If
          End If
          parts2(k) = "-"
       ElseIf junk <> "" Then
          
          junks = Val(junk)
          If junks > tMax2 Then
             tMax2 = junks
             If InStr(1, junk, "/") = 0 Then
                 tMaxSTR = junks
             Else
                 tMaxSTR = junk
             End If
          End If
          parts2(k) = "*"
       End If
    Next k
    If tMax = 0 Then tMax = tMax2
    If AA <> 0 Then
      On Error Resume Next
      tMax = tMax * AA * AA2
      On Error GoTo errhandl
    End If
    For k = 0 To UBound(Parts) - 1
      If parts2(k) <> "" Then
        junk2 = junk2 & parts2(k) & tMaxSTR & " ~ "
      Else
        junk2 = junk2 & parts2(k) & " " & " ~ "
      End If
    Next k
    T(1, i) = junk2
    temp.MoveNext
  Wend
  GetWeekMaxes = T
errhandl:
End Function
Public Sub OpenWeek(OpenDate As Date)
On Error GoTo errOut
Call Daily_LeaveCell
EList.Visible = False
TEnter.Visible = False
  
  TodaysCol = Weekday(OpenDate)
  Dim temp As Recordset, Week As Date, i As Long, k As Long
  Dim temp2 As Recordset
  Dim junk As String, Row As Long, Parts() As String
  Dim tIDs As New Collection, tServings As New Collection
  Dim Datas() As String
  ReDim Datas(1, 0)
  If CurrentUser.Username = "" Then Exit Sub
  Daily.Rows = 1
  Erase IDs
  ReDim IDs(1)
  Erase Formula
  ReDim Formula(1)
  Erase Calories
  ReDim Calories(8, 1)
  Erase Disabled
  ReDim Disabled(8, 1)
  Dim EmptyWeek As Boolean
  
  EList.Text = ""
  EList.Visible = False
  
  TEnter.Text = ""
  TEnter.Visible = False
  
  Week = FindFirstDay(OpenDate)
  Set temp = DB.OpenRecordset("Select * from ExerciseLog where (Week = #" & FixDate(Week) & "#) and (user = """ & CurrentUser.Username & """) ORDER BY ExerciseLog.Order;", dbOpenDynaset)
  Row = 0
  i = 0
  If temp.EOF Then
    temp.Close
    Set temp = Nothing
    On Error GoTo errOut
    Week = DateAdd("ww", -1, Week)
    Set temp = DB.OpenRecordset("Select * from ExerciseLog where (Week = #" & FixDate(Week) & "#) and (user = """ & CurrentUser.Username & """)  ORDER BY ExerciseLog.Order;", dbOpenDynaset)
    If temp.EOF Then
       Set temp = DB.OpenRecordset("Select * from ExercisePlans where (ExercisePlanid = 0) and (user = """ & CurrentUser.Username & """);", dbOpenDynaset)
       Datas = GetWeekMaxes(temp, 0)
       
       EmptyWeek = True
    Else
       If m_UpdateValues Then
         Datas = GetWeekMaxes(temp, 1)
       Else
         Datas = GetWeekMaxes(temp, 0)
       End If
       i = UBound(Datas, 2)
    End If
  Else
    i = 0
    While Not temp.EOF
      i = i + 1
      ReDim Preserve Datas(1, i)
      Datas(0, i) = temp.Fields("ExerciseID")
      Datas(1, i) = temp.Fields("WeekInfo")
      temp.MoveNext
    Wend
  End If
  If i = 0 Then
     ReDim Datas(1, 0)
     EmptyWeek = True
  End If
  
  
  Set temp = Nothing
  Opening = True
  For i = 1 To UBound(Datas, 2)
      Row = Row + 1
      If Row >= Daily.Rows - 1 Then
        Daily.Rows = Daily.Rows + 1
        ReDim Preserve IDs(Daily.Rows)
        ReDim Preserve Formula(Daily.Rows)
        ReDim Preserve Calories(8, Daily.Rows)
        ReDim Preserve Disabled(8, Daily.Rows)
      End If
      junk = Datas(0, i)
      If Left$(junk, 2) = "-1" Then
        junk = Trim$(Right$(junk, Len(junk) - 2))
        Call EList.SetTextAndID(junk, -1111)
      Else
        EList.SelectedID = Val(junk)
      End If
      LRow = Row
      Lcol = 0
      Override = True
      Call GetFormula(EList.Text, EList.SelectedID, Row)
      'Call Daily_LeaveCell
      
      
      junk = Datas(1, i)
      Parts = Split(junk, "~", , vbBinaryCompare)
      Daily.Row = LRow
      For k = 0 To UBound(Parts) - 1
          junk = Trim$(Parts(k))
          Lcol = k + 1
          If Left$(junk, 1) = "*" Then
            junk = Right$(junk, Len(junk) - 1)
            Daily.Col = k + 1
            Daily.CellForeColor = RGB(200, 200, 200)
            Disabled(Lcol, LRow) = True
          End If
          If Left$(junk, 1) = "-" Then
            junk = Right$(junk, Len(junk) - 1)
            Daily.Col = k + 1
            Daily.CellBackColor = RGB(50, 50, 50)
            Disabled(Lcol, LRow) = True
          End If
          Call DoCalories(junk, Row, k + 1)
      Next k
      Override = False
  Next i
  Daily.Col = 0
  
  
errOut:
  Opening = False
  
  Daily.Rows = Daily.Rows + 1
  ReDim Preserve IDs(Daily.Rows)
  ReDim Preserve Formula(Daily.Rows)
  ReDim Preserve Calories(8, Daily.Rows)
  ReDim Preserve Disabled(8, Daily.Rows)

  On Error Resume Next
  If Not temp Is Nothing Then temp.Close
  Set temp = Nothing
  For i = 1 To 7
    DoTotal i
  Next i
  
  Call RedrawTotals
  If EmptyWeek Then Daily.TextMatrix(1, 0) = "Type to enter exercise here"

End Sub

Public Sub AddDataBase(sDB As Database)
On Error Resume Next
    Set mDB = sDB
    EList.AddDataBase sDB
    Set EWeights = DB.OpenRecordset("Select * from abbrevexercise;", dbOpenDynaset)
    SC.AddCode "Weight = " & CurrentUser.Weight
    SC.AddCode "Height = " & CurrentUser.Height '/ 2.54
End Sub

Private Sub Daily_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call Daily_MouseUp(0, 0, CSng(KeyAscii), 0)
End Sub

Private Sub GetFormula(Text As String, ID As Long, Row As Long)
On Error GoTo errhandl
     Dim i As Long
     Daily.TextMatrix(Row, 0) = Text
     IDs(Row) = ID
     If IDs(Row) > 0 Then
         Set EWeights = Nothing
         Set EWeights = DB.OpenRecordset("Select * from abbrevexercise;", dbOpenDynaset)
         EWeights.MoveFirst
         Call EWeights.FindNext("index = " & IDs(Row))
         Formula(Row) = EWeights.Fields("Formula")
     ElseIf IDs(Row) = 0 Then
         Formula(Row) = ""
     ElseIf IDs(Row) = -1111 Then
         Formula(Row) = "0"
     End If
     If Not Opening Then
        For i = 1 To 7
            DoTotal i
        Next i
     End If
errhandl:
End Sub
Private Sub DoCalories(Text As String, Row As Long, Col As Long)
On Error GoTo errhandl
     Dim junk As String, junk2() As String, i As Long
     If TEnter.Visible Or Override Then
        junk = Text
        Daily.TextMatrix(Row, Col) = junk
        If Not Opening Then
             Opening = True
             Daily.Row = Row
             Daily.Col = Col
             If Left$(junk, 1) = "*" Then
                junk = Right$(junk, Len(junk) - 1)
                Daily.CellForeColor = RGB(200, 200, 200)
                Daily.CellBackColor = vbWhite
                Daily.Text = junk
                Disabled(Lcol, LRow) = True
             ElseIf Left$(junk, 1) = "-" Then
                junk = Right$(junk, Len(junk) - 1)
                Daily.CellForeColor = vbBlack
                Daily.CellBackColor = RGB(50, 50, 50)
                Daily.Text = junk
                Disabled(Lcol, LRow) = True
             ElseIf junk <> "" Then
                Daily.CellForeColor = vbBlack
                Daily.CellBackColor = vbWhite
                Disabled(Col, Row) = False
             ElseIf junk = "" Then
                Daily.CellForeColor = vbBlack
                Daily.CellBackColor = vbWhite
                Disabled(Col, Row) = False
                Calories(Col, Row) = 0
             End If
             Opening = False
        End If
        If Trim$(junk) <> "" Then
         '  If Calories(Col, Row) <> 0 Then
             If InStr(1, junk, "/", vbBinaryCompare) <> 0 Then
                junk2 = Split(junk, "/")
                For i = 0 To UBound(junk2)
                 SC.AddCode "Par" & i & " = " & Val(junk2(i))
                Next i
             Else
                SC.AddCode "Par0 = " & Val(junk)
             End If
             On Error Resume Next
             Calories(Col, Row) = SC.Eval(Formula(Row))
          ' End If
        Else
             Calories(Col, Row) = 0
        End If
     End If
errhandl:
End Sub


Private Sub Daily_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
  Call DeleteRow
End If
End Sub

Private Sub Daily_LeaveCell()
On Error GoTo errhandl
Dim i As Long
If Not Opening Then
    If LRow <> -1 Then
        If Lcol = 0 Then
            If EList.Visible Or Override Then
                Call GetFormula(EList.Text, EList.SelectedID, LRow)
                EList.Visible = False
                EList.Text = ""
            End If
        ElseIf TEnter.Visible Or Override Then
           
            Call DoCalories(TEnter.Text, LRow, Lcol)
            TEnter.Visible = False
            TEnter.Text = ""
        End If
        DoTotal Lcol
    End If
End If
errhandl:
End Sub
Private Sub RedrawTotals()
   Dim i As Long, T As Single
   T = Daily.RowPos(Daily.Rows - 1) + 3 * Daily.RowHeight(Daily.Row - 1)
  
  ' T = Daily.Height - Totals(1).Height - 4
   If T + Daily.RowHeight(Daily.Row - 1) > Daily.Height Then
     T = Daily.Height
   End If
   Totals(0).Top = T
   Totals(0).Left = Daily.ColPos(1) - Totals(0).Width + ScaleX(2, vbPixels, UserControl.ScaleMode)
   For i = 1 To 7
    Totals(i).Top = T
    Totals(i).Left = Daily.ColPos(i) + ScaleX(2, vbPixels, UserControl.ScaleMode)
   Next i
End Sub
Private Sub DoTotal(Col As Long)

Dim i As Long, sum As Single
On Error Resume Next
If Col <> 0 Then
    sum = 0
    For i = 0 To UBound(Calories, 2)
      If Not Disabled(Col, i) Then
           If Calories(Col, i) < 0 Then Calories(Col, i) = 0
           sum = sum + Calories(Col, i)
      End If
    Next i
    Totals(Col).Text = Round(sum)
End If

Call RedrawTotals
   
If Col = TodaysCol Then
  RaiseEvent TodaysCalories(Round(sum))
End If
End Sub


Private Sub Daily_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If Button = 2 Then
   Dim i As Long
   For i = 0 To Daily.Rows - 1
     If Y < Daily.RowPos(i) Then
       Y = Daily.RowPos(i - 1)
       Exit For
     End If
   Next i
   LRow = i - 1
   If Daily.Row = Daily.RowSel Then
      Daily.Row = LRow
   End If
   'Daily.Row = LRow
   Call frmMain.PopUpMenu(PopUpMenu, , X, Y)
   Exit Sub
End If
If (Button <> 2) And Not CaloriesMode Then
Dim React As Boolean, LLCol As Long, llRow As Long
If Button = 1 Or Button = 0 Then
  LLCol = Daily.Col
  llRow = Daily.Row
End If

Call Daily_LeaveCell
If Daily.Row <> Daily.RowSel Or Daily.Col <> Daily.ColSel Then
  EList.Visible = False
  TEnter.Visible = False
  Exit Sub
End If
'Todo: this is a dirty little  workaround for having to write this as a stand
'alone procedure, need to clean this up

If Button = 1 Or Button = 0 Then
  Lcol = Daily.Col
  LRow = Daily.Row
End If


If Button = -1 Then
  LRow = Shift
End If

If Button = -2 Then
  Lcol = Shift
  If Lcol > 7 Then
    LRow = LRow + 1
    Lcol = 0
  End If
End If
'If Button = 0 Then
'  LRow = Daily.Row
'  Lcol = Daily.Col
'End If
React = False
If Lcol = 0 Then
    EList.Width = Daily.ColWidth(0)
    EList.Height = Daily.RowHeight(LRow)
    EList.SuggestedHeight = Daily.RowHeight(LRow)
    EList.Left = Daily.ColPos(0) + 50
    EList.Top = Daily.RowPos(LRow) + 50
    Dim junk As String
    junk = Daily.TextMatrix(LRow, Lcol)
    If Trim$(junk) <> "Type to enter exercise here" Then
       EList.Text = junk
    Else
       EList.Text = ""
    End If
    'EList.Reset
    EList.Visible = True
    EList.ZOrder
    If Button = 0 Then EList.Text = Chr$(X)
    EList.SetFocus
    React = True
Else
  If Formula(LRow) <> "" And Daily.TextMatrix(LRow, 0) <> "" Then

    'TEnter.Width = Daily.ColWidth(Lcol)
    'TEnter.Height = Daily.RowHeight(LRow)
    TEnter.Left = Daily.ColPos(Lcol) + 50
    TEnter.Top = Daily.RowPos(LRow) + 50
    Call TEnter.SetLabels(Daily.TextMatrix(LRow, 0))
    TEnter.Text = Daily.TextMatrix(LRow, Lcol)
    'tenter.Reset
    TEnter.Visible = True
    If Button = 0 Then
       
       TEnter.Text = Chr$(X)
       'TEnter.SelStart = 1
       'TEnter.SelLength = 0
    End If
    TEnter.SetFocus
    React = True
  Else
    TEnter.Text = ""
  End If
End If
If React Then
  If LRow >= Daily.Rows - 1 Then
    Daily.Rows = Daily.Rows + 1
    ReDim Preserve IDs(Daily.Rows)
    ReDim Preserve Formula(Daily.Rows)
    ReDim Preserve Calories(8, Daily.Rows)
    ReDim Preserve Disabled(8, Daily.Rows)
  End If
End If
End If
errhandl:
End Sub



Private Sub EList_ItemSelected(SelectedID As Long)
On Error Resume Next
   Call Daily_LeaveCell
   Call Daily_MouseUp(-1, LRow + 1, 0, 0)
   Daily.SetFocus
End Sub

Private Sub EList_NoneSelected()
On Error Resume Next
If EList.Visible Then
  Call Daily_LeaveCell
  Call Daily_MouseUp(-1, LRow + 1, 0, 0)
End If
End Sub


Private Sub TEnter_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If KeyCode = 9 Or KeyCode = 13 Then
      KeyCode = 0
   End If
      'Call TEnter_KeyUp(KeyCode, Shift)
End Sub

Private Sub TEnter_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 9 Or KeyAscii = 13 Then

   KeyAscii = 0
End If
End Sub

Private Sub TEnter_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 9 Or KeyCode = 40 Then
  
  Call Daily_MouseUp(-1, CInt(LRow), 0, 0)
  TEnter.Visible = False
  Daily.SetFocus
  KeyCode = 0
End If
If KeyCode = 38 Then
   Call Daily_MouseUp(-1, LRow - 1, 0, 0)
   Daily.SetFocus
   KeyCode = 0
End If

End Sub

'*************************************************************8

Private Sub UserControl_Initialize()
On Error Resume Next
ReDim IDs(1)
ReDim Formula(1)
ReDim Calories(8, 1)
ReDim Disabled(8, 1)

LRow = -1
Dim Top(7) As String
Top(0) = "Exercise"
Top(1) = "Sunday"
Top(2) = "Mon"
Top(3) = "Tues"
Top(4) = "Wed"
Top(5) = "Thurs"
Top(6) = "Fri"
Top(7) = "Sat"

Dim i As Long, j As Long, FG As MSFlexGrid
For i = 0 To 2
   For j = 0 To 7
     Daily.TextArray(j) = Top(j)
   Next j
Next i
Daily.ColWidth(0) = UserControl.TextWidth(Space(60)) * 2
For i = 1 To 7
  If i > 1 Then Load Totals(i)
  Totals(i).Visible = True
  Totals(i).ZOrder
  Totals(i).Left = Daily.ColPos(i) + 50
  Totals(i).Width = Daily.ColWidth(i)
  Totals(i).Text = 0
Next i

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
   Daily.Move 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - Totals(1).Height
   Dim i As Long, T As Single
   Call RedrawTotals
   Text1.Left = (UserControl.ScaleWidth - Text1.Width) / 2
   Text1.Top = UserControl.ScaleHeight - Text1.Height
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next
    Daily.BackColorBkg = PropBag.ReadProperty("BackColor", 8421504)
    m_AddLastWeeks = PropBag.ReadProperty("AddLastWeeks", m_def_AddLastWeeks)
    m_UpdateValues = PropBag.ReadProperty("UpdateValues", m_def_UpdateValues)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("BackColor", Daily.BackColorBkg, 8421504)
    Call PropBag.WriteProperty("AddLastWeeks", m_AddLastWeeks, m_def_AddLastWeeks)
    Call PropBag.WriteProperty("UpdateValues", m_UpdateValues, m_def_UpdateValues)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Daily,Daily,-1,BackColorBkg
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
On Error Resume Next
    BackColor = Daily.BackColorBkg
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    Daily.BackColorBkg() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DoFormat() As Variant
On Error Resume Next
   Call SaveWeek(Curdate)
   Call OpenWeek(Curdate)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Recalculate() As Variant
On Error Resume Next
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get AddLastWeeks() As Boolean
On Error Resume Next
    AddLastWeeks = m_AddLastWeeks
End Property

Public Property Let AddLastWeeks(ByVal New_AddLastWeeks As Boolean)
On Error Resume Next
    m_AddLastWeeks = New_AddLastWeeks
    PropertyChanged "AddLastWeeks"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get UpdateValues() As Boolean
On Error Resume Next
    UpdateValues = m_UpdateValues
End Property

Public Property Let UpdateValues(ByVal New_UpdateValues As Boolean)
On Error Resume Next
    m_UpdateValues = New_UpdateValues
    PropertyChanged "UpdateValues"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    m_AddLastWeeks = m_def_AddLastWeeks
    m_UpdateValues = m_def_UpdateValues
End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

        Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function

Public Property Get Font() As Font
On Error Resume Next
    Set Font = Daily.Font
  
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
   Set Daily.Font = Font
   Set TEnter.Font = Font
   PropertyChanged "Font"
End Property
