VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl AdvancedFlex 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   ScaleHeight     =   7215
   ScaleWidth      =   8820
   Begin CalorieBalance.UnitListBox Units 
      Height          =   2175
      Left            =   5040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   End
   Begin CalorieBalance.AutoComplete AC 
      Height          =   3375
      Left            =   1320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5953
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
   Begin CalorieBalance.PieChart Balance 
      Height          =   1920
      Left            =   5880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4905
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3387
      MaskPicture     =   "Flex.ctx":0000
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
   Begin VB.PictureBox FlagBox 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TServing 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5220
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   4575
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -60
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   -2147483626
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
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
   Begin CalorieBalance.ProgressBars PB 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3413
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
   Begin CalorieBalance.ProgressBars PB2 
      Height          =   1935
      Left            =   2760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4725
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3413
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
   Begin VB.Label TotalLabel 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "AdvancedFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public ReadOut As Single

Private ACCol As Long
Private UnitCol As Long
Private ServingCol As Long
Private FlagCol As Long
Private ReservedCol  As Long
Private LastObject As Object
Private LastRow As Long
Private LastCol As Long
'Default Property Values:
Const m_def_ShowTotals = True
Const m_def_Text = ""
Const PromptPhrase = "Type food name here..."

Const m_def_ShowAsPercent = True
Const m_def_Changed = False
Const m_def_SaveDay = 0

Event RowUpdated()
Event PresentPopup(PopUp As Object, Un, X As Single, Y As Single, MealRow As Boolean, MealName As String)
Event MakeOfficalMeal(MealRowNumber As Long, MealTime As Long)
Event InternetSearch(SearchText As String)

'Property Variables:
Dim m_YFixed As Boolean
Dim m_ShowTotals As Boolean
Dim m_Text As String

Dim m_ShowAsPercent As Boolean
Dim m_Changed As Boolean
Dim m_SaveDay As Variant
Dim mDisplayDate As Date


Dim mDB As Database
Dim Maxes As Calories

Dim Headers() As String
Dim Username As String
Dim SelectedIDS() As Long
Dim GramsSelected() As Single
Dim ServingsSelected() As Single
Dim ExtraSelected()
Dim MealIDs() As Long
Dim Stuff() As Single
Dim Totals() As Single, TMacro(4) As Single
Dim Maxs() As Single
Dim fat() As Single, carbs() As Single, Prot() As Single, fiber() As Single, sugar() As Single
Dim PopUp As Object
Dim m_Calories As Single
Dim ExerciseCals As Single
Dim YTitles
Dim NoLeave As Boolean


Public OrderMeals As Boolean
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage


Public Sub EndIt()
On Error Resume Next
       Call SaveSetting(App.Title, "Settings", "ACWidth", FG.ColWidth(ACCol))
       Call SaveSetting(App.Title, "Settings", "UnitWidth", FG.ColWidth(UnitCol))

End Sub

Public Sub SetBackGround(back)
On Error Resume Next
   UserControl.BackColor = back
   Balance.BackColor = back
End Sub
Public Function PrintDay(Filename As String) As Variant
On Error GoTo errhandl
   Dim ff As Long
   Dim i As Long, junk As String, thing As String
   Dim temp As Recordset, MLs As Recordset
   Dim Outlines() As String
    
    
  Dim MealPlans As Recordset
  Dim ID As Long, Serving As String, Unit As String
  'open the days foods and load them into the data fields
  Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.date)=#" & FixDate(DisplayDate) & "#) AND (DaysInfo.user='" & Username & "')) ORDER BY daysinfo.order,itemid;", dbOpenDynaset)
  
  If temp.EOF Then
     ReDim Outlines(3, 1)
  End If
  'get the date
    i = 1
    Do While Not temp.EOF
      Unit = ""
      Serving = ""
      thing = ""
      ID = 0
      ID = temp.Fields("itemID")
      
      Set MLs = DB.OpenRecordset("select * from abbrev where index=" & ID & ";", dbOpenDynaset)
      thing = MLs("foodname")
      On Error Resume Next
      
      
      If ID <> 0 And ID <> -1111 Then
         Unit = temp.Fields("unit")
         Serving = temp.Fields("Servings")
      End If
      ReDim Preserve Outlines(3, i)
      If ID <= -200 And ID <> -555 Then
         ID = Abs(-200 - ID)
         Set MLs = DB.OpenRecordset("select * from meals " _
           & "where user='" & CurrentUser.Username & "' and " _
           & "mealnumber = " & ID & " and " _
           & "entrydate=#" & FixDate(mDisplayDate) & "#;", dbOpenDynaset)
         Serving = Unit
         Serving = ""
         
         If MLs.EOF Then
           Unit = FG.TextMatrix(i - 1, UnitCol) = "Make into Meal"
         End If
         MLs.Close
         Set MLs = Nothing
      Else
      
      
      End If
      temp.MoveNext
      
      Outlines(0, i) = Serving
      Outlines(1, i) = Unit
      Outlines(2, i) = thing
      i = i + 1
    Loop
    
   ff = FreeFile
   Open Filename For Output As #ff
   Print #ff, "<html><body><h1>Days foods for " & CurrentUser.Username & "</h1>"
   Print #ff, "<table width =""100%"">"
   Print #ff, "<tr><td><b>Amount</b></td>"
   Print #ff, "<td><b>Foodname</b></td>"
   Print #ff, "<td><b>Calories</b></td></tr>"
   
   For i = 1 To FG.Rows - 1
        Print #ff, "<tr >"
        If Trim$(Outlines(0, i)) = "" Then
         If (i <= UBound(Outlines, 2)) Then
          If (Trim$(Outlines(2, i)) <> "") Then
            Print #ff, "   <td style=""border-bottom-style: solid; ""> &nbsp;&nbsp;&nbsp;&nbsp;</td>"
            Print #ff, "   <td style=""border-bottom-style: solid; "">" & Outlines(2, i) & "</td>"
            Print #ff, "   <td style=""border-bottom-style: solid; "">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
          End If
          End If
        Else
            Print #ff, "   <td >" & Outlines(0, i) & " " & Outlines(1, i) & "</td>"
            Print #ff, "   <td >" & Outlines(2, i) & "</td>"
            Print #ff, "   <td >" & Round(Stuff(0, i), 1) & "</td>"
        End If
        Print #ff, "</tr>"
   Next i
   Print #ff, "</table><table>"
   Print #ff, "<tr>"
   For i = 0 To UBound(Headers)
     Print #ff, "<td><b>" & Headers(i) & "</b></td>"
      
   Next i
   Print #ff, "</tr><tr>"
   For i = 0 To UBound(Headers)
     Print #ff, "<td><b>" & Round(Totals(i)) & "</b></td>"
   Next i
   Print #ff, "</tr></table>   </body></html>"
   Close #ff
   Exit Function
errhandl:
Close #ff
MsgBox "Unable to print." & vbCrLf & Err.Description, vbOKOnly, ""
'Resume 'todo remove this
End Function
Public Sub DeleteMeal()
On Error GoTo errhandl
  Dim i As Long, j As Long, RS As Long, re As Long
  Call LeaveCell
  
  AC.Visible = False
  Units.Visible = False
  TServing.Visible = False
  
 ' Call SaveDay(DisplayDate)
  RS = FG.Row
  Dim r As Recordset
  Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" _
    & CurrentUser.Username & "' and daysinfo.order=" & RS & ";", dbOpenDynaset)
  Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" _
    & CurrentUser.Username & "' and meal=" & r("meal") & ";", dbOpenDynaset)
  Dim MealNumber As Long
  MealNumber = r("meal")
  While Not r.EOF
    r.Delete
    r.MoveNext
  Wend
  r.Close
  Set r = Nothing
  
  
  Dim ms As Recordset, ms2 As Recordset
  Set ms = DB.OpenRecordset("select * from meals where meals.user='" & CurrentUser.Username _
     & "' and entrydate=#" & FixDate(mDisplayDate) & "# and mealnumber=" & MealNumber & ";", dbOpenDynaset)
  
  If Not ms.EOF Then
  
     ms.Delete
     If Not OrderMeals Then
        Set ms = DB.OpenRecordset("select * from meals where meals.user='" & CurrentUser.Username _
            & "' and entrydate=#" & FixDate(mDisplayDate) & "# order by mealnumber;", dbOpenDynaset)
        MealNumber = 0
        While Not ms.EOF
           ms.Edit
           ms("mealnumber") = MealNumber
           MealNumber = MealNumber + 1
           ms.Update
           ms.MoveNext
        Wend
     End If
  End If
  Set ms = Nothing
  Set ms2 = Nothing
  
errhandl:
  On Error Resume Next
  Call OpenDay(mDisplayDate)

End Sub
Private Sub SyncMealPlanner()
On Error GoTo errhandl
  Dim ms As Recordset, ms2 As Recordset
  
  Set ms = DB.OpenRecordset("SELECT * from meals " _
      & "where meals.user='" & CurrentUser.Username & _
      "' and entrydate=#" & FixDate(mDisplayDate) & "#;", dbOpenDynaset)
  While Not ms.EOF
    Set ms2 = DB.OpenRecordset("select * from daysinfo where user ='" & CurrentUser.Username _
       & "' and date=#" & FixDate(mDisplayDate) & "# and mealid=" & ms("id") & ";", dbOpenDynaset)
    If ms2.EOF And ms2.BOF Then
      ms.Delete
    End If
    ms.MoveNext
  Wend
  Set ms = Nothing
  Set ms2 = Nothing
  
  Dim MealNumber As Long
  'Exit Sub
    If Not OrderMeals Then
        Set ms = DB.OpenRecordset("select * from meals where meals.user='" & CurrentUser.Username _
            & "' and entrydate=#" & FixDate(mDisplayDate) & "# order by mealnumber;", dbOpenDynaset)
        MealNumber = 0
        While Not ms.EOF
           ms.Edit
           ms("mealnumber") = MealNumber
           MealNumber = MealNumber + 1
           ms.Update
           ms.MoveNext
        Wend
     End If
errhandl:
End Sub
Public Sub DeleteRows(Optional nRows As Long = 1)
On Error GoTo errhandl
  Dim i As Long, j As Long, RS As Long, re As Long
  Call LeaveCell
  
  AC.Visible = False
  Units.Visible = False
  TServing.Visible = False
  
  re = FG.RowSel
  RS = FG.Row
  If re < RS Then
     i = re
     re = RS
     RS = i
  End If
  Dim r As Recordset
  Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" _
    & CurrentUser.Username & "';", dbOpenDynaset)
  While Not r.EOF
    If r("order") >= RS And r("order") <= re Then
       r.Delete
    End If
    r.MoveNext
  Wend
    
  r.Close
  Set r = Nothing
  On Error Resume Next
  Call SyncMealPlanner
errhandl:
 
  Call OpenDay(mDisplayDate)
End Sub

Public Function SelectRow(Y As Single) As Long
Dim cy As Single, i As Long
For i = 0 To FG.Rows - 1
  cy = cy + FG.RowHeight(i)
  If cy > Y Then
    FG.RowSel = i
    FG.Row = i
    SelectRow = i
    Exit For
  End If
Next i
End Function
Public Sub InsertRows()
On Error GoTo errhandl
  Dim i As Long, j As Long, RS As Long, re As Long
  
  Call LeaveCell
  AC.Visible = False
  Units.Visible = False
  TServing.Visible = False

  re = FG.RowSel
  RS = FG.Row
  If re < RS Then
     i = re
     re = RS
     RS = i
  End If
  
  Dim r As Recordset
  Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" _
    & CurrentUser.Username & "';", dbOpenDynaset)
    
  While Not r.EOF
    If r("order") >= RS Then
       r.Edit
       r("order") = r("order") + (re - RS) + 1
       r.Update
    End If
    r.MoveNext
  Wend
  For i = RS To re
     r.AddNew
     r("user") = CurrentUser.Username
     r("date") = FixDate(mDisplayDate)
     r("itemID") = -555
     r("order") = i
     r.Update
  Next i
    
  r.Close
  Set r = Nothing
  
' Exit Sub
errhandl:
On Error Resume Next
  Call OpenDay(mDisplayDate)

End Sub
Public Sub Paste()

On Error GoTo errhandl
 Dim lines() As String, Parts() As String, junk As String
 Dim RS As Long, n As Long, i As Long, j As Long, k As Long
 Dim r As Recordset
 
 'clear off the entry stuff
 Call LeaveCell
 AC.Visible = False
 Units.Visible = False
 TServing.Visible = False
 'get the new information
 junk = Clipboard.GetText
 'save what is on the grid

 'get the lines.  user may paste the information in from another program, so
 'we have to check each line for data
 If InStr(1, junk, vbCrLf, vbBinaryCompare) <> 0 Then
    lines = Split(junk, vbCrLf, , vbBinaryCompare)
 Else
    ReDim lines(0)
    lines(0) = junk
 End If
 RS = FG.RowSel
 'move all the data up
 Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" & CurrentUser.Username & "' and [order]>=" & RS & ";", dbOpenDynaset)
 While Not r.EOF
    r.Edit
    r("order") = r("order") + UBound(lines)
    r.Update
    r.MoveNext
 Wend
 r.Close
 Set r = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) & "# and user='" & CurrentUser.Username & "' and [order]<" & RS & " order by [order];", dbOpenDynaset)
 Dim ID As Long, MealN As Long, LunchPassed As Boolean
 While Not r.EOF
    ID = r("itemid")
    If ID <= -200 Then
       If ID = -200 Then MealN = 0
       If ID = -201 And Not LunchPassed Then MealN = 1
       If ID = -202 Then
          LunchPassed = True
          MealN = 2
       End If
       If ID = -201 And LunchPassed Then MealN = 3
       If ID = -203 Then MealN = 4
       If ID = -204 Then MealN = 5
    End If
    r.MoveNext
 Wend
 'now enter the new data
 Dim A As Recordset
 For i = 0 To UBound(lines)
    Parts = Split(lines(i), vbTab)
    If LCase$(Parts(0)) <> "foodname" Then
       Set A = DB.OpenRecordset("select * from abbrev where foodname='" & Parts(0) & "';", dbOpenDynaset)
       If ID <= -200 Then
          If ID = -200 Then MealN = 0
          If ID = -201 And Not LunchPassed Then MealN = 1
          If ID = -202 Then
             LunchPassed = True
             MealN = 2
          End If
          If ID = -201 And LunchPassed Then MealN = 3
          If ID = -203 Then MealN = 4
          If ID = -204 Then MealN = 5
       End If
       
       If A.EOF = False Then
          r.AddNew
          r("date") = FixDate(mDisplayDate)
          r("user") = CurrentUser.Username
          r("itemid") = A("index")
          If UBound(Parts) > 1 Then
             r("unit") = Parts(1)
             r("servings") = Val(Parts(2))
          End If
          r("order") = i + RS
          r("meal") = MealN
          r.Update
       End If
       
       A.Close
       Set A = Nothing
    End If
   
 Next i
 Call OpenDay(mDisplayDate)
 Exit Sub
errhandl:
MsgBox Err.Description, vbOKOnly, ""

End Sub
Public Sub Copy()
On Error GoTo errhandl
  Dim junk As String, i As Long, j As Long
  Call LeaveCell
  AC.Visible = False
  Units.Visible = False
  TServing.Visible = False
  
  junk = ""
  Dim RS As Long, re As Long
  RS = FG.Row
  re = FG.RowSel
  If RS > re Then
    i = RS
    RS = re
    re = i
  End If
  'now get all the data
  For j = 0 To FG.Cols - 1
     junk = junk & FG.TextArray(j) & vbTab
  Next j
  junk = junk & vbCrLf
  For i = RS To re
    For j = 0 To FG.Cols - 1
      junk = junk & FG.TextMatrix(i, j) & vbTab
    Next j
    If i <> FG.RowSel Then junk = junk & vbCrLf
  Next i
  'put it all on the clipboard
  Clipboard.Clear
  Clipboard.SetText junk
  'FG.RowSel = FG.Row
  'FG.ColSel = FG.Col
  
  Exit Sub
errhandl:
MsgBox Err.Description, vbOKOnly, ""

End Sub



Public Sub SetPopUpMenu(newMenu As Object)
On Error Resume Next
  Set PopUp = newMenu
End Sub

Public Function GetNutrients(ID As Long, Table As String, Field As String, Optional Names As Collection) As Collection
On Error GoTo errhandl
Dim sTemp As Recordset
Dim temp As Collection
Dim junk As String, JunkName As String
    Set sTemp = mDB.OpenRecordset("SELECT * FROM " & Table & " WHERE (" & Field & " = " & _
                                  ID & ");", dbOpenDynaset)
     If Err = 3075 Then
    'Here we got a bug!!
        Exit Function
      '
     End If
    
     If Not sTemp.RecordCount = 0 Then
          Set temp = New Collection
          Set Names = New Collection
          sTemp.MoveFirst
          Dim i As Long
          On Error Resume Next
          For i = 0 To sTemp.Fields.Count - 1
            JunkName = sTemp.Fields(i).Name
            junk = sTemp.Fields(i)
            If Err.Number <> 0 Then
              junk = ""
              Err.Clear
            End If
            temp.Add junk, JunkName
            Names.Add JunkName, JunkName
          
          Next i
     End If
     Set GetNutrients = temp
     
    sTemp.Close
    Set sTemp = Nothing
errhandl:
    
End Function
Public Function SetRow(Row As Long, ID As Long, nServing As Single, nUnit As String, ExtraText As String, Optional ShowTotals As Boolean = True) As Boolean
On Error GoTo errhandl
  Dim r As Long
  r = Row
  If FG.Rows - 1 <= Row Or Row > UBound(SelectedIDS) Then
    FG.Rows = r + 4
    Call MacroNut(FG.Rows - 1)
  End If
      
      
  Dim temp As Recordset
  Dim junk As String, i As Long
  junk = AC.Translate(ID)
  SelectedIDS(r) = ID
  FG.TextMatrix(r, ACCol) = junk
  If junk <> "" And ID > -200 Then
        On Error Resume Next
        GramsSelected(r) = Units.Translate(ID, nUnit)
        If Err.Number <> 0 Then GramsSelected(r) = 0
        FG.TextMatrix(r, UnitCol) = nUnit
        'prevent serving from showing a zero if there is no data
        SetRow = True
  ElseIf ID <= -200 Then
       SetRow = True
  Else
       SetRow = False
  End If
  If ID <> 0 And nUnit <> "" And SetRow Then
    ServingsSelected(r) = nServing
    FG.TextMatrix(r, ServingCol) = ConvertDecimalToFraction(nServing)
    FG.TextMatrix(r, FlagCol) = ExtraText
  ElseIf ID > -200 Then
    For i = 0 To FG.Cols - 1
      FG.TextMatrix(r, i) = ""
    Next i
    SetRow = False
  End If
  If ID <= -200 And nUnit <> "" Then
    FG.TextMatrix(r, UnitCol) = nUnit
    FG.TextMatrix(r, ServingCol) = ""
  End If
  
  If SetRow Then Call UpdateRow(r, Not ShowTotals)
  Exit Function
errhandl:
 ' Resume 'todo remove this
End Function
Public Function GetRow(Row As Long, ID As Long, Serving As Single, Unit As String, ExtraText As String, Optional Grams As Single, Optional Foodname As String) As Variant
On Error Resume Next
ID = SelectedIDS(Row)
Serving = ServingsSelected(Row)
Unit = FG.TextMatrix(Row, UnitCol)
Grams = GramsSelected(Row)
ExtraText = ExtraSelected(Row)
Foodname = FG.TextMatrix(Row, ACCol)
If Row > UBound(SelectedIDS) Then ID = -1
End Function
Private Sub UpdateRow(Row As Long, Optional SkipTotals As Boolean = False)
On Error Resume Next
   Dim temp As Collection, i As Long, X As Single, j As Long, junk As String
   Dim cc As Single
   Dim Calories As Single
   Dim Trow As Long, Tcol As Long
   Dim RS As Recordset
   On Error GoTo errhandl
   
   Trow = FG.Row
   Tcol = FG.Col
   
   If SelectedIDS(Row) = 0 Or (ServingsSelected(Row) = 0 And SelectedIDS(Row) > -200) Or FG.TextMatrix(Row, UnitCol) = "" Then
      Set RS = DB.OpenRecordset("select * from daysinfo where (daysinfo.order=" & Row _
         & " and daysinfo.date=#" & FixDate(mDisplayDate) & "#" _
         & " and user='" & CurrentUser.Username & "');", dbOpenDynaset)
      If Not RS.EOF Then
         RS.Edit
         RS("itemid") = 0
         RS.Update
      End If
   End If
   If SelectedIDS(Row) = 0 Then
   
       Exit Sub
   Else
      Set RS = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) _
         & "# and user='" & CurrentUser.Username & "' and daysinfo.order<" & Row & " order by daysinfo.order;", dbOpenDynaset)
      Dim LMeal As Long
      If SelectedIDS(Row) < -199 Then
         LMeal = Abs(-200 - SelectedIDS(Row))
      End If
      While Not RS.EOF
           If RS("itemid") < -199 Then
             i = Abs(-200 - RS("itemid"))
             If i > LMeal Then LMeal = i
           ElseIf RS("meal") >= 0 And RS("meal") > LMeal Then
             LMeal = RS("meal")
           End If
           RS.MoveNext
      Wend

      Set RS = DB.OpenRecordset("select * from daysinfo where date=#" & FixDate(mDisplayDate) _
         & "# and user='" & CurrentUser.Username & "' and daysinfo.order=" & Row & ";", dbOpenDynaset)
      
      If Not RS.EOF Then
         RS.Edit
      Else
         RS.AddNew
      End If
      junk = FG.TextMatrix(Row, UnitCol)
      If Not (RS("itemid") = SelectedIDS(Row) And RS("unit") = junk And RS("servings") = ServingsSelected(Row)) Then
      
          RS("date") = mDisplayDate 'FixDate(mDisplayDate)
          RS("user") = CurrentUser.Username
          RS("itemid") = SelectedIDS(Row)
          RS("unit") = FG.TextMatrix(Row, UnitCol)
          RS("servings") = ServingsSelected(Row)
            RS("order") = Row
            RS("meal") = LMeal
            RS("mealid") = -1
            RS.Update
      End If
      Set RS = Nothing
      
   End If
   
   
   
   NoLeave = True
   If SelectedIDS(Row) <= -200 Then
      FG.Row = Row
      FG.Col = ACCol
      FG.CellFontBold = True
      FG.CellBackColor = RGB(225, 225, 225)
      FG.Col = UnitCol
      FG.CellBackColor = RGB(225, 225, 225)
      If LCase$(FG.Text) = "make into meal" Then
          FG.CellForeColor = vbBlue
          FG.TextMatrix(Row, UnitCol) = "Make Into Meal"
      End If
      For i = ServingCol To FG.Cols - 1
         FG.Col = i
         FG.Text = ""
         FG.CellBackColor = RGB(225, 225, 225)
      Next i
      fat(Row) = 0
      carbs(Row) = 0
      Prot(Row) = 0
      fiber(Row) = 0
      sugar(Row) = 0
      GoTo DTot
   End If
   
   'clear off any formating from row
   FG.Row = Row
   For i = 1 To FG.Cols - 1
      FG.Col = i
      FG.CellFontBold = False
      FG.CellBackColor = vbWhite
      FG.CellForeColor = 0
   Next i
   FG.Col = ACCol
   FG.CellFontBold = False
   FG.CellBackColor = vbWhite
   
   '*********************************8
   
   Set temp = GetNutrients(SelectedIDS(Row), "Abbrev", "Index")
   If temp Is Nothing Then
      NoLeave = False
      Exit Sub
   End If
   cc = GramsSelected(Row) / 100 * ServingsSelected(Row)
   
   'now read out the totals
   
   'On Error GoTo skipto
   For i = 0 To UBound(Headers)
      If Headers(i) = "Carbohydrate" Then
        junk = temp("Carbs")
      ElseIf Headers(i) = "Calories Net" Then
        junk = temp("calories")
      Else
        junk = temp(Headers(i))
      End If
      If Headers(i) = "kj" Then
        junk = temp("calories") * 4.1868
      End If
      X = Val(junk)
      X = X * cc
      If m_ShowAsPercent Then
        If Maxs(i) = 0 Then
          FG.TextMatrix(Row, ReservedCol + i) = 0
        Else
          FG.TextMatrix(Row, ReservedCol + i) = Round(X / Maxs(i) * 100, 1) & "%"
        End If
      Else
         FG.TextMatrix(Row, ReservedCol + i) = Round(X, 1)
      End If
      Stuff(i, Row) = X
   Next i
   
'skipto:
   On Error Resume Next
   fat(Row) = temp("Fat") * cc
   carbs(Row) = temp("Carbs") * cc
   Prot(Row) = temp("Protein") * cc
   fiber(Row) = temp("Fiber") * cc
   sugar(Row) = temp("Sugar") * cc
   
   FG.Row = Trow
   FG.Col = Tcol
   
   Changed = True
   If SkipTotals Then
     NoLeave = False
     Exit Sub
   End If
DTot:
   On Error GoTo errhandl
   For i = 0 To UBound(Headers)
      Totals(i) = 0
      For j = 0 To UBound(Stuff, 2)
         Totals(i) = Totals(i) + Stuff(i, j)
      Next j
   Next i
   m_Calories = Totals(0)
   Dim f As Single, c As Single, p As Single, fb As Single, s As Single
   For j = 0 To UBound(Stuff, 2)
      f = f + fat(j)
      c = c + carbs(j)
      p = p + Prot(j)
      fb = fb + fiber(j)
      s = s + sugar(j)
   Next j
   On Error Resume Next
   Dim Pcnt As Single
   If m_ShowTotals Then
   For i = 0 To UBound(Headers)
      If i = 0 Then
         Pcnt = Round((Totals(i)) / Maxs(i) * 100)
        
         If i > 5 Then
           Call PB2.UpdateLine(i - 6, Pcnt, Round(Totals(i)))
         Else
           Call PB.UpdateLine(i, Pcnt, Round(Totals(i)))
         End If
      ElseIf LCase$(Headers(i)) = "calories net" Then
         Pcnt = Round((m_Calories - ExerciseCals) / Maxs(0) * 100)
        
        ' If i > 5 Then
        '   Call PB2.UpdateLine(i - 6, Pcnt, Round(m_Calories - ExerciseCals))
        ' Else
           Call PB.UpdateLine(i, Pcnt, Round(m_Calories - ExerciseCals))
        ' End If
      
      Else
         Pcnt = Round(Totals(i) / Maxs(i) * 100)
         If Headers(i) = "Fiber" Or Headers(i) = "Protein" Then
            If Pcnt > 100 Then Pcnt = 100
         End If
         If i > 5 Then
            Call PB2.UpdateLine(i - 6, Pcnt, Round(Totals(i)))
         Else
            Call PB.UpdateLine(i, Pcnt, Round(Totals(i)))
         End If
      End If
   Next i
   
   PB.Draw
   PB2.Draw
   
   Call FigurePercentages(Balance, Totals(0), f, s, c, p, fb)
   TMacro(0) = f
   TMacro(1) = s
   TMacro(2) = c
   TMacro(3) = p
   TMacro(4) = fb
   
   End If
  
   NoLeave = False
   RaiseEvent RowUpdated
   Exit Sub
errhandl:
   
   NoLeave = False
   'Resume ' todo
End Sub
Private Sub MacroNut(Rows As Long)
  On Error Resume Next
  ReDim Preserve SelectedIDS(Rows)
  ReDim Preserve GramsSelected(Rows)
  ReDim Preserve ServingsSelected(Rows)
  ReDim Preserve Stuff(UBound(Headers), Rows)
  ReDim Preserve ExtraSelected(Rows)
  ReDim Preserve fat(Rows)
  ReDim Preserve fiber(Rows)
  ReDim Preserve Prot(Rows)
  ReDim Preserve carbs(Rows)
  ReDim Preserve sugar(Rows)
  ReDim Preserve MealIDs(Rows)
  
End Sub
Public Sub SetHeads(Heads() As String)
  On Error Resume Next
  Dim i As Long
  ReDim Headers(UBound(Heads))
  Erase Stuff
  ReDim Stuff(UBound(Heads), FG.Rows - 1)
  ReDim Totals(UBound(Heads))
  ReDim Maxs(UBound(Heads))
  
  FG.TextArray(ACCol) = "FoodName"
  FG.TextArray(UnitCol) = "Unit"
  FG.TextArray(ServingCol) = "Serving"
  FG.TextArray(FlagCol) = "Flags"
  
  FG.Cols = ReservedCol + UBound(Heads) + 1
  PB.Clear
  On Error Resume Next
  For i = 0 To 6 'UBound(Heads)
    Maxs(i) = Maxes(Heads(i))
    Headers(i) = Heads(i)
    FG.TextArray(i + ReservedCol) = Heads(i)
    PB.AddLine 0, Heads(i), 0
  Next i
  PB.Draw
  PB2.Clear
  For i = 7 To UBound(Heads)
    Maxs(i) = Maxes(Heads(i))
    Headers(i) = Heads(i)
    FG.TextArray(i + ReservedCol) = Heads(i)
    PB2.AddLine 0, Heads(i), 0
  Next i
  
  PB2.Draw
  Call UserControl_Resize
End Sub

Private Sub LocateObject(c As Long, r As Long)
  On Error GoTo errhandl
   Dim T2 As Long
   T2 = r - (Int(FG.Height / FG.RowHeight(1)) + 2)
   If T2 > 0 Then FG.TopRow = r
   If TypeOf LastObject Is UnitListBox Then
     LastObject.Move FG.ColPos(c) + 50, FG.RowPos(r) + 50, FG.ColWidth(c)
   ElseIf TypeOf LastObject Is TextBox Then
     LastObject.Move FG.ColPos(c) + 50, FG.RowPos(r) + 50, FG.ColWidth(c), FG.RowHeight(r)
   Else
     LastObject.Move FG.ColPos(c) + 50, FG.RowPos(r) + 50, FG.ColWidth(c), FG.RowHeight(r)
   End If
   LastObject.Visible = True
   LastObject.TabStop = False
   LastObject.SetFocus
errhandl:
End Sub

Private Sub LeaveCell()
 
  Dim change As Boolean, junk As String
  On Error GoTo exitS
  change = False
  If LastObject.Visible Then
    junk = LastObject.Text
    If FG.TextMatrix(LastRow, LastCol) <> junk Then
       FG.TextMatrix(LastRow, LastCol) = junk
       change = True
    End If
    LastObject.Visible = False
  End If
  'only update the display if we see a change in the value
  If change Then
      If SelectedIDS(LastRow) <> 0 And GramsSelected(LastRow) <> 0 And ServingsSelected(LastRow) <> 0 Then
         Call UpdateRow(LastRow)
      End If
      If SelectedIDS(LastRow) <= -200 Then
         Call UpdateRow(LastRow)
      End If
  End If
  Exit Sub
exitS:
  LastObject.Visible = False
End Sub
Private Sub MoveCell(Col As Long, Row As Long)

On Error GoTo errhandl
If Row = 0 Then Exit Sub

FG.Col = Col
FG.Row = Row


If Row + 1 >= FG.Rows Or Row > UBound(SelectedIDS) Then
  FG.Rows = FG.Rows + 4
  Call MacroNut(FG.Rows - 1)
End If
LastRow = Row
LastCol = Col
If Col = ACCol Then
   Set LastObject = AC
   Call LocateObject(Col, Row)
   LastObject.SuggestedHeight = FG.RowHeight(Row)
   If FG.Text = PromptPhrase Then FG.Text = ""
   
   Call AC.SetTextAndID(FG.Text, SelectedIDS(Row))
   AC.ZOrder
ElseIf Col = UnitCol Then
   If SelectedIDS(Row) <> 0 And SelectedIDS(Row) > -200 Then
       Call Units.LoadListBox(SelectedIDS(Row))
       Set LastObject = Units
       Call LocateObject(Col, Row)
       If FG.Text = "" Then
         GramsSelected(Row) = Units.GramsSelected
         Units.Text = ""
       Else
         Units.Text = FG.Text
         GramsSelected(Row) = Units.GramsSelected
       End If
   ElseIf LCase$(FG.TextMatrix(Row, Col)) = "make into meal" Then
       RaiseEvent MakeOfficalMeal(Row, Abs(-200 - SelectedIDS(Row)))
       'FG.TextMatrix(Row, Col) = ""
   End If
ElseIf Col = ServingCol Then
  If SelectedIDS(Row) > -200 Then
   Set LastObject = TServing
   Call LocateObject(Col, Row)
  
   TServing.Text = FG.Text
   TServing.SelStart = Len(FG.Text)
  End If
ElseIf Col = FlagCol Then
  ' Set LastObject = FlagBox
  'Call LocateObject(Col, Row)
  'flagbox.text = fg.text
End If

errhandl:
End Sub

Public Sub ForceLoseFocus()
On Error Resume Next
  If AC.Visible Then Call AC.CloseBox
  If Units.Visible Then Call Units.UserControl_ExitFocus
  If TServing.Visible Then Call TServing_KeyPress(13)
End Sub
Private Sub AC_ExitFocus()
On Error Resume Next
Call AC.CloseBox
End Sub

Private Sub UpdateAC(SelectedID As Long)
On Error Resume Next
    Dim j As Long
    Dim Tcol As Long, Trow As Long
    Trow = FG.Row
    Tcol = FG.Col
    FG.TextMatrix(LastRow, UnitCol) = Units.Text
    GramsSelected(LastRow) = Units.GramsSelected
    FG.Row = LastRow
   
    If SelectedID <= -200 Then
       SelectedIDS(LastRow) = SelectedID
       FG.TextMatrix(LastRow, UnitCol) = "Make Into Meal"
    '   Call LeaveCell
       NoLeave = True
       For j = 0 To FG.Cols - 1
          FG.Col = j
          FG.CellBackColor = RGB(225, 225, 225)
          FG.TextMatrix(LastRow, j) = ""
       Next j
       FG.Col = UnitCol
       FG.CellForeColor = vbBlue
       FG.TextMatrix(LastRow, UnitCol) = "Make Into Meal"
       
       FG.Col = ACCol
'       FG.Row = LastRow + 1
       FG.CellBackColor = RGB(225, 225, 225)
       FG.CellFontBold = True
       NoLeave = False
       
    ElseIf SelectedID > -200 And SelectedIDS(LastRow) <= -200 Then
      SelectedIDS(LastRow) = SelectedID
      NoLeave = True
       For j = 0 To FG.Cols - 1
          FG.Col = j
          FG.CellBackColor = vbWhite
          FG.CellForeColor = 0
       Next j
       FG.Col = ACCol
       FG.CellBackColor = vbWhite
       FG.CellFontBold = False
       NoLeave = False
    End If
    SelectedIDS(LastRow) = SelectedID

End Sub

Private Sub AC_ItemSelected(SelectedID As Long)
On Error GoTo errhandl
  If SelectedID = -999 Then
     
     RaiseEvent InternetSearch(AC.Text)
     Exit Sub
  End If
  Call UpdateAC(SelectedID)
  Call AC_TabPress(False)
  Exit Sub
errhandl:
  If DoDebug Then MsgBox Err.Description, vbOKOnly, ""
End Sub

Private Sub AC_TabPress(Shift As Boolean)
On Error GoTo errhandl
  
  
   If SelectedIDS(LastRow) > -200 Then
     Call MoveCell(LastCol + 1, LastRow)
   Else
     Call MoveCell(ACCol, LastRow + 1)
   End If
errhandl:
End Sub

Public Sub DropRow(Text As String, Optional RowY As Single = -1)
On Error GoTo dout


   Dim i As Long, j As String
   Dim minRow As Long
   If RowY <> -1 Then
      minRow = SelectRow(RowY) - 2
      InsertRows
   Else
     minRow = -1
   End If
   
   For i = 0 To FG.Rows
      j = FG.TextMatrix(i, ACCol)
      If Trim$(j) = "" And i > minRow Then
         Call MoveCell(ACCol, i)
         AC.Text = Text
         AC.CloseWithText
         GoTo dout
      End If
   Next i
  
dout:
   
End Sub
Public Sub DropMeal(Text As String, dreplace As Boolean, DropDate As Date, Optional dShow As Boolean = True, Optional MealNumber As Long = -2)


On Error GoTo dout
Dim RS As Recordset, rs2 As Recordset, j As Long, Parts() As String
Dim ms(5) As String, nS(5) As Long, i As Long, MealName As String
If InStr(1, Text, "~~~") <> 0 Then
  Parts = Split(Text, "~~~")
  Text = Parts(2)
  MealName = Parts(1)
End If
ms(0) = "breakfast"
ms(1) = "brunch"
ms(2) = "lunch"
ms(3) = "snack"
ms(4) = "dinner"
ms(5) = "treat"
nS(0) = -200
nS(1) = -201
nS(2) = -202
nS(3) = -203
nS(4) = -204
nS(5) = -205


'determine which meal this is in the day
Set RS = DB.OpenRecordset("SELECT MealPlanner.*, MealDefinition.* " _
   & "FROM MealPlanner INNER JOIN MealDefinition ON MealPlanner.MealID = MealDefinition.MealID " _
   & "WHERE (((MealPlanner.Mealid)=" & Text & "));", dbOpenDynaset)
 'if the meal number is specified, use that, otherwise, use the default
If MealNumber = -2 Or OrderMeals = False Then
  For j = 0 To 5
   If LCase$(RS("meal")) = ms(j) Then
     i = nS(j)
     Exit For
   End If
  Next j
Else
  j = MealNumber
  i = nS(j)
End If

'remove the old meal
If dreplace Then
    If MealNumber = -2 Then MealNumber = j
    Set rs2 = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username & "' and " _
      & "date=#" & FixDate(DropDate) & "# and meal=" & MealNumber & ";", dbOpenDynaset)
    While Not rs2.EOF
      
      rs2.Delete
      rs2.MoveNext
    Wend
ElseIf OrderMeals = False Then
'get the largest meal entered in for this day
    Set rs2 = DB.OpenRecordset("select meal from daysinfo where " _
    & "user='" & CurrentUser.Username & "' and " _
    & "date=#" & FixDate(DropDate) & "#;", dbOpenDynaset)
    If IsNull(rs2("meal")) Then
       j = 0
    Else
      j = -1
      While Not rs2.EOF
         If j < rs2("meal") Then j = rs2("meal")
         rs2.MoveNext
      Wend
      j = j + 1
    End If
End If

Dim MealPosition As Long
MealPosition = j

Set rs2 = DB.OpenRecordset("select * from meals where user='" _
       & CurrentUser.Username & "' and entrydate=#" _
       & FixDate(DropDate) & "# and mealnumber=" & j & ";", dbOpenDynaset)

If rs2.EOF Then
  rs2.AddNew
Else
  rs2.Edit
End If

rs2("mealid") = Text
rs2("user") = CurrentUser.Username
rs2("mealnumber") = j
rs2("entrydate") = FixDate(DropDate)
rs2("downloaded") = 1
rs2.Update

'now get the index of the meal so it can be tracked
Set rs2 = DB.OpenRecordset("select * from meals where user='" _
       & CurrentUser.Username & "' and entrydate=#" _
       & FixDate(DropDate) & "# and mealnumber=" & j & ";", dbOpenDynaset)
       
Dim MealIndex As Long
MealIndex = rs2("id")
Set rs2 = Nothing

'now enter the meal ingredients

Dim orderN As Long
Set rs2 = DB.OpenRecordset("select * from daysinfo " & _
     "WHERE (((DaysInfo.date)=#" & FixDate(DropDate) & "#) AND " & _
     "(DaysInfo.user='" & CurrentUser.Username & "')) ORDER BY daysinfo.order" & _
     ";", dbOpenDynaset)
If Not rs2.EOF Then
  rs2.MoveLast
  orderN = rs2("order") + 1
Else
  orderN = 1
End If


rs2.AddNew
rs2("date") = DropDate 'FixDate(DropDate)
rs2("user") = CurrentUser.Username
rs2("itemid") = i
rs2("unit") = Left$(MealName, RS("unit").Size)
rs2("meal") = j
rs2("order") = orderN
rs2("mealid") = MealIndex
rs2.Update

While Not RS.EOF
   orderN = orderN + 1
   rs2.AddNew
   rs2("date") = DropDate 'FixDate(DropDate)
   rs2("user") = CurrentUser.Username
   rs2("itemid") = RS("abbrevid")
   rs2("unit") = RS("unit")
   rs2("servings") = RS("serving")
   rs2("meal") = j
   rs2("mealid") = MealIndex
   rs2("order") = orderN
   rs2.Update
   RS.MoveNext
Wend

RS.Close
rs2.Close
Set RS = Nothing
Set rs2 = Nothing

'if the meals need to be ordered then renumber the order parameter, otherwise just leave it as it falls

If OrderMeals Then
    Set rs2 = DB.OpenRecordset("select * from daysinfo " & _
     "WHERE (((DaysInfo.date)=#" & FixDate(DropDate) & "#) AND " & _
     "(DaysInfo.user='" & CurrentUser.Username & "')) ORDER BY daysinfo.meal, daysinfo.order" & _
     ";", dbOpenDynaset)
    orderN = 1
    While Not rs2.EOF
        rs2.Edit
        rs2("order") = orderN
        rs2.Update
        orderN = orderN + 1
        rs2.MoveNext
    Wend
End If



If dShow Then
  'MsgBox "Now Show"
  Call OpenDay(DropDate)
  Call UpdateRow(1)
End If
Exit Sub
dout:
'If DoDebug Then MsgBox "DropMeal" & vbCrLf & Err.Description
'If DoDebug Then Resume
End Sub


Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
  Call DeleteRows
End If

End Sub

Private Sub FG_LeaveCell()
 On Error Resume Next
  If Not NoLeave Then

    If LastCol = ServingCol Then ServingsSelected(LastRow) = Val(ConvertFractionToDecimal(TServing.Text))
  
    Call LeaveCell
  End If
End Sub

Private Sub FG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl

Dim Col As Long, Row As Long
Col = FG.MouseCol
Row = FG.MouseRow
Call LeaveCell

If Row = 0 Then Exit Sub
If (FG.Row <> FG.RowSel Or FG.Col <> FG.ColSel) And Button = 1 Then Exit Sub

If Button = 1 Then
  Call MoveCell(Col, Row)
Else 'if it is  a right click then line up the popup menu with the next line and show the options
  Dim i As Long, j As Long
  On Error Resume Next
  For i = 0 To FG.Rows - 1
     If Y < FG.RowPos(i) Then
       Y = FG.RowPos(i - 1)
       j = i
       Exit For
     End If
   Next i
   LastRow = i - 1
   If FG.Row <> FG.RowSel Then
      
   Else
      FG.Row = LastRow
   End If
   RaiseEvent PresentPopup(PopUp, j, X, Y, (SelectedIDS(Row) <= -200), FG.TextMatrix(i - 1, ACCol))
   'Call frmMain.PopUpMenu(PopUp, , X, Y)
End If
errhandl:
End Sub



Private Sub FG_Scroll()
On Error GoTo errhandl
If LastObject.Visible Then
  Call LocateObject(LastCol, LastRow)
End If
errhandl:
End Sub



Private Sub TServing_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = 13 Or KeyAscii = 9 Then
      ServingsSelected(LastRow) = Val(ConvertFractionToDecimal(TServing.Text))
      Call MoveCell(ACCol, LastRow + 1)
      KeyAscii = 0
  End If
End Sub

Private Sub Units_ItemSelected(Item As String, Conversion As Single)
On Error Resume Next
  GramsSelected(LastRow) = Conversion
  Call MoveCell(LastCol + 1, LastRow)
End Sub

Private Sub Units_TabPress(ShiftPressed As Boolean)
On Error Resume Next
   Call MoveCell(LastCol + 1, LastRow)

End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
  YFixed = False
  AC.InternetSearch = False
  
  ReDim Headers(0)
    
 
  Call MacroNut(FG.Rows - 1)
  
    Set LastObject = AC
    'usercontrol.Font.Name=
    'Call SetFont(UserControl.Font)
    FG.ColWidth(ACCol) = GetSetting(App.Title, "Settings", "ACWidth", TextWidth(Space(60)))
    FG.ColWidth(UnitCol) = GetSetting(App.Title, "Settings", "UnitWidth", TextWidth(Space(30)))
    AC.SetListBox Units
    NoLeave = False
End Sub
Private Sub SetFont(nFont As Object)
   On Error Resume Next
    Set UserControl.Font = nFont
    FG.Font.Charset = nFont.Charset
    FG.Font.Bold = nFont.Bold
    FG.Font.Name = nFont.Name
    FG.Font.Size = nFont.Size + 1
    FG.Font.Weight = nFont.Weight
    Set AC.Font = nFont
    Set Units.Font = nFont
    
    
    FG.ColWidth(ACCol) = TextWidth(Space(60))
    FG.ColWidth(UnitCol) = TextWidth(Space(30))

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=FG,FG,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
  On Error Resume Next
    BackColor = FG.BackColorBkg
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    FG.BackColorBkg() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=FG,FG,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Determines the color used to draw text on each part of the FlexGrid."
On Error Resume Next
    ForeColor = FG.ForeColor
    
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
On Error Resume Next
    FG.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=FG,FG,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
On Error Resume Next
    Enabled = FG.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next
    FG.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=FG,FG,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets the default font or the font for individual cells."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = FG.Font
    
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
   Call SetFont(New_Font)
   PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
On Error Resume Next
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
On Error Resume Next
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the FlexGrid. This includes all text, pictures, and cell formatting."
On Error Resume Next
  m_Calories = 0
  Call LeaveCell
  AC.Visible = False
  Units.Visible = False
  TServing.Visible = False
  FG.Rows = 1
  DoEvents
  FG.Rows = 5
  ReDim SelectedIDS(FG.Rows - 1)
  ReDim GramsSelected(FG.Rows - 1)
  ReDim ServingsSelected(FG.Rows - 1)
  ReDim SelectedIDS(0)
  ReDim GramsSelected(0)
  ReDim ServingsSelected(0)
  ReDim Stuff(UBound(Headers), 0)
  ReDim ExtraSelected(0)
  ReDim fat(0)
  ReDim fiber(0)
  ReDim Prot(0)
  ReDim carbs(0)
  ReDim sugar(0)
  ReDim MealIDs(0)
  Call MacroNut(FG.Rows - 1)
  
  Call UpdateRow(0)
  
  If m_YFixed Then Call SetYFixed(YTitles)
  Dim i As Long
  For i = 0 To PB.Rows - 1
    Call PB.UpdateLine(i, 0, 0)
  Next i
  
  PB.Draw
  For i = 0 To PB2.Rows - 1
    Call PB2.UpdateLine(i, 0, 0)
  Next i
  PB2.Draw
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function OpenDay(Today As Date) As Variant

  mDisplayDate = Today
  On Error GoTo errhandl
  Dim temp As Recordset, junk As String
  Dim MLs As Recordset
  
  
  Call Clear
  On Error Resume Next
  Dim i As Long
  For i = 0 To UBound(Headers)
    Maxs(i) = Maxes(Headers(i))
  Next i
  
  Dim MealPlans As Recordset
  Dim ID As Long, Serving As Single, Unit As String
  'open the days foods and load them into the data fields
  Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.date)=#" & FixDate(Today) & "#) AND (DaysInfo.user='" & Username & "')) ORDER BY daysinfo.order,itemid;", dbOpenDynaset)
  'get the date
   i = 0
  'if there are records, then put them into the data record
  If Not temp.EOF Then
     If Not (temp.EOF = True And temp.BOF = True) Then
        temp.MoveFirst
        i = 1
        Do While Not temp.EOF
          ID = 0
          ID = temp.Fields("itemID")
          On Error Resume Next
          Unit = ""
          Serving = 0
          
          If ID <> 0 And ID <> -1111 Then
             Unit = temp.Fields("unit")
             Serving = temp.Fields("Servings")
          End If
          temp.Edit
          temp("order") = i
          temp.Update

          If (Unit <> "" And Serving <> 0) Or (ID <= -200 And ID <> -555) Then
            Call SetRow(i, ID, Serving, Unit, "", False) ' Then i = i + 1
          End If
          
          If ID <= -200 And Trim$(Unit) = "" And ID <> -555 Then
           
             ID = Abs(-200 - ID)
             Set MLs = DB.OpenRecordset("select * from meals " _
               & "where user='" & CurrentUser.Username & "' and " _
               & "mealnumber = " & ID & " and " _
               & "entrydate=#" & FixDate(mDisplayDate) & "#;", dbOpenDynaset)
             If MLs.EOF Then
                NoLeave = True
                FG.TextMatrix(i, UnitCol) = "Make into Meal"
                FG.Col = UnitCol
                FG.Row = i
                FG.CellForeColor = vbBlue
                NoLeave = False
             End If
             MLs.Close
             Set MLs = Nothing
            
          End If
          temp.MoveNext
         
          i = i + 1
        Loop
     End If
    End If
    If FG.Rows = 5 And Trim$(FG.TextMatrix(1, 0)) = "" Then
      FG.TextMatrix(1, 0) = PromptPhrase
    End If
    'If FG.TextMatrix(FG.Rows - 1, ACCol) <> "" Then
       FG.Rows = FG.Rows + 4
    'End If
    
    If FG.Rows <= 2 Then
        Dim PT As Single
        Dim f As Single, c As Single, fb As Single, p     As Single, Calories As Single, s As Single
        Calories = Nutmaxes("Calories")
        f = Nutmaxes("Fat")
        c = Nutmaxes("Carbs")
        fb = Nutmaxes("Fiber")
        p = Nutmaxes("Protein")
        s = Nutmaxes("Sugar")
        Call Module1.FigurePercentages(Balance, Calories, f, s, c, p, fb)
    End If
    temp.Close
    Set temp = Nothing
    Call UpdateRow(1, False)
    Exit Function
errhandl:
    MsgBox Err.Description, vbOKOnly, ""
  ' Resume
End Function


Public Sub SaveDay(Today As Date)

  On Error GoTo errhandl
  Dim i As Long, junk As String
  Dim temp As Recordset, temp2 As Recordset, DI As Recordset
  Dim EditSet As Boolean
  If Username = "" Then Exit Sub
  Call LeaveCell
  
  

  On Error GoTo errhandl
 'shouldnt this have a full save to database to backup other actions???
    Dim ID As Long, Serving As Single, Unit As String, junks As Long
    Dim MealN As Long
    MealN = -1
    For i = 1 To FG.Rows - 1
      Call GetRow(i, ID, Serving, Unit, "")
      If (ID <> 0 And ID <> -1111 And ((Unit <> "" And Serving <> 0) Or (ID <= -200))) Or (ID <= -200) Then
           'now update the ingrediants
           Set temp2 = DB.OpenRecordset("SELECT ABBREV.index, ABBREV.Usage From ABBREV WHERE (((ABBREV.index)=" & ID & "));", dbOpenDynaset)
          
           On Error Resume Next
           junks = temp2.Fields("Usage") + 1
           If Err.Number <> 0 Then
             junks = 0
             Err.Clear
           End If
           temp2.Edit
           If junks > 2000 Then junks = 2000
           temp2.Fields("Usage") = junks
           
           temp2.Update
           temp2.Close
           Set temp2 = Nothing
      End If
     Next i
    
    ' Set temp = Nothing
     On Error GoTo errhandl
     'now put in the daily log maintence
    Set temp = DB.OpenRecordset("Select * from dailylog where ((date = #" & FixDate(Today) & "#) and (user = """ & Username & """));", dbOpenDynaset)
    'if there is nothing then put in a record with the proper start
    If temp.RecordCount = 0 Then
       temp.AddNew
       temp.Fields("user") = Username
       temp.Fields("date") = Today 'FixDate(Today)
    Else
       temp.Edit
    End If
    'put in the calories information
    temp.Fields("Calories") = m_Calories
    temp("bmr") = CurrentUser.BMR
    Dim f As Single, c As Single, p As Single, fb As Single, s As Single
    Dim j As Long
    For j = 0 To UBound(Stuff, 2)
      f = f + fat(j)
      c = c + carbs(j)
      p = p + Prot(j)
      fb = fb + fiber(j)
      s = s + sugar(j)
    Next j
    
    temp.Fields("Sugar") = s
    temp.Fields("Fat") = f
    temp.Fields("Carbs") = c
    temp.Fields("Protein") = p
    temp.Fields("Fiber") = fb
    
    temp.Update
    Set temp = Nothing
  Exit Sub
errhandl:
  MsgBox "Unable to save." & vbCrLf & Err.Description, vbOKOnly, ""
  
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddDataBase(DB As Database, nUserName As String, Today As Date, Optional NewMaxes As Calories = Nothing) As Calories
   On Error Resume Next
   Set mDB = DB
   If NewMaxes Is Nothing Then
    Set AddDataBase = Maxes
   Else
    Set Maxes = NewMaxes
   End If
   Username = nUserName
   AC.AddDataBase DB
   Units.AddDataBase DB
   
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   On Error Resume Next
    YFixed = False
    m_Changed = m_def_Changed
    m_ShowAsPercent = m_def_ShowAsPercent
    m_Text = m_def_Text
    ShowTotals = m_def_ShowTotals
    AC.InternetSearch = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   On Error Resume Next
    YFixed = PropBag.ReadProperty("Yfixed", False)
    m_ShowAsPercent = PropBag.ReadProperty("ShowAsPercent", m_def_ShowAsPercent)
    FG.BackColorBkg = PropBag.ReadProperty("BackColor", &H80000005)
    FG.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    FG.Enabled = PropBag.ReadProperty("Enabled", True)
    Call SetFont(PropBag.ReadProperty("Font", Ambient.Font))
  ' AC.InternetSearch = True
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
   
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    ShowTotals = PropBag.ReadProperty("ShowTotals", m_def_ShowTotals)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim ROH As Single
Dim TW As Single, th As Single, LW As Single
Dim BW As Single, BH As Single
Dim No As Long, Nd As Long, n As Long, i As Long, j As Long
Dim T As Single, fsize As Single
Dim cc As Single

If m_ShowTotals Then
    BW = ScaleX(1.5, vbInches, UserControl.ScaleMode)
    BH = ScaleY(1.5, vbInches, UserControl.ScaleMode)
    
    cc = ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
    If cc < 718 Then
        T = 1.5 * cc / 718
        BW = ScaleX(T, vbInches, UserControl.ScaleMode)
        BH = ScaleY(T, vbInches, UserControl.ScaleMode)
        fsize = 10 * cc / 718
    Else
        fsize = 10
    End If
    
        ReadOut = BH
        T = ScaleHeight - BH
        PB.Move 0, T, BW * 2, BH
        PB2.Move BW * 2 + 10, T, BW * 2, BH
        PB.Font.Size = Round(fsize)
        PB.Font.Bold = True
        PB2.Font.Size = Round(fsize)
        PB2.Font.Bold = True
        
        Balance.Move BW * 4 + 20, T, BW, BH
        FG.Move 0, 0, ScaleWidth, T
        
    
Else
    FG.Move 0, 0, ScaleWidth, ScaleHeight
End If

If LastObject.Visible Then
  Call LocateObject(LastCol, LastRow)
End If

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
    Call PropBag.WriteProperty("Yfixed", YFixed, False)
    Call PropBag.WriteProperty("BackColor", FG.BackColorBkg, &H80000005)
    Call PropBag.WriteProperty("ForeColor", FG.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", FG.Enabled, True)
    Call PropBag.WriteProperty("Font", FG.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("SaveDay", m_SaveDay, m_def_SaveDay)
    Call PropBag.WriteProperty("Rows", FG.Rows, 2)
    Call PropBag.WriteProperty("Changed", m_Changed, m_def_Changed)
    Call PropBag.WriteProperty("ShowAsPercent", m_ShowAsPercent, m_def_ShowAsPercent)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("ShowTotals", m_ShowTotals, m_def_ShowTotals)
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=FG,FG,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in a FlexGrid."
   On Error Resume Next
    Rows = FG.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
On Error Resume Next
    FG.Rows() = New_Rows
    PropertyChanged "Rows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,false
Public Property Get Changed() As Boolean
Attribute Changed.VB_MemberFlags = "400"
On Error Resume Next
    Changed = m_Changed
End Property

Public Property Let Changed(ByVal New_Changed As Boolean)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 387
    m_Changed = New_Changed
    PropertyChanged "Changed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowAsPercent() As Boolean
On Error Resume Next
    ShowAsPercent = m_ShowAsPercent
End Property

Public Property Let ShowAsPercent(ByVal New_ShowAsPercent As Boolean)
On Error Resume Next
    m_ShowAsPercent = New_ShowAsPercent
    PropertyChanged "ShowAsPercent"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
On Error Resume Next
    Text = FG.Text
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get ShowTotals() As Boolean
On Error Resume Next
    ShowTotals = m_ShowTotals
End Property

Public Property Let ShowTotals(ByVal New_ShowTotals As Boolean)
On Error Resume Next
    m_ShowTotals = New_ShowTotals
    PB.Visible = New_ShowTotals
    PB2.Visible = New_ShowTotals
    Balance.Visible = New_ShowTotals
    Call UserControl_Resize
    PropertyChanged "ShowTotals"
End Property
'Public Property Get DisplayDate() As date
'   DisplayDate = mDisplayDate
'End Property
'Public Property Let DisplayDate(ByVal new_date As date)
  
'   mDisplayDate = new_date
'   Call OpenDay(new_date)
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get YFixed() As Boolean
On Error Resume Next
  YFixed = m_YFixed
End Property

Public Property Let YFixed(ByVal New_YFixed As Boolean)
On Error Resume Next
  m_YFixed = New_YFixed
  If YFixed = False Then
    FG.FixedCols = 0
    ACCol = 0
    UnitCol = 1
    ServingCol = 2
    FlagCol = 3
    ReservedCol = 3
  Else
    FG.FixedCols = 1
    ACCol = 1
    UnitCol = 2
    ServingCol = 3
    FlagCol = 4
    ReservedCol = 4
  End If
End Property

Public Sub SetYFixed(Titles)
On Error Resume Next
   YTitles = Titles
   Dim i As Long
   If FG.Rows < UBound(Titles) + 2 Then FG.Rows = UBound(Titles) + 4
   For i = 0 To UBound(Titles)
      FG.TextMatrix(i, 0) = Titles(i)
   Next i
End Sub
Public Sub Refresh()
On Error Resume Next
  Balance.DrawGraph
End Sub

Public Function GetTotalCalories(MacroNutrients() As Single)
On Error Resume Next
  MacroNutrients = TMacro
  GetTotalCalories = m_Calories
  
End Function

Public Sub SetExerciseCals(New_Cals As Single)
On Error Resume Next
   ExerciseCals = New_Cals
   Dim i As Long, Pcnt As Single
   For i = 0 To UBound(Headers)
       If LCase$(Headers(i)) = "calories net" Then
         Pcnt = Round((m_Calories - ExerciseCals) / Maxs(0) * 100)
         If i > 5 Then
           Call PB2.UpdateLine(i - 6, Pcnt, Round(m_Calories - ExerciseCals))
         Else
           Call PB.UpdateLine(i, Pcnt, Round(m_Calories - ExerciseCals))
           PB.Draw
         End If
       End If
   Next
   Call UpdateRow(1)
End Sub
Public Function GetAllTotals() As exCollection
On Error GoTo errhandl
  Dim i As Long, j As Long
  Dim JunkCol As Collection, junk As Boolean
  Dim headCol As Long
  Dim X, Serving As Single, ID As Long, Unit As String, Grams As Single
  Dim Names As Collection
  Dim TotGrams As Single
  Dim AllTotals() As Single
  
  TotGrams = 0
  
  
  
  For i = 1 To UBound(SelectedIDS)
    junk = True
    ID = SelectedIDS(i)
    Serving = ServingsSelected(i)
    Grams = GramsSelected(i)
    TotGrams = TotGrams + Grams
    If ID = 0 Or ID = -1111 Then junk = False
    If Serving = 0 Then junk = False
    
    
    If junk Then
      Set Names = Nothing
      Set JunkCol = GetNutrients(ID, "abbrev", "index", Names)
      ReDim Preserve AllTotals(Names.Count)
      For j = 1 To JunkCol.Count
        X = Val(JunkCol(j))
        X = X * Serving
        X = X / 100 * Grams
        AllTotals(j) = AllTotals(j) + X
      Next j
    End If
  Next i
  
  Dim JunkName As String
  If Names Is Nothing Then
    Exit Function
  Else
    Set GetAllTotals = New exCollection
    For i = 1 To Names.Count
      JunkName = LCase$(Names(i))
      If JunkName <> "foodgroup" And JunkName <> "index" And JunkName <> "ndb_no" And JunkName <> "foodname" Then
         GetAllTotals.Add AllTotals(i), JunkName
      End If
    Next i
  End If
  GetAllTotals.Add TotGrams, "Total Grams"
  Exit Function
errhandl:
  MsgBox "Cannot do totals." & vbCrLf & Err.Description, vbOKOnly, ""
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

      Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
