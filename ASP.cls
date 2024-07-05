VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Dim CodeSplit() As String
Dim HTMLString() As String
Dim codeString As String
Dim ASPString() As String
Dim ASPSplitCount
Dim HTMLSplitCount
Dim CodeSplitCount
Dim Template As String

Private HTMLOutput As String
Private GS As uGraphSurface
Private SC As ScriptControl
Private LastTime As Long
Private OutPath As String
Private Graphnumber As Long
Private SessionCookies As New Collection
Private mForm As New Collection

Private Type MinInfo
  ID As Long
  Calories As Single
  Weight As Single
  eDate As Date
End Type

Public PassPhrase As String

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function DoDaysNutrition() As String
On Error GoTo errhandl

  'CurrentUser.Username = "brian"
  'DisplayDate = "12/8/2007"
  Dim junk As String, ff As Long, temp As String
  Dim i As Long
  ff = FreeFile
  Open App.path & "\resources\daily\nutritiontemp.htm" For Input As #ff
    temp = ""
    While Not EOF(ff)
      Line Input #ff, junk
      temp = temp & junk
    Wend
  Close #ff
  Dim s As Single
  Dim RS As Recordset, Abb As Recordset, W As Recordset
  Dim Cols As Collection, grid() As Single, j As Long, nS As Collection
  Set RS = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username & _
     "' and date=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
  While Not RS.EOF
    Set Abb = DB.OpenRecordset("select * from abbrev where index=" & RS("itemid") & ";", dbOpenDynaset)
    If j = 0 Then
       Set Cols = New Collection
       Set nS = New Collection
       ReDim grid(Abb.Fields.Count - 1)
       For i = 0 To Abb.Fields.Count - 1
          Cols.Add i, Abb.Fields(i).Name
          nS.Add Abb.Fields(i).Name
       Next i
    End If
    If RS("itemid") >= 0 Then
    Set W = DB.OpenRecordset("select * from weight where index=" & RS("itemid") & " and msre_desc='" & RS("unit") & "';", dbOpenDynaset)
    For i = 0 To Abb.Fields.Count - 1
    '  Debug.Print VarType(Abb(i)), Abb(i)
      If Not IsNull(Abb(i)) Then
        If VarType(Abb(i)) <> 8 Then
         s = Abb(i) / 100 * W("gm_wgt") * RS("servings") / W("amount")
         grid(i) = grid(i) + s
        End If
      End If
    Next i
    End If
    j = j + 1
    RS.MoveNext
  Wend
  Set RS = Nothing
  Set W = Nothing
  Dim ii As Recordset
  Set ii = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If Not nS Is Nothing Then
  On Error Resume Next
  Dim js As Single
  For i = 1 To nS.Count
    junk = nS(i)
    junk = Replace(junk, " ", "_")
    If LCase$(junk) = "carbs" Then junk = "carbohydrates"
    'Debug.Print junk
    'Debug.Print nS(i), Cols(nS(i)), grid(Cols(nS(i))), ii(nS(i))
    js = grid(Cols(nS(i)))
    temp = Replace(temp, "#" & junk & "#", Round(js, 1), , , vbTextCompare)
    temp = Replace(temp, "#" & junk & "p#", Round(js / ii(nS(i)) * 100), , , vbTextCompare)
  Next i
End If
'Clipboard.SetText temp
  DoDaysNutrition = temp
  Set ii = Nothing
  Exit Function
errhandl:
   
  If DoDebug Then
    Debug.Print Err.Description
    Stop
    Resume
  End If
End Function
Public Function sort(inArray)
On Error GoTo errhandl
Dim T, i As Long, j As Long, k As Long
ReDim T(UBound(inArray, 1), UBound(inArray, 2))
j = 0
For i = 0 To UBound(inArray, 2)
  If inArray(0, i) <> 0 Then
    For k = 0 To UBound(inArray, 1)
     T(k, j) = inArray(k, i)
    Next k
     j = j + 1
  End If
Next i

ReDim Preserve inArray(UBound(inArray, 1), j - 1)

For i = 0 To UBound(inArray, 2)
    For k = 0 To UBound(inArray, 1)
     inArray(k, i) = T(k, i)
    Next k
Next i

errhandl:

End Function



Public Sub GetRows(QueryString, Rows, indexs)


    On Error GoTo Err_Proc
   Dim RS As Recordset
   Dim i As Long, j As Long, T As New Collection
   Set RS = DB.OpenRecordset(QueryString, dbOpenDynaset)
   If RS.EOF Then
     ReDim Rows(0, 0)
     Exit Sub
   End If
   RS.MoveLast
   RS.MoveFirst
   ReDim Rows(RS.RecordCount, RS.Fields.Count - 1)
   For j = 0 To RS.RecordCount - 1
     For i = 0 To RS.Fields.Count - 1
      Rows(j, i) = RS.Fields(i)
     Next i
     RS.MoveNext
   Next j
   For i = 0 To RS.Fields.Count - 1
     T.Add i, RS.Fields(i).Name
   Next i
   Dim junk As String
   For i = 0 To UBound(indexs)
     junk = indexs(i)
     If LCase$(junk) = "weightloss" Then junk = "weight"
     indexs(i) = T(junk)
   Next i
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "GetRows", Err.Description
    Resume Exit_Proc


End Sub
Public Sub PrepReports()
    Dim DL As Recordset
    Dim nDays As Long, MinDate As Date
    Dim rDays() As MinInfo, n As Long
    Dim MaxIndex As Long
    
    Set DL = DB.OpenRecordset("Select * from dailylog where user='small';", dbOpenDynaset)
    While Not DL.EOF
      DL.Delete
      DL.MoveNext
    Wend
    DL.Close
     Dim i As Long

    
    Set DL = DB.OpenRecordset("select max(id) as maxit from dailylog;", dbOpenDynaset)
    MaxIndex = DL("maxit") + 1
    DL.Close
    
    Set DL = DB.OpenRecordset("Select * from dailylog where user='" & CurrentUser.Username & "' " _
    & " and date<# " & FixDate(Today) & "# order by date;", dbOpenDynaset)
    MinDate = DL("date")
    ReDim rDays(0)
    While Not DL.EOF
      n = DateDiff("d", MinDate, DL("date"))
      If n > UBound(rDays) Then
        ReDim Preserve rDays(n)
      End If
      With rDays(n)
        .ID = DL("Id")
        .Calories = DL("calories")
        .Weight = DL("weight")
        .eDate = DL("date")
      End With
      DL.MoveNext
    Wend
    DL.Close
    Dim Fields()
    Dim cc() As Long
    
    Dim found As Boolean
    For nDays = 0 To 0
           For n = 0 To UBound(rDays) - nDays - 1
               If rDays(n).Calories <> 0 And rDays(n).Weight <> 0 And _
                  rDays(n + nDays).Calories <> 0 And rDays(n + nDays + 1).Weight <> 0 Then
                  found = False
                  For i = n To n + nDays
                     If rDays(i).Calories = 0 Then found = True
                  Next i
                  If Not found Then
                     Set DL = DB.OpenRecordset("Select * from dailylog where user='" & CurrentUser.Username & "' " _
                       & " and date>=# " & FixDate(rDays(n).eDate) & "# and date <=#" & FixDate(rDays(n + nDays).eDate) & "# order by date;", dbOpenDynaset)
                     Erase Fields
                     ReDim Fields(DL.Fields.Count)
                     Erase cc
                     ReDim cc(DL.Fields.Count)
                     
                     While Not DL.EOF
                       For i = 0 To DL.Fields.Count - 1
                         If LCase$(DL(i).Name) = "calories" Then
                         If Val(DL(i) & " ") <> 0 Then
                          Fields(i) = Fields(i) + DL(i)
                          
                          cc(i) = cc(i) + 1
                         End If
                         Else
                         If Val(DL(i) & " ") <> 0 Then
                          Fields(i) = Fields(i) + DL(i)
                          cc(i) = cc(i) + 1
                         End If
                         End If
                       Next i
                       DL.MoveNext
                     Wend
                     DL.Close
                     Set DL = DB.OpenRecordset("Select * from dailylog;", dbOpenDynaset)
                     DL.AddNew
                     On Error Resume Next
                     For i = 1 To DL.Fields.Count - 1
                        If LCase$(DL(i).Name) = "calories" Then
                           DL(i) = Fields(i)
                        Else
                           DL(i) = Fields(i) / cc(i)
                        End If
                     Next i
                    ' On Error GoTo 0
                     DL("user") = "small"
                     
                     DL("calories") = DL("calories") / (nDays + 1)
                     DL("date") = DateAdd("d", n, MinDate) 'FixDate(DateAdd("d", n, MinDate))
                     
                     DL("weight") = (rDays(n + nDays + 1).Weight - rDays(n).Weight) / (nDays + 1)
                 
                     DL("Id") = MaxIndex
                     
                     DL.Update
                     MaxIndex = MaxIndex + 1
                  End If
               End If
           Next n
    Next nDays

    
    Exit Sub
  
End Sub


Public Sub SaveForm(d, ByRef Cancel As Boolean)
If d Is Nothing Then Exit Sub
Cancel = True
Dim values As New Collection, Names As New Collection
Dim i As Long, Fdate As Date
Dim e, Es, junk As String
Dim NewName As Boolean
Set Es = d.getElementsByTagName("input")
Fdate = "2000-01-01"
'gather all the form elements
For Each e In Es
  If LCase$(e.Type) = "radio" Then
     If e.Checked Then
       values.Add e.Value, e.Name
       Names.Add e.Name
     End If
  ElseIf LCase$(e.Type) = "checkbox" Then
     On Error Resume Next
     Names.Add e.Name
     
     If e.Checked Then
       junk = ""
       junk = e.Value
       If junk = "" Then
          values.Add True, e.Name
       Else
          values.Add junk, e.Name
       End If
       If Err.Number = 457 Then
          junk = values(e.Name) & "," & e.Value
          values.Remove e.Name
          values.Add junk, e.Name
       End If
     Else
       junk = ""
       junk = e.Value
       If junk = "" Then
          values.Add False, e.Name
       End If
       'values.Add False, E.name
     End If
    ' On Error GoTo 0
  ElseIf Trim$(e.Name) <> "" Then
     values.Add e.Value, e.Name
     Names.Add e.Name
     If LCase$(e.Name) = "log_date" Then
        Fdate = e.Value
     End If
     If LCase$(e.Name) = "goweb" Then
        Cancel = False
     End If
     If LCase$(e.Name) = "profile_user" Then
        NewName = True
     End If

  End If
Next
Set Es = d.getElementsByTagName("select")
For Each e In Es
     values.Add e.Value, e.Name
     Names.Add e.Name
Next
Set Es = d.getElementsByTagName("textarea")
For Each e In Es
     values.Add e.innerText, e.Name
     Names.Add e.Name
Next
junk = ""


If NewName Then
  junk = Trim$(values("profile_user"))
  If junk <> "" Then CurrentUser.Username = values("Profile_user")
End If
'now sort them into the needed categories

Dim p As Recordset, L As Recordset, Idl As Recordset, POpen As Boolean, LOpen As Boolean, IdlOpen As Boolean
Dim PFound As Boolean, Lfound As Boolean, j As Long

Set p = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If p.EOF Then
  p.AddNew
  p("user") = CurrentUser.Username
  p.Update
  Set p = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
End If
Set Idl = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If Idl.EOF Then
  Idl.AddNew
  Idl("user") = CurrentUser.Username
  Idl.Update
  Set Idl = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
End If
'if there is  date field then put info into the log
If Fdate <> "2000-01-01" Then
  Set L = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "' and date=#" & FixDate(Fdate) & "#;", dbOpenDynaset)
End If
'clear out the form thing
Set mForm = Nothing
Set mForm = New Collection
On Error Resume Next
For i = 1 To mForm.Count - 1
  mForm.Remove i
Next i

For i = 1 To Names.Count
  junk = Names(i)
  mForm.Add Trim$(values(i)), junk
  PFound = False
  Lfound = False
  
  If InStr(1, junk, "Ideals_", vbTextCompare) <> 0 Then
    If IdlOpen = False Then
       IdlOpen = True
       Idl.Edit
    End If
    junk = Replace(LCase$(junk), "ideals_", "")
    For j = 0 To Idl.Fields.Count - 1
       If junk = LCase$(Replace(Replace(Idl(j).Name, " ", "_"), "-", "_")) Then
         Idl(j) = Val(values("ideals_" & junk))
       End If
    Next j
  ElseIf InStr(1, junk, "profile_", vbTextCompare) <> 0 Then
       PFound = True
       If POpen = False Then
          p.Edit
          POpen = True
       End If
       Dim junk2 As String
       junk2 = Replace(junk, "profile_", "", , , vbTextCompare)
       p(junk2) = values(junk)
       If LCase$(junk2) = "user" Then
          CurrentUser.Username = values(junk)
       End If
  ElseIf Fdate <> "2000-01-01" And InStr(1, junk, "log_", vbTextCompare) <> 0 Then
       If LOpen = False Then
          LOpen = True
          If L.EOF Then
            L.AddNew
            L("date") = Fdate 'FixDate(Fdate)
            L("user") = CurrentUser.Username
          Else
            L.Edit
          End If
       End If
       L(Replace(junk, "log_", "", , , vbTextCompare)) = values(junk)
  End If
Next i

Dim SS As Double
SS = p("sugar") + p("carbs") + p("protein") + p("fat")
p("sugar") = p("sugar") / SS
p("fat") = p("fat") / SS
p("protein") = p("protein") / SS
p("carbs") = p("carbs") / SS
If POpen Then p.Update
If LOpen Then L.Update
If IdlOpen Then Idl.Update

On Error Resume Next
p.Close
L.Close
Idl.Close
Set p = Nothing
Set L = Nothing
Set Idl = Nothing


End Sub

Public Function Form(vName As String)
 On Error Resume Next
  Form = mForm(vName)
End Function
Public Function Session(vName As String, Optional Value = -10000000)
If Value = -10000000 Then
  Session = SessionCookies(vName)
Else
  On Error Resume Next
  SessionCookies.Add Value, vName
  Session = Err.Number
End If
End Function
Public Sub HTMLWrite(s)


    On Error GoTo Err_Proc
     HTMLOutput = HTMLOutput & s

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "HTMLWrite", Err.Description
    Resume Exit_Proc


End Sub
Public Sub HTMLLine(s)
On Error Resume Next
  Dim junk As String
  junk = s
  If InStr(1, junk, "@@~~@@") <> 0 Then
     HTMLOutput = HTMLOutput & Replace(junk, "@@~~@@", vbCrLf)
  Else
     HTMLOutput = HTMLOutput & s
  End If
End Sub

Public Function MakeFoodLog(InputType As String, nDays As Long) As String
On Error GoTo errhandl
Dim StartDate As Date
Dim nWeeks As Single
Dim RS As Recordset
Dim cc As Single, i As Long, j As Single
Dim Nuts As New Collection, NutNames As New Collection
Dim TT As Date
Dim values As New Collection
Dim OutString As String
Dim ff As Long, jj As Single
Dim jjj As Single

If LCase$(InputType) = "user" Then
   nWeeks = Val(InputBox("Please enter the number of weeks you wish to include in report?", "", nDays))
ElseIf LCase$(InputType) = "last" Then
   nWeeks = LastTime / 7
Else
   nWeeks = nDays / 7
End If

LastTime = nWeeks * 7
StartDate = DateAdd("d", Int(Val(nWeeks) * -7), DisplayDate)



ff = FreeFile
    
    For j = 0 To UBound(WatchHeads)
      values.Add 0, WatchHeads(j)
    Next j

OutString = OutString & "<table width=""100%"" border=0 cellspacing=0 cellpadding=0>"

For i = DateDiff("d", DisplayDate, StartDate) To 0
   TT = DateAdd("d", i, DisplayDate)
   Set RS = DB.OpenRecordset("SELECT DaysInfo.*, Abbrev.* " _
     & "FROM (DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID = Abbrev.Index) " _
     & "where user ='" & CurrentUser.Username & "' and date=#" & FixDate(TT) & "# order by [order];", dbOpenDynaset)
    For j = 0 To UBound(WatchHeads)
      values.Remove WatchHeads(j)
      values.Add 0, WatchHeads(j)
    Next j
   If Not RS.EOF Then
     nDays = nDays + 1
     OutString = OutString & "<tr><td colspan=4><h2>" & TT & "</h2></td></tr>"
     While Not RS.EOF
       If RS("index") < 0 Then
          OutString = OutString & "<tr><td  bgColor=#c0c0c0  colspan = 3><b>" & RS("Foodname") & "</b><!--AbbrevID=" & RS("index") & "&ndb_no=" & RS("ndb_no") & "--></td>"  '</tr>"
          For j = 0 To 4
            OutString = OutString & "<td bgColor=#c0c0c0>&nbsp;" & WatchHeads(j) & "</td>"
          Next j
          OutString = OutString & "</tr>" & vbCrLf
       Else
          cc = Module1.TranslateUnitToGrams(RS("index"), RS("unit"))
          OutString = OutString & "<tr><td><!--grams=" & (cc * RS("servings")) & "-->" & Module1.ConvertDecimalToFraction(RS("servings")) & "</td><td>" & RS("unit") & "</td><td>" & RS("Foodname") & "<!--AbbrevID=" & RS("index") & "&ndb_no=" & RS("ndb_no") & "--></td>"
          
          cc = cc * RS("servings") / 100 'rs("gm_wgt") / rs("amount") * rs("servings") / 100
          jj = 0
          For j = 0 To 4
             jjj = cc * Val(RS(WatchHeads(j)) & "")
             OutString = OutString & "<td>&nbsp; " & Round(jjj, 1) & "</td>"
             jj = values(WatchHeads(j))
             values.Remove (WatchHeads(j))
             values.Add jj + jjj, WatchHeads(j)
          Next j
          OutString = OutString & "</tr>" & vbCrLf
       End If
       RS.MoveNext
     Wend
     OutString = OutString & "<tr><td height=""3px"" colspan=20 style=""border-bottom-style:solid"">&nbsp;</td>"
     OutString = OutString & "<tr><td>Totals</td><td>&nbsp;</td><td>&nbsp;</td>"
     For j = 0 To 4 'UBound(WatchHeads)
       OutString = OutString & "<td >" & Round(values(WatchHeads(j))) & "</td>"
     Next j
     OutString = OutString & "</tr>"
   End If
   
   Set RS = DB.OpenRecordset("SELECT ExerciseLog.Week, AbbrevExercise.ExerciseName, AbbrevExercise.Formula, ExerciseLog.weekInfo, ExerciseLog.User, ExerciseLog.Order " _
                           & " FROM ExerciseLog INNER JOIN AbbrevExercise ON ExerciseLog.exerciseID=AbbrevExercise.Index " _
                           & " Where (((ExerciseLog.Week) = #" & FixDate(FindFirstDay(TT)) & "#) And ((ExerciseLog.USER) = '" & CurrentUser.Username & "')) " _
                           & " ORDER BY ExerciseLog.Order;", dbOpenDynaset)
   Dim junk As String, junks() As String, first As Boolean
   first = False
   
   If RS.EOF = False Then
       While Not RS.EOF
         junk = RS("weekinfo")
         junks = Split(junk, "~")
         j = Abs(DateDiff("d", TT, FindFirstDay(TT)))
         If Trim$(junks(j)) <> "" Then
           If first = False Then OutString = OutString & "<tr><td  colspan=20><table><tr><td bgColor=#c0c0c0 colspan=2><h3>Exercise Statistics</h3></td></tr>"
           OutString = OutString & "<tr><td>" & RS("exercisename") & "</td><td>" & junks(j) & "</td></tr>"
           first = True
         End If
         RS.MoveNext
       Wend
       If first Then OutString = OutString & "</table><br><br></td></tr>" & vbCrLf
   End If
   
   Set RS = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "' and date=#" & FixDate(TT) & "#;", dbOpenDynaset)
   If RS.EOF = False Then
     If RS("calories") <> 0 Or RS("exercise_cal") <> 0 Then
         OutString = OutString & "<tr><td colspan=20><table width='100%' cellpadding=0 cellspacing=0><tr><td bgColor=#c0c0c0 colspan=20 > " _
                          & "<h3>Day's Information</h3></td></tr><tr>"
          For j = 3 To RS.Fields.Count - 1
             If RS.Fields(j).Name = "e1" Then Exit For
             If j Mod 2 = 0 Then
                OutString = OutString & "<td><stronger>" & CleanDaily(RS.Fields(j).Name) & "</stronger></td>"
             Else
                OutString = OutString & "<td bgColor=#c0c0c0> <stronger>" & CleanDaily(RS.Fields(j).Name) & "</stronger></td>"
             End If
          Next j
          OutString = OutString & "</tr><tr>"
          For j = 3 To RS.Fields.Count - 1
             If RS.Fields(j).Name = "e1" Then Exit For
             If LCase$(RS.Fields(j).Name) <> "comments" Then
               If j Mod 2 = 0 Then
                OutString = OutString & "<td align=center>" & Round(RS.Fields(j), 2) & "</td>"
               Else
                OutString = OutString & "<td bgColor=#c0c0c0 align=center>" & Round(RS.Fields(j), 2) & "</td>"
               End If
             Else
               If j Mod 2 = 0 Then
                OutString = OutString & "<td align=center></td>"
               Else
                OutString = OutString & "<td bgColor=#c0c0c0 align=center></td>"
               End If
             End If
          Next j
          OutString = OutString & "</tr></table></td></tr>"
      End If
   End If
Next i
OutString = OutString & "</table>"

HTMLOutput = HTMLOutput & OutString
MakeFoodLog = OutString
Exit Function
errhandl:
HTMLOutput = HTMLOutput & OutString
MakeFoodLog = OutString

End Function

Private Function CleanDaily(junk As String) As String


    On Error GoTo Err_Proc
   If LCase$(junk) = "exercise_cal" Then junk = "Exercise Cals"
   If LCase$(junk) = "sore" Then junk = "Soreness"
   CleanDaily = junk
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "ASP", "CleanDaily", Err.Description
    Resume Exit_Proc


End Function


Public Sub RadioOptions(varname, Range, default)


    On Error GoTo Err_Proc
Dim junk As String, i As Long, j As Long
j = Val(SC.Eval("log(""" & varname & """)"))
If j = 0 Then j = default
For i = 1 To Range
  If j = i Then
    junk = junk & "<option value=" & i & " selected >" & i & "</option>"
  Else
    junk = junk & "<option value=" & i & ">" & i & "</option>"
  End If
Next i
HTMLOutput = HTMLOutput & junk
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "RadioOptions", Err.Description
    Resume Exit_Proc


End Sub

Public Sub EquationGraph( _
CaptionName As String, _
Xname As String, Yname As String, _
xvalue As String, yEquation As String)

On Error Resume Next
Dim RS As Recordset, k As Long, dd As Single
Set RS = DB.OpenRecordset("select * from dailylog where user='small';", dbOpenDynaset)
GS.Reset
GS.CaptionFont.Size = 13
GS.OverSizeX = True
GS.OverSizeY = True
GS.ShowLegend = False
GS.YAxisName = Yname
GS.XAxisName = Xname
xvalue = Trim$(xvalue)
Dim DisplayData() As Single, i As Long, junk As String
While Not RS.EOF
   ReDim Preserve DisplayData(1, k)
   For i = 0 To RS.Fields.Count - 1
     If LCase$(RS.Fields(i).Name) <> "bmr" Then
        junk = RS.Fields(i).Name & "=" & RS(i)
        Call SC.ExecuteStatement(junk)
     End If
   Next i
   
   DisplayData(0, k) = RS(Replace(xvalue, "p_", "", , , vbTextCompare))
   DisplayData(1, k) = SC.Eval(Replace(yEquation, "=", ""))
   dd = RS("calories")
   If InStr(1, xvalue, "p_", vbTextCompare) <> 0 Then
        If xvalue = "p_fat" Then DisplayData(0, k) = DisplayData(0, k) * 9 / dd * 100
        If xvalue = "p_protein" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
        If xvalue = "p_carbs" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
        If xvalue = "p_sugar" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
   End If
   RS.MoveNext
   k = k + 1
Wend

If UBound(DisplayData, 2) < 2 Then
   GS.DrawGraph
   GS.DrawCentered "Not Enough Datapoints." & vbCrLf & "Please record more days in journal."
   GoTo Exitsub

End If

Dim slope As Double, B As Double
Dim T() As Single
Call Module1.LinearFit(DisplayData, slope, B)
ReDim T(1, UBound(DisplayData, 2))
For i = 0 To UBound(DisplayData, 2)
   T(1, i) = slope * DisplayData(0, i) + B
   T(0, i) = DisplayData(0, i)
Next i
    
GS.MonthAxis = False
GS.AddLine "", gLine, vbRed, T
GS.AddLine "", gScatter, 0, DisplayData
GS.DrawGraph
Exitsub:
Dim Filename As String
Filename = "E" & Graphnumber & ".bmp"
Graphnumber = Graphnumber + 1
GS.SaveBitmap OutPath & "\" & Filename
HTMLOutput = HTMLOutput & "<img src=""" & OutPath & "\" & Filename & """ border=0>"
End Sub
Public Sub Graph( _
CaptionName As String, _
Xname As String, _
Yname As String, _
ShowLinear As Boolean, _
Data)
On Error Resume Next

Dim DisplayData() As Single, i As Long
ReDim DisplayData(1, UBound(Data, 2))
For i = 0 To UBound(Data, 2)
   DisplayData(0, i) = Data(0, i)
   DisplayData(1, i) = Data(1, i)
Next i
GS.Reset
GS.CaptionFont.Size = 13
GS.OverSizeX = True
GS.OverSizeY = True
GS.ShowLegend = False
GS.YAxisName = Yname
GS.XAxisName = Xname
  

If UBound(DisplayData, 2) < 2 Then
   GS.DrawGraph
   GS.DrawCentered "Not Enough Datapoints." & vbCrLf & "Please record more days in journal."
   GoTo Exitsub
End If

Dim slope As Double, B As Double
Dim T() As Single
Call Module1.LinearFit(DisplayData, slope, B)
ReDim T(1, UBound(DisplayData, 2))
For i = 0 To UBound(DisplayData, 2)
   T(1, i) = slope * DisplayData(0, i) + B
   T(0, i) = DisplayData(0, i)
Next i
    
GS.MonthAxis = False
GS.AddLine "", gLine, vbRed, T
GS.AddLine "", gScatter, 0, DisplayData
GS.DrawGraph
Exitsub:
Dim Filename As String
Filename = "N" & Graphnumber & ".bmp"
Graphnumber = Graphnumber + 1

GS.SaveBitmap OutPath & "\" & Filename
HTMLOutput = HTMLOutput & "<img src=""" & OutPath & "\" & Filename & """ border=0>"
End Sub
Public Sub LinearGraph( _
CaptionName As String, _
Xname As String, _
Yname As String, _
Table As String, _
USER As String, _
xvalue As String, _
yValue As String, _
Optional ShowLinear As Boolean = True)
On Error Resume Next
Dim RS As Recordset, k As Long, dd As Single
If LCase$(USER) = "currentuser" Then USER = CurrentUser.Username

Set RS = DB.OpenRecordset("select * from " & Table & _
  " where user='" & USER & "';", dbOpenDynaset)
  
GS.Reset
GS.CaptionFont.Size = 13
GS.OverSizeX = True
GS.OverSizeY = True
GS.ShowLegend = False
GS.YAxisName = Yname
GS.XAxisName = Xname
xvalue = Trim$(xvalue)
yValue = Trim$(yValue)
Dim DisplayData() As Single, i As Long
While Not RS.EOF
   ReDim Preserve DisplayData(1, k)
   DisplayData(0, k) = RS(Replace(xvalue, "p_", "", , , vbTextCompare))
   DisplayData(1, k) = RS(Replace(yValue, "p_", "", , , vbTextCompare))
   dd = RS("calories")
   If InStr(1, xvalue, "p_", vbTextCompare) <> 0 Then
        If xvalue = "p_fat" Then DisplayData(0, k) = DisplayData(0, k) * 9 / dd * 100
        If xvalue = "p_protein" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
        If xvalue = "p_carbs" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
        If xvalue = "p_sugar" Then DisplayData(0, k) = DisplayData(0, k) * 4 / dd * 100
   End If
   If InStr(1, yValue, "p_", vbTextCompare) <> 0 Then
        If yValue = "p_fat" Then DisplayData(1, k) = DisplayData(1, k) * 9 / dd * 100
        If yValue = "p_protein" Then DisplayData(1, k) = DisplayData(1, k) * 4 / dd * 100
        If yValue = "p_carbs" Then DisplayData(1, k) = DisplayData(1, k) * 4 / dd * 100
        If yValue = "p_sugar" Then DisplayData(1, k) = DisplayData(1, k) * 4 / dd * 100
   End If
   RS.MoveNext
   k = k + 1
Wend

If UBound(DisplayData, 2) < 2 Then
   GS.DrawGraph
   GS.DrawCentered "Not Enough Datapoints." & vbCrLf & "Please record more days in journal."
   GoTo Exitsub
End If

Dim slope As Double, B As Double
Dim T() As Single
If ShowLinear Then
Call Module1.LinearFit(DisplayData, slope, B)
ReDim T(1, UBound(DisplayData, 2))
For i = 0 To UBound(DisplayData, 2)
   T(1, i) = slope * DisplayData(0, i) + B
   T(0, i) = DisplayData(0, i)
Next i
SC.ExecuteStatement "m=" & slope
SC.ExecuteStatement "b=" & B
End If
GS.MonthAxis = False
If ShowLinear Then GS.AddLine "", gLine, vbRed, T
GS.AddLine "", gScatter, 0, DisplayData
GS.DrawGraph
Exitsub:
Dim Filename As String
Filename = "N" & Graphnumber & ".bmp"
Graphnumber = Graphnumber + 1

GS.SaveBitmap OutPath & "\" & Filename
HTMLOutput = HTMLOutput & "<img src=""" & OutPath & "\" & Filename & """ border=0>"
End Sub

Public Function AverageNutrition(DateMethod As String, nDays As Long, Nuts As Collection, NutNames As Collection, Optional MealNumber As Long = -2)
Dim RS As Recordset
Dim cc As Single, i As Long, j As Single

Dim StartDate As Date
On Error Resume Next
If LCase$(DateMethod) = "user" Then
  nDays = 7 * Val(InputBox("Please enter the number of weeks you wish to include in report?", "", nDays))
ElseIf LCase$(DateMethod) = "last" Then
  nDays = LastTime
End If
If nDays <= 0 Then nDays = 1
LastTime = nDays

Dim mDays As Long
If nDays = 1 Then
   Set RS = DB.OpenRecordset("select * from dailylog where user ='" & CurrentUser.Username & "' and date=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
Else
   Set RS = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "' and " _
   & "date >=#" & FixDate(StartDate) & "# and date<=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
End If
While Not RS.EOF
  If RS("calories") <> 0 Then mDays = mDays + 1
  RS.MoveNext
Wend
If mDays = 0 Then mDays = 1

If nDays = 1 Then
 If MealNumber = -2 Then
  Set RS = DB.OpenRecordset("SELECT DaysInfo.*, Abbrev.*, Weight.* " _
    & "FROM (DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID = Abbrev.Index) INNER JOIN Weight ON (DaysInfo.ItemID = Weight.Index) AND (DaysInfo.Unit = Weight.Msre_Desc) " _
    & "where user ='" & CurrentUser.Username & "' and date=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
 Else
  Set RS = DB.OpenRecordset("SELECT DaysInfo.*, Abbrev.*, Weight.* " _
    & "FROM (DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID = Abbrev.Index) INNER JOIN Weight ON (DaysInfo.ItemID = Weight.Index) AND (DaysInfo.Unit = Weight.Msre_Desc) " _
    & "where user ='" & CurrentUser.Username & "' and date=#" & FixDate(DisplayDate) & "# and meal=" & MealNumber & ";", dbOpenDynaset)
 End If
'  nDays = 1
Else
  StartDate = DateAdd("d", -1 * nDays + 1, DisplayDate)
  If MealNumber = -2 Then
    Set RS = DB.OpenRecordset("SELECT DaysInfo.*, Abbrev.*, Weight.* " _
      & "FROM (DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID = Abbrev.Index) INNER JOIN Weight ON (DaysInfo.ItemID = Weight.Index) AND (DaysInfo.Unit = Weight.Msre_Desc) " _
      & "where user ='" & CurrentUser.Username & "' and " _
      & "date>=#" & FixDate(StartDate) & "# and date<=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
  Else
    Set RS = DB.OpenRecordset("SELECT DaysInfo.*, Abbrev.*, Weight.* " _
      & "FROM (DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID = Abbrev.Index) INNER JOIN Weight ON (DaysInfo.ItemID = Weight.Index) AND (DaysInfo.Unit = Weight.Msre_Desc) " _
      & "where user ='" & CurrentUser.Username & "' and " _
      & "date>=#" & FixDate(StartDate) & "# and date<=#" & FixDate(DisplayDate) & "# and meal=" & MealNumber & ";", dbOpenDynaset)
  End If
End If


  
For i = 0 To RS.Fields.Count - 1
   Nuts.Add 0, RS.Fields(i).Name
   NutNames.Add RS.Fields(i).Name
Next i
  
While Not RS.EOF
  cc = Module1.TranslateUnitToGrams(RS("abbrev.index"), RS("unit"))
  cc = cc * RS("servings") / 100
  For i = 0 To RS.Fields.Count - 1
    j = Nuts(RS.Fields(i).Name)
    Nuts.Remove RS.Fields(i).Name
    j = j + cc * Val(RS(i)) / mDays
    Nuts.Add j, RS.Fields(i).Name
  Next i
  RS.MoveNext
Wend
RS.Close
Dim Cals As Single, ex As Single

If nDays = 1 Then
   Set RS = DB.OpenRecordset("select * from dailylog where user ='" & CurrentUser.Username & "' and date=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
   Cals = Val(RS("calories"))
   ex = Val(RS("exercise_cal"))
Else
   Set RS = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "' and " _
   & "date >=#" & FixDate(StartDate) & "# and date<=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
   Cals = 0
   ex = 0
   While Not RS.EOF
     Cals = Cals + Val(RS("calories")) / mDays
     ex = ex + Val(RS("exercise_cal")) / mDays
     RS.MoveNext
   Wend
End If
If MealNumber = -2 Then
  Nuts.Remove "Calories"
  Nuts.Add Cals, "Calories"
End If
Nuts.Add ex, "ExCal"
Nuts.Add Cals - ex, "NetCal"

End Function
Public Function BarNutrition(DateMethod As String, nDays As Long) As String
Dim RS As Recordset
Dim cc As Single, i As Long, j As Single
Dim rs2 As Recordset
Dim Nuts As New Collection, NutNames As New Collection

On Error Resume Next
Call AverageNutrition(DateMethod, nDays, Nuts, NutNames)

Set RS = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If RS.EOF Then
  Set RS = DB.OpenRecordset("select * from ideals where user='anyuser';", dbOpenDynaset)
End If
Set rs2 = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
RS.Edit
RS("calories") = rs2("calories")
RS("fat") = rs2("calories") * rs2("fat") / 9
RS("carbs") = rs2("calories") * rs2("carbs") / 4
RS("protein") = rs2("calories") * rs2("protein") / 4
RS.Update
Nuts.Add Val(rs2("bmr")), "BMR"
On Error Resume Next
Dim NutPercents As New Collection
For i = 0 To RS.Fields.Count - 1
  j = Nuts(RS.Fields(i).Name)
  NutPercents.Add Round(j / RS(i) * 100, 2), RS.Fields(i).Name
Next i
NutPercents.Add Nuts("netcal") / rs2("calories") * 100, "NCalories"

Dim ff As Long, PP As Single, Template As String, junk As String
ff = FreeFile
Open App.path & "\resources\temp\TodaysNutrients.htm" For Input As #ff
While Not EOF(ff)
  Line Input #ff, junk
  Template = Template & junk & vbCrLf
Wend
Close #ff

Template = Replace(Template, "##Calories", Round(Nuts("Calories")), , , vbTextCompare)
Template = Replace(Template, "##Calfat", Round(Nuts("fat") / 9), , , vbTextCompare)
Template = Replace(Template, "##fat", Round(Nuts("fat")), , , vbTextCompare)
Template = Replace(Template, "##sfat", Round(Nuts("saturated fat")), , , vbTextCompare)
Template = Replace(Template, "##mfat", Round(Nuts("monounsaturated fat")), , , vbTextCompare)
Template = Replace(Template, "##pfat", Round(Nuts("Polyunsaturated Fat")), , , vbTextCompare)
Template = Replace(Template, "##cholesterol", Round(Nuts("cholesterol")), , , vbTextCompare)
Template = Replace(Template, "##sodium", Round(Nuts("sodium")), , , vbTextCompare)
Template = Replace(Template, "##carbs", Round(Nuts("carbs")), , , vbTextCompare)
Template = Replace(Template, "##fiber", Round(Nuts("fiber")), , , vbTextCompare)
Template = Replace(Template, "##sugar", Round(Nuts("sugar")), , , vbTextCompare)
Template = Replace(Template, "##protein", Round(Nuts("protein")), , , vbTextCompare)
Template = Replace(Template, "##protein", Round(Nuts("protein")), , , vbTextCompare)
Template = Replace(Template, "##protein", Round(Nuts("protein")), , , vbTextCompare)
Template = Replace(Template, "##bmr", Round(Nuts("bmr")), , , vbTextCompare)
Template = Replace(Template, "##ncals", Round(Nuts("netcal")), , , vbTextCompare)
Template = Replace(Template, "##bcals", Round(Nuts("excal")), , , vbTextCompare)
Template = Replace(Template, "##bmr", Round(Nuts("bmr")), , , vbTextCompare)


Dim O As New Collection
O.Add "NCalories"
O.Add "Fat"
O.Add "Saturated Fat"
O.Add "Monounsaturated fat"
O.Add "Polyunsaturated fat"
O.Add "Sodium"
O.Add "Cholestrol"
O.Add "Carbs"
O.Add "Fiber"
O.Add "Sugar"
O.Add "Protein"

For i = 1 To O.Count
  PP = NutPercents(O(i)) / 2
  If PP > 200 Then PP = 200
  PP = PP * 300 / 100
  junk = "<img height=""10px"" src=""narancscsik.gif"" width=""" & PP & "px"">" & Round(NutPercents(O(i))) & "%"
  Template = Replace(Template, "%%" & O(i), junk, , , vbTextCompare)
Next i

Dim B As New Collection
B.Add "Vitamin A"
B.Add "Vitamin C"
B.Add "Calcium"
B.Add "Iron"

B.Add "Magnesium"
B.Add "Phosphorus"
B.Add "Potassium"

B.Add "Zinc"
B.Add "Copper"
B.Add "Manganese"
B.Add "Selenium"

B.Add "Thiamin"
B.Add "Riboflavin"
B.Add "Niacin"
B.Add "Pantothenic acid"
B.Add "Vitamin B6"
B.Add "Folate"
B.Add "Vitamin B12"

B.Add "Retinol"
B.Add "Vitamin E"
B.Add "Vitamin K"
B.Add "Alpha-carotene"

For i = 1 To B.Count
  PP = NutPercents(B(i)) / 2
  If PP > 200 Then PP = 200
  
  PP = PP * 300 / 100
  junk = "<img height=""10px"" src=""narancscsik.gif"" width=""" & PP & "px"">" & Round(NutPercents(B(i))) & "%"
  Template = Replace(Template, "%%" & B(i), junk, , , vbTextCompare)

Next i
RS.Close
rs2.Close
BarNutrition = Template
HTMLOutput = HTMLOutput & Template

End Function


Public Sub MLinearRegress(Cols() As String, depend As String)


    On Error GoTo Err_Proc
'set up the linear regression
Dim RS As Recordset
Set RS = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "';", dbOpenDynaset)
Dim X() As Single, Y() As Single
Dim dd As Single, i As Long, nut As String, j As Long
While Not RS.EOF
  ReDim Preserve X(UBound(Cols), i)
  ReDim Preserve Y(i)
  dd = RS("calories")
  For j = 0 To UBound(Cols)
    nut = LCase$(Cols(j))
    X(j, i) = RS(Replace(nut, "p_", "", , , vbTextCompare))
    If nut = "p_fat" Then X(j, i) = X(j, i) * 9 / dd * 100
    If nut = "p_protein" Then X(j, i) = X(j, i) * 4 / dd * 100
    If nut = "p_carbs" Then X(j, i) = X(j, i) * 4 / dd * 100
    If nut = "p_sugar" Then X(j, i) = X(j, i) * 4 / dd * 100
  Next j
  i = i + 1
  RS.MoveNext
Wend

Dim Coeff() As Single, yy As Single
ReDim BCorrections(1, 6, UBound(Y))

Coeff = MathMod.LinRegr(UBound(X), UBound(Y), X, Y)
If Coeff(1) <> 0 Then
   For i = 0 To UBound(Y)
       yy = Coeff(0)
       For j = 1 To UBound(Coeff)
         yy = yy + Coeff(j) * X(j, i)
       Next j
       BCorrections(1, 1, i) = (Y(i) - (yy - Coeff(1) * X(1, i) - Coeff(0))) * 7
       BCorrections(0, 1, i) = X(1, i)
       
       For j = 2 To 6
         BCorrections(1, j, i) = Y(i) - (yy - Coeff(j) * X(j, i))
         BCorrections(0, j, i) = X(j, i)
       Next j
    Next i

    
End If

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "MLinearRegress", Err.Description
    Resume Exit_Proc


End Sub

Public Sub LinBMR(m_BMR, t_CPP, A)


    On Error GoTo Err_Proc
Dim RS As Recordset, k As Long, dd As Single
Set RS = DB.OpenRecordset("select * from dailylog where user='small';", dbOpenDynaset)
Dim DisplayData() As Single, i As Long
While Not RS.EOF
   ReDim Preserve DisplayData(1, k)
   DisplayData(0, k) = RS("calories")
   A = A + RS("calories")
   DisplayData(1, k) = RS("weight")
   RS.MoveNext
   k = k + 1
Wend
Set RS = Nothing
If k = 0 Then
   m_BMR = CurrentUser.BMR
   t_CPP = CurrentUser.CalPound
   A = 0
   Exit Sub
End If
A = A / k
Dim slope As Double, B As Double
Dim T() As Single
Call Module1.LinearFit(DisplayData, slope, B)


    t_CPP = Round(1 / slope, 2)
    If t_CPP > 2000 And t_CPP < 4500 And slope <> 0 Then
        m_BMR = Round(Abs(B / slope), 2)
    Else
        t_CPP = 3500
        m_BMR = 0
        For i = 0 To UBound(DisplayData, 2)
           m_BMR = m_BMR - DisplayData(1, i) * 500 + DisplayData(0, i)
        Next i
        If i > 1 Then
            m_BMR = m_BMR / (i)
        Else
            m_BMR = CurrentUser.BMR
        End If
        If m_BMR < 1000 Or m_BMR > 4000 Then
           m_BMR = CurrentUser.BMR
        End If
    End If

  SC.ExecuteStatement "BMR = " & m_BMR
  SC.ExecuteStatement "Cpp = " & t_CPP
  SC.ExecuteStatement "average=" & A
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "LinBMR", Err.Description
    Resume Exit_Proc


End Sub

Public Sub TimeGraph(CaptionName As String, Yname As String, Ynut As String)
On Error Resume Next
'set up the graph
GS.Reset
GS.Cls
GS.XAxisName = ""
GS.CaptionFont.Size = 13
GS.YAxisName = Yname
GS.CaptionName = CaptionName
Dim temp As Recordset, cc As Long, Y As Long, O As Date
Dim d As Date, m As Long, Y2 As Long, O2 As Date
Dim dYear As Long, NM As Date, Ws, W
Dim T()
Set temp = DB.OpenRecordset("select * from dailylog where user = '" _
    & CurrentUser.Username & "'  order by date;", dbOpenDynaset)
    
If temp.EOF Then Exit Sub
cc = 0
Y = Year(temp("date"))
O = Y & "-01-01"
While Not temp.EOF
     DoEvents
     d = temp("date")
     m = Month(d)
     Y2 = Year(d)
     O2 = DateHandler.IsoDateString(CInt(m), 1, CInt(Y2))
     dYear = (Y2 - Y) * 12
     NM = DateAdd("m", 1, O2)
     Ws = Int(DateDiff("d", O2, NM))
     W = Int(DateDiff("d", O2, d))
     
     If temp(Trim$(Ynut)) <> 0 Then
         ReDim Preserve T(1, cc)
         T(0, cc) = m + W / Ws + dYear
         T(1, cc) = temp(Trim$(Ynut))
         cc = cc + 1
     End If
     temp.MoveNext
Wend
cc = cc - 1
GS.AutoScaleY = True
GS.Reset
If cc > 1 Then
   GS.AddLine "", gLine, vbRed, T
   GS.DrawGraph
Else
   GS.DrawGraph
   GS.DrawCentered "Not Enough Datapoints." & vbCrLf & "Please record more days in journal."
End If
GS.ShowLegend = False
Dim Filename As String
Filename = "T" & Graphnumber & ".bmp"
Graphnumber = Graphnumber + 1
GS.SaveBitmap OutPath & "\" & Filename
HTMLOutput = HTMLOutput & "<img src=""" & OutPath & "\" & Filename & """ border=0>"
End Sub

Private Sub ConvertASP()


    On Error GoTo Err_Proc
    CodeSplitCount = 0
    Dim lines() As String, i As Long
    lines = Split(Template, vbCrLf)
    codeString = ""
    For i = 0 To UBound(lines)
      codeString = codeString & Trim$(lines(i)) & vbCrLf
    Next i
    
    
    codeString = Replace(codeString, "response.write", "response.htmlwrite", , , vbTextCompare)
    codeString = Replace(codeString, "<%=", "@~@~<%response.htmlwrite ")
    codeString = Replace(codeString, "<% =", "@~@~<%response.htmlwrite ")
    codeString = Replace(codeString, "%>" & vbCrLf, "%>@@~~@@")
    codeString = Replace(codeString, "%>", "<%")
    
    If InStr(1, codeString, "<%=") <> 0 Then
       codeString = Replace(codeString, "<%=", "@~@~<%" & (Chr(34) & " " & Chr(38) & " "))
    End If
    If Left(codeString, 2) <> "<%" Then
        CodeSplit() = Split(vbCrLf & codeString, "<%")
    Else
        CodeSplit() = Split(codeString, "<%")
    End If
    If UBound(CodeSplit()) > 0 Then
        Do Until CodeSplitCount >= UBound(CodeSplit()) + 1
            ParseHTML
            ParseASP
        Loop
    Else
        ParseHTML
    End If
    Template = Join(CodeSplit, "")
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "ConvertASP", Err.Description
    Resume Exit_Proc


End Sub
Private Sub ParseHTML()


    On Error GoTo Err_Proc
    HTMLSplitCount = 0
    HTMLString() = Split(CodeSplit(CodeSplitCount), vbCrLf)
    Do Until HTMLSplitCount > UBound(HTMLString())
        HTMLString(HTMLSplitCount) = Replace(HTMLString(HTMLSplitCount), Chr(34), Chr(34) & Chr(34))
        If EmptyString(HTMLString(HTMLSplitCount)) = False Then
           If InStr(1, HTMLString(HTMLSplitCount), "@~@~", vbTextCompare) <> 0 Then
              HTMLString(HTMLSplitCount) = "response.htmlline " & Chr(34) & HTMLString(HTMLSplitCount) & Chr(34)
              HTMLString(HTMLSplitCount) = Replace(HTMLString(HTMLSplitCount), "@~@~", "")
           ElseIf Right$(HTMLString(HTMLSplitCount), 6) = "@@~~@@" Then
              HTMLString(HTMLSplitCount) = "response.htmlline " & Chr(34) & HTMLString(HTMLSplitCount) & Chr(34)
           Else
              HTMLString(HTMLSplitCount) = "response.htmlline " & Chr(34) & HTMLString(HTMLSplitCount) & "@@~~@@" & Chr(34)
           End If
        End If
        HTMLSplitCount = HTMLSplitCount + 1
    Loop
    CodeSplit(CodeSplitCount) = Join(HTMLString, vbCrLf) & vbCrLf
    CodeSplitCount = CodeSplitCount + 1
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "ParseHTML", Err.Description
    Resume Exit_Proc


End Sub
Private Sub ParseASP()


    On Error GoTo Err_Proc
    ASPSplitCount = 0
    If (CodeSplitCount < UBound(CodeSplit()) + 1) Then
        ASPString() = Split(CodeSplit(CodeSplitCount), vbCrLf)
        ASPSplitCount = 0
        Do Until ASPSplitCount >= UBound(ASPString())
            'additional code can go here
            ASPSplitCount = ASPSplitCount + 1
        Loop
        CodeSplit(CodeSplitCount) = Join(ASPString, vbCrLf) & vbCrLf
        
        CodeSplitCount = CodeSplitCount + 1
    End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "ASP", "ParseASP", Err.Description
    Resume Exit_Proc


End Sub
Private Function EmptyString(TestString As String) As Boolean


    On Error GoTo Err_Proc
    If Trim(Replace(TestString, vbTab, "")) = vbCrLf Then
        EmptyString = True
    ElseIf Trim(Replace(TestString, vbTab, "")) = "" Then
        EmptyString = True
    End If
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "ASP", "EmptyString", Err.Description
    Resume Exit_Proc


End Function
Public Function Init(mScriptControl As ScriptControl, mGS As uGraphSurface)


    On Error GoTo Err_Proc
Set SC = mScriptControl
Set GS = mGS

SC.AddObject "Response", Me, True
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "ASP", "Init", Err.Description
    Resume Exit_Proc


End Function

Public Function LoadTemplate(mOutPath As String, Filename As String)
On Error GoTo errhandl2

Graphnumber = 0
OutPath = mOutPath
HTMLOutput = ""
Template = ""
Dim ff As Long, junk As String
ff = FreeFile
Open Filename For Input As #ff
While Not EOF(ff)
  Line Input #ff, junk
  Template = Template & junk & vbCrLf
Wend
Close #ff

Call ConvertASP
On Error GoTo errhandl
'Open App.path & "\log.vb" For Output As #ff
'Print #ff, template
'Close #ff
If DoDebug Then Clipboard.SetText Template

SC.AddCode Template
'SC.ExecuteStatement template
ff = FreeFile
LoadTemplate = HTMLOutput
Exit Function
errhandl:
On Error Resume Next
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & SC.Error.Description
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & Err.Description
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & Err.Number
HTMLOutput = HTMLOutput & vbCrLf & "<br>Line and Character= " & SC.Error.Line & ", " & SC.Error.Column
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & SC.Error.Text
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & Replace(Replace(Replace(Replace(Replace(Template, "<br>", vbCrLf, , , vbTextCompare), "<", "&lt;"), ">", "&gt;"), vbCrLf, "<br>"), "@@~~@@", "")
LoadTemplate = HTMLOutput
Exit Function
errhandl2:
HTMLOutput = HTMLOutput & vbCrLf & "<br>" & Err.Description
LoadTemplate = HTMLOutput
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
