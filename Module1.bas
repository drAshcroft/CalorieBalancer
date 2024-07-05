Attribute VB_Name = "Module1"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.

Option Explicit
Public Declare Function htmlHelpTopic Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As String) As Long
    
'this turns off the help files when it is set to true
Global Const DoDebug = False
Global Const NoTrial = False
Global Const FreeVersion = False
Global Const TrialDays = 90000000
Global Const RegNow = False
Global Const Version = "408"
Global Const RestartTrial = False
Global StartingProgram As Boolean
Global CloseProgram As Boolean

Global FirstRun As Boolean
Global NoQuestions As Boolean
Global HelpWindowHandle As Long
Global HelpPath As String
Global Branding As New Collection
Global frmMain As frmMainO

Public Paid As Boolean


Public Type Users
    Username As String
    Weight As Single
    Height As Single
    BMR As Single
    CalPound As Single
    Password As String
    HashNumber As String
End Type
Private Type DayInfo
     Fdate As Date
     Calories As Single
     Exercise_Cal As Single
     Weight As Single
     BMR As Single
     bfp As Single
     Valid As Boolean
     sugar As Single
     fat As Single
     Protein As Single
     fiber As Single
     carbs As Single
     DisComfort As Single
End Type



Global DisplayDate As Date
Public CurrentUser As Users
Public DB As Database
Public WatchHeads() As String
Public FirstChanged As Date
Global Nutmaxes As Calories
Global Today As Date


Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_CLOSE_ALL = &H12
Public Const HH_DISPLAY_TOPIC = &H0

  Const aCup = 0
  Const aTBSP = 1
  Const aTSP = 2
  Const aMili = 3
  Const aCInch = 5
  Const aFLOZ = 4
  Const aQuart = 6
  'weights
  Const aPound = 0
  Const aOZ = 1
  Dim Converts(6, 6) As Single
  Dim ConvertsW(1, 1) As Single
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage
Private Function SyncAbbrev(FoodGroup As Long) As String
On Error GoTo errhandl
            Dim junk As String, Abbrev As Recordset, i As Long
            Dim weights As Recordset, junk2 As String
            
            junk = ""
            Set Abbrev = DB.OpenRecordset("select * from abbrev where foodgroup='" & Right("0000" & FoodGroup, 4) & "' order by usage;", dbOpenDynaset)
            i = 0
            While (Not Abbrev.EOF) And i < 10
               junk = junk & vbTab & vbTab & Abbrev("foodname") & "|||"
               junk = junk & 300000 + Abbrev("index")
               If IsNull(Abbrev("calories")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("calories")
               End If
               If IsNull(Abbrev("fat")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("fat")
               End If
               If IsNull(Abbrev("carbs")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("carbs")
               End If
               If IsNull(Abbrev("protein")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("protein")
               End If
               If IsNull(Abbrev("sugar")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("sugar")
               End If
               If IsNull(Abbrev("fiber")) Then
                  junk = junk & "|||0"
               Else
                  junk = junk & "|||" & Abbrev("fiber")
               End If
               'Dim weights As Recordset, junk2 As String
               junk2 = ""
               Set weights = DB.OpenRecordset("select * from weight where index=" & Abbrev("index") & ";", dbOpenDynaset)
               While Not weights.EOF
                  junk2 = junk2 & weights("msre_desc") & "@@" & weights("gm_wgt") / weights("amount") & "~~"
                  weights.MoveNext
               Wend
               If Len(junk2) > 1 Then junk2 = Left(junk2, Len(junk2) - 1)
               junk = junk & "|||" & junk2
               junk = junk & vbCrLf
               i = i + 1
               Abbrev.MoveNext
            Wend
            SyncAbbrev = junk
Exit Function
errhandl:
If DoDebug Then
   Call MsgBox(Err.Description)
   Resume Next
End If
End Function

Public Sub SyncWithEve()
  On Error GoTo errhandl
    Dim FoodgroupS  As Recordset
    Dim Second As Recordset
    Dim Abbrev As Recordset, weights As Recordset, junk2 As String
    Dim i As Long
    Dim junk As String
    Set FoodgroupS = DB.OpenRecordset("select * from foodgroups where parentnumber=-10;", dbOpenDynaset)
    While Not FoodgroupS.EOF
       junk = junk & FoodgroupS("category") & "|||M" & FoodgroupS("catnumber") & vbCrLf
       Set Second = DB.OpenRecordset("select * from foodgroups where parentnumber =" & FoodgroupS("catnumber") & ";", dbOpenDynaset)
       If Not Second.EOF Then
          
          While Not Second.EOF
            junk = junk & vbTab & Second("category") & "|||I" & Second("catnumber") & vbCrLf
            junk = junk & SyncAbbrev(Second("catnumber"))
            Second.MoveNext
          Wend
       Else
          junk = junk & SyncAbbrev(FoodgroupS("catnumber"))
       End If
       FoodgroupS.MoveNext
    Wend
    Debug.Print junk
    Clipboard.Clear
    Clipboard.SetText junk
    Exit Sub
errhandl:
    If DoDebug Then
       Call MsgBox(Err.Description)
       Resume Next
     End If
End Sub

Public Sub InitASP(SC As ScriptControl, GS As uGraphSurface, mASP As ASP)
On Error GoTo errhandl
SC.Reset

Dim RS As Recordset, i As Long, Logs As New Collection
Dim Prof As New Collection, ideals As New Collection

Set RS = DB.OpenRecordset("select * from dailylog where user='" & CurrentUser.Username & "' " _
  & "and date=#" & FixDate(DisplayDate) & "#;", dbOpenDynaset)
SC.ExecuteStatement "comments="" "" "
If Not RS.EOF Then
    For i = 0 To RS.Fields.Count - 1
      If LCase$(RS(i).Name) <> "date" Then
        If (Not IsNull(RS(i))) Then
           Logs.Add Trim$(RS(i)), Trim$(RS(i).Name)
        Else
           Logs.Add " ", Trim$(RS(i).Name)
        End If
      End If
    Next i
    RS.Close
Else
    Set RS = DB.OpenRecordset("select * from dailylog;", dbOpenDynaset)
    For i = 0 To RS.Fields.Count - 1
           Logs.Add " ", Trim$(RS(i).Name)
    Next i
    RS.Close

End If
Set RS = DB.OpenRecordset("select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If Not RS.EOF Then
   For i = 0 To RS.Fields.Count - 1
     If Not IsNull(RS(i)) Then
       ideals.Add Trim$(RS(i)), Trim$(RS(i).Name)
     End If
   Next i
End If
Set RS = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
If Not RS.EOF Then
  For i = 0 To RS.Fields.Count - 1
    If Not IsNull(RS(i)) Then
     Prof.Add Trim$(RS(i)), Trim$(RS(i).Name)
    Else
     Prof.Add "", Trim$(RS(i).Name)
    End If
  Next i
End If
'Dim Abbrev As New exCollection
Dim Units As New Collection
Set RS = DB.OpenRecordset("select * from units;", dbOpenDynaset)
For i = 0 To RS.Fields.Count - 1
   Units.Add Trim$(RS.Fields(i).Value & " "), Trim$(RS.Fields(i).Name & " ")
Next i
RS.Close
Set RS = Nothing
Set RS = DB.OpenRecordset("select * from abbrev;", dbOpenDynaset)

SC.AddObject "Ideals", ideals, True
SC.AddObject "Profile", Prof, True
SC.AddObject "Log", Logs, True
SC.AddObject "Abbrev", RS, True
SC.AddObject "Units", Units, True
SC.ExecuteStatement "Log_Date=""" & DisplayDate & """"


'rs.Fields(i).name
Call mASP.Init(SC, GS)
errhandl:
End Sub

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
On Error GoTo errhandl
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
errhandl:
End Function
Private Sub loadConverts(ByRef Converts() As Single, ByRef ConvertsW() As Single)
On Error Resume Next
  ConvertsW(aPound, aPound) = 1
  ConvertsW(aPound, aOZ) = 16
  ConvertsW(aOZ, aOZ) = 1
  ConvertsW(aOZ, aPound) = 1 / 16
  
  Converts(aCup, aTBSP) = 16
  Converts(aCup, aTSP) = 48
  Converts(aCup, aMili) = 236.588238
  Converts(aCup, aCInch) = 14.4375001
  Converts(aCup, aFLOZ) = 8
  Converts(aCup, aQuart) = 0.25
  
  Converts(aTBSP, aTSP) = 3
  Converts(aTBSP, aMili) = 14.7867648
  Converts(aTBSP, aCInch) = 0.902343754
  Converts(aTBSP, aFLOZ) = 0.5
  Converts(aTBSP, aQuart) = 0.015625
  
  Converts(aTSP, aMili) = 4.92892161
  Converts(aTSP, aCInch) = 0.300781251
  Converts(aTSP, aFLOZ) = 0.166666667
  Converts(aTSP, aQuart) = 0.00520833333
  
  Converts(aMili, aCInch) = 0.0610237441
  Converts(aMili, aFLOZ) = 0.0338140226
  Converts(aMili, aQuart) = 0.0010566882
  
  Converts(aCInch, aFLOZ) = 0.554112552
  Converts(aCInch, aQuart) = 0.0173160172
  
  Converts(aFLOZ, aQuart) = 0.03125
  Dim i As Long, j As Long
  For i = 0 To UBound(Converts, 1)
    For j = 0 To UBound(Converts, 2)
       If i = j Then
          Converts(i, j) = 1
       ElseIf Converts(i, j) = 0 Then
          If Converts(j, i) <> 0 Then Converts(i, j) = 1 / Converts(j, i)
       End If
    Next j
  Next i
  

End Sub

Private Sub ConvertUnitsEngine(sItems() As String, ItemGrams() As Single, OutNames As Collection, OutGrams As Collection)
On Error GoTo errhandl
  'Dim sListItems() As String
  Set OutNames = New Collection
  Set OutGrams = New Collection
  
  Dim junk As String, Vs(6) As Boolean, Ws(1) As Boolean
  Dim i As Long, j As Long, k As Long
  Dim nVs(6) As Boolean, nWs(1) As Boolean
  Dim NamesV(6) As String, NamesW(1) As String
  Dim hasVs As Boolean, hasWs As Boolean
  Dim CNV As New Collection, RVV As New Collection
  Dim CNW As New Collection, RVW As New Collection
  
  NamesV(aCup) = "cup"
  NamesV(aTBSP) = "Tbsp"
  NamesV(aTSP) = "tsp"
  NamesV(aMili) = "ml"
  NamesV(aCInch) = "cubic inch"
  NamesV(aFLOZ) = "fl oz"
  NamesV(aQuart) = "quart"
  NamesW(aPound) = "lbs"
  NamesW(aOZ) = "oz"
  
  CNV.Add "cup": RVV.Add aCup, "cup"
  CNV.Add "tablespoon": RVV.Add aTBSP, "tablespoon"
  CNV.Add "tsp": RVV.Add aTSP, "tsp"
  CNV.Add "cubic inch": RVV.Add aCInch, "cubic inch"
  CNV.Add "tbsp": RVV.Add aTBSP, "tbsp"
  CNV.Add "fl oz": RVV.Add aFLOZ, "fl oz"
  CNV.Add "fluid ounce": RVV.Add aFLOZ, "fluid ounce"
  CNV.Add "fl. oz.": RVV.Add aFLOZ, "fl. oz."
  CNV.Add "quart": RVV.Add aQuart, "quart"
  CNV.Add "ml": RVV.Add aMili, "ml"

  CNW.Add "lbs": RVW.Add aPound, "lbs"
  CNW.Add "oz": RVW.Add aOZ, "oz"
  CNW.Add "ounce": RVW.Add aOZ, "ounce"
  CNW.Add "pound": RVW.Add aPound, "pound"
  CNW.Add "lb": RVW.Add aPound, "lb"
  
  Dim VFound As Boolean, ji As Long, junk2 As String
  'check what is already available
  'loop through the list and then check against all the units abov
  'if a volume unit is not found then check for weight units
  'this is done to make sure that fluid ounces are not considered ounces
  For i = 0 To UBound(sItems)
    junk = LCase$(sItems(i))
    VFound = False
    For j = 1 To CNV.Count
      junk2 = CNV(j)
      ji = InStr(1, junk, junk2, vbTextCompare)
      If ji <> 0 Then
         If ji + Len(junk2) - 1 = Len(junk) Then
            VFound = True
            Vs(RVV(junk2)) = True
         Else
            If InStr(1, junk, junk2 & " ", vbTextCompare) <> 0 Or InStr(1, junk, junk2 & ",", vbTextCompare) <> 0 Then
              VFound = True
              Vs(RVV(junk2)) = True
            End If
         End If
      End If
    Next j
    If Not VFound Then
      For j = 1 To CNW.Count
        junk2 = CNW(j)
        ji = InStr(1, junk, junk2, vbTextCompare)
        If ji <> 0 Then
          If ji + Len(junk2) - 1 = Len(junk) Then
            Ws(RVW(junk2)) = True
          Else
            If InStr(1, junk, junk2 & " ", vbTextCompare) <> 0 Or InStr(1, junk, junk2 & ",", vbTextCompare) <> 0 Then
              Ws(RVW(junk2)) = True
            End If
          End If
        End If
      Next j
    End If
  Next i
  
 
  'now indicate which units need to be made
  For i = 0 To UBound(Vs)
    If Vs(i) = True Then
      hasVs = True
      Exit For
    End If
  Next i
  If hasVs Then
     For j = 0 To UBound(Vs)
        If Not Vs(j) Then nVs(j) = True
     Next j
  End If
  For i = 0 To UBound(Ws)
    If Ws(i) = True Then
      hasWs = True
      Exit For
    End If
  Next i
  If hasWs Then
     For j = 0 To UBound(Ws)
        If Not Ws(j) Then nWs(j) = True
     Next j
  End If
  'take off the useless units from the output list
  nVs(aQuart) = False
  nVs(aCInch) = False
  'nWs(aOZ) = True
  'nVs(aFLOZ) = False
  
  'now do the conversion for all the needed units
  Dim SS As String, FIndex As Long, Ffound As Boolean
  For i = 0 To UBound(sItems)
    junk = LCase$(sItems(i))
    SS = ""
    VFound = False
    'volume
    For j = 1 To CNV.Count
      junk2 = CNV(j)
      ji = InStr(1, junk, junk2, vbTextCompare)
      Ffound = False
      If ji <> 0 Then
         If ji + Len(junk2) - 1 = Len(junk) Then
           Ffound = True
         ElseIf InStr(1, junk, junk2 & " ", vbTextCompare) <> 0 Or InStr(1, junk, junk2 & ",", vbTextCompare) <> 0 Then
           Ffound = True
         End If
         If Ffound Then
            FIndex = RVV(junk2)
            SS = NamesV(FIndex)
            VFound = True
         End If
      End If
    Next j
    If SS <> "" Then
       Call AddList(FIndex, Converts, i, nVs, ItemGrams, sItems, SS, NamesV, OutNames, OutGrams)
    End If
    If Not VFound Then
       For j = 1 To CNW.Count
         junk2 = CNW(j)
         ji = InStr(1, junk, junk2, vbTextCompare)
         Ffound = False
         If ji <> 0 Then
           If ji + Len(junk2) - 1 = Len(junk) Then
             Ffound = True
           ElseIf InStr(1, junk, junk2 & " ", vbTextCompare) <> 0 Or InStr(1, junk, junk2 & ",", vbTextCompare) <> 0 Then
             Ffound = True
           End If
           If Ffound Then
             FIndex = RVW(junk2)
             SS = NamesW(FIndex)
           End If
         End If
         If SS <> "" Then
           Call AddList(FIndex, ConvertsW, i, nWs, ItemGrams, sItems, SS, NamesW, OutNames, OutGrams)
         End If
       Next j
    End If
  Next i
  

  Exit Sub
errhandl:

Resume Next

End Sub
Public Sub ConvertUnits(lb As ListBox, ByRef Conv() As Single, ByRef cc As Long)
On Error GoTo errhandl
  Dim s() As String, i As Long
  Dim OutNames As Collection, OutGrams As Collection
  ReDim s(lb.ListCount - 1)
  For i = 0 To lb.ListCount - 1
     s(i) = lb.List(i)
  Next i
  Call ConvertUnitsEngine(s, Conv, OutNames, OutGrams)
  lb.AddItem "--Conversions--"
  Conv(cc) = 0
  cc = cc + 1
  ReDim Preserve Conv(cc)
  For i = 1 To OutNames.Count
     lb.AddItem OutNames(i)
     Conv(cc) = OutGrams(OutNames(i))
     cc = cc + 1
     ReDim Preserve Conv(cc)
  Next i
errhandl:

End Sub

Private Sub AddList(FromI As Long, Converts() As Single, i As Long, Valids() As Boolean, BaseGrams() As Single, baseNames() As String, SS As String, ValidNames() As String, OutNames As Collection, OutGrams As Collection)
On Error Resume Next
       Dim base As String
       base = baseNames(i)
       Dim j As Long, junk As String, k As Long
       For j = 0 To UBound(Valids)
         If Valids(j) = True Then
           'LB.AddItem Replace(base, ss, ValidNames(j), , , vbTextCompare)
           'Conv(cc) = Conv(i) * Converts(j, FromI)

           junk = Replace(base, SS, ValidNames(j), , , vbTextCompare)
           For k = 0 To UBound(baseNames)
              If LCase$(junk) = LCase$(baseNames(k)) Then GoTo errOut
           Next k
           For k = 1 To OutNames.Count
              If LCase$(junk) = LCase$(OutNames(k)) Then GoTo errOut
           Next k
           On Error GoTo errOut
          
           OutGrams.Add BaseGrams(i) * Converts(j, FromI), junk
           OutNames.Add junk
errOut:
           Err.Clear
         End If
       Next j
End Sub

Public Function TranslateUnitToGrams(FoodID As Long, Unit As String) As Single
      On Error Resume Next
      Dim temp As Recordset
      Set temp = DB.OpenRecordset("SELECT *" _
                              & " From weight " _
                              & " WHERE ((index=" & FoodID & ") and " _
                              & "(msre_desc = '" & Unit & "'));", dbOpenDynaset)
       
      If temp.EOF Then
         temp.Close
         Set temp = Nothing
         If LCase$(Unit) = "grams" Then
           TranslateUnitToGrams = 1
           Exit Function
         End If
         If Unit = "OZ." Then
           TranslateUnitToGrams = 28.34
           Exit Function
         End If
         Set temp = DB.OpenRecordset("select * from weight where index=" & FoodID & ";", dbOpenDynaset)
         Dim Conv() As Single, sItems() As String, UnitNames As Collection, UnitGrams As Collection
         Dim i As Long
         i = -1
         While Not temp.EOF
            i = i + 1
            ReDim Preserve sItems(i)
            ReDim Preserve Conv(i)
            sItems(i) = temp("msre_desc")
            Conv(i) = temp("gm_wgt") / temp("amount")
            temp.MoveNext
         Wend
         Call ConvertUnitsEngine(sItems, Conv, UnitNames, UnitGrams)
         TranslateUnitToGrams = 0
         TranslateUnitToGrams = UnitGrams(Unit)
         
         Set temp = Nothing
         If TranslateUnitToGrams = 0 Then GoTo lasttry
         Exit Function
      End If
      
      
      TranslateUnitToGrams = temp.Fields("gm_wgt").Value / temp.Fields("amount").Value
      Set temp = Nothing
      Exit Function
lasttry:

      Set temp = DB.OpenRecordset("SELECT *" _
                              & " From weight " _
                              & " WHERE ((index=" & FoodID & ") and " _
                              & "(msre_desc like '" & Left$(Unit, 3) & "*'));", dbOpenDynaset)
       
      If Not temp.EOF Then
         TranslateUnitToGrams = temp.Fields("gm_wgt").Value / temp.Fields("amount").Value
         Err.Clear
         
      End If
      
      temp.Close
      Set temp = Nothing
         
         
End Function



Public Sub FigurePercentages(Balance As PieChart, Calories As Single, fat As Single, sugar As Single, carbs As Single, Protien As Single, fiber As Single, _
 Optional f As Single, Optional c As Single, Optional p As Single, Optional s As Single)
 Dim RS As Recordset
 On Error Resume Next
 If Calories <> 0 Then
   f = fat * 9 / Calories * 100
   c = (carbs - fiber - sugar) * 4 / Calories * 100
   s = sugar * 4 / Calories * 100
   p = Protien * 4 / Calories * 100
   
   If f + c + s + p < 100 Then
     c = 100 - (f + s + p)
   End If
Else
   Set RS = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
   f = RS("fat") * 100
   c = RS("carbs") * 100
   s = RS("sugar") * 100
   p = RS("protein") * 100
   
   RS.Close
   Set RS = Nothing
End If
   
   Balance.Reset
   If f <> 0 Then Balance.AddSlice f, "Fat " & Round(f) & "%", vbRed
   If c <> 0 Then Balance.AddSlice c, "Starch " & Round(c) & "%", vbBlue
   If s <> 0 Then Balance.AddSlice s, "Sugar " & Round(s) & "%", RGB(100, 100, 255) 'RGB(255, 255, 0)
   If p <> 0 Then Balance.AddSlice p, "Protein " & Round(p) & "%", vbGreen
   Balance.DrawGraph

End Sub

Public Function DoCalories(ByVal sex, age, Height, Weight, bodyfat)
Dim schorate
On Error Resume Next
schorate = 0
sex = LCase(sex)
If sex = "female" Then
    schorate = 1.1 * (4.12 * Weight * 1.2 * (1 - bodyfat / 100) + 659)
    If age < 60 Then
         schorate = 1.1 * (6.72 * Weight * 1.2 * (1 - bodyfat / 100) + 487)
         If age > 29 Then schorate = 1.1 * (3.69 * Weight * 1.2 * (1 - bodyfat / 100) + 846)
         If age < 18 Then schorate = 1.1 * (6.07 * Weight * 1.2 * (1 - bodyfat / 100) + 693)
    End If
End If
If sex = "male" Then
    schorate = 1.1 * (5.31 * Weight * 1.15 * (1 - bodyfat / 100) + 588)
    If age < 60 Then
         schorate = 1.1 * (6.83 * Weight * 1.15 * (1 - bodyfat / 100) + 692)
         If age > 29 Then schorate = 1.1 * (5.2 * Weight * 1.15 * (1 - bodyfat / 100) + 873)
         If age < 18 Then schorate = 1.1 * (8.02 * Weight * 1.15 * (1 - bodyfat / 100) + 658)
    End If
End If
Dim baserate
baserate = 0
If sex = "female" Then baserate = 655 + 4.35 * Weight * 1.2 * (1 - bodyfat / 100) + 4.7 * Height - 4.7 * age
If sex = "male" Then baserate = 66 + 6.23 * Weight * 1.15 * (1 - bodyfat / 100) + 12.7 * Height - 6.8 * age
If sex = "female" Then baserate = (9.99 * Weight) + (6.25 * Height) - (4.92 * age) _
  + 166 * 0 - 161

If sex = "male" Then baserate = (9.99 * Weight) + (6.25 * Height) - (4.92 * age) _
  + 166 * 1 - 161


DoCalories = Round(baserate, 1) 'Round((schorate + baserate) / 2)
    
    
End Function
Public Function DoBFP(bodybuild, BMI, BFPMeasure, sex, waist, neck, Height, hips, Weight, wrist, forearm)
   On Error Resume Next
    Dim bfp As Single
    If BFPMeasure = "Est" Then
     If LCase$(sex) = "male" Then
        bfp = Round((Weight - Height ^ 2 * (22 + bodybuild) / 702) / Weight * 100, 1)
     Else
        bfp = Round((Weight - Height ^ 2 * (20 + bodybuild) / 702) / Weight * 100, 1)
     End If
    ElseIf BFPMeasure = "Cloth" Then
      If LCase$(sex) = "male" Then
         bfp = Round(86.01 * Log10(waist - neck) - 70.041 * Log10(Height) + 36.76, 1)
      Else
         bfp = Round(-71.938 + 105.42 * Log10(Weight) + 0.4396 * hips - 0.5086 * wrist - 3.997 * forearm - 1.3085 * Height - 1.354 * neck, 1)
      End If
    End If
    DoBFP = bfp
End Function

Public Function SelectedList(L As ListBox) As Long
On Error Resume Next
  Dim i As Long
  For i = 0 To L.ListCount - 1
     If L.Selected(i) Then
        SelectedList = i
        Exit Function
     End If
  Next i
  SelectedList = -1
End Function

Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
  ' On Error
End Function
Private Sub DeleteOldData()
On Error Resume Next
   Dim O As Date
   Dim d As Date
   Dim i As Long
   Dim W As Long
   Dim temp As Recordset
   d = Today
   O = Year(d) & "-01-01"   '"1/1/" & Year(d)
   i = DateDiff("ww", O, d, vbSunday, vbFirstJan1)
   d = DateAdd("ww", i, O)
   d = DateAdd("m", -4, d)
   W = Weekday(d, vbSunday)
   While W > 1
     d = DateAdd("d", 1, d)
     W = Weekday(d, vbSunday)
   Wend
   
   Set temp = DB.OpenRecordset("SELECT [DailyLog].[date], DailyLog.*, [DailyLog].[User] " & _
                               "From DailyLog " & _
                               "WHERE ((([DailyLog].[date])<#" & FixDate(d) & "#) " & _
                               "And (([DailyLog].[User])='" & CurrentUser.Username & "'));", dbOpenDynaset)

   While Not temp.EOF
      temp.Delete
      temp.MoveNext
   Wend
   
   temp.Close
   Set temp = Nothing
   
   Set temp = DB.OpenRecordset("SELECT * " & _
                               "From ExerciseLog " & _
                               "WHERE ((Week<#" & FixDate(d) & "#) " & _
                               "And (User='" & CurrentUser.Username & "'));", dbOpenDynaset)
   While Not temp.EOF
     temp.Delete
     temp.MoveNext
   Wend
   temp.Close
   Set temp = Nothing
End Sub

Public Sub EndIt(Cancel As Integer)
   On Error Resume Next
   Dim RS As Recordset, ID As Long
   Set RS = DB.OpenRecordset("select * from mealplanner where mealname='blank';", dbOpenDynaset)
   ID = RS("Mealid")
   RS.Delete
 
   RS.Close
   If ID <> 0 Then
   Set RS = DB.OpenRecordset("select * from mealdefinition where mealid=" & ID & ";", dbOpenDynaset)
   While Not RS.EOF
      RS.Delete
      RS.MoveNext
   Wend
   RS.Close
   End If
   Dim temp As Recordset
   Set temp = DB.OpenRecordset("Select lastweek from profiles where user = '" & CurrentUser.Username & "';", dbOpenDynaset)
   If Abs(DateDiff("d", temp("lastweek"), Today)) >= 3 And frmMain.mnuRemind.Checked Then
      Dim ret As VbMsgBoxResult
      ret = MsgBox("Please enter your weight and journal information", vbYesNoCancel, "")
      If ret = vbYes Then
       ' Call frmJournal.ChangeMode(1)
        frmJournal.Show vbModal, frmMain

      End If
      If ret = vbCancel Then
         Cancel = 1
        Exit Sub
      End If
   End If
   Set temp = Nothing
    Cancel = 0
    
   
    Call SaveDay(DisplayDate)
  
    
  
    Call DeleteOldData
   
  

End Sub

Sub Main()
   
On Error Resume Next
   Today = FixDate(Date)
   'msgbox "1"
  

   ' InitCommonControlsVB
   Dim Installdate As Date
   Dim RS As Recordset, AlreadyLoaded As Boolean
   
   
   If InStr(1, Interaction.Command$, "exit", vbTextCompare) <> 0 Then
      Installdate = GetSetting(App.Title, "Settings", "FD", "200-01-01")
      Call OpenURL("http://www.caloriebalancediet.com/feedback.asp?Version=CalorieBalance" & Version & "&InstalledDays=" & FixDate(Installdate))
      End
   End If
   'msgbox "2"
   Dim DoUpdateScript As Boolean
   DoUpdateScript = False
   If InStr(1, Interaction.Command$, ".cbm", vbTextCompare) <> 0 Then
      
      
      Dim DBOut As Database, OtherProfiles As Boolean
      Set DBOut = OpenDatabase(Interaction.Command$)
      Set RS = DBOut.OpenRecordset("select * from profiles;", dbOpenDynaset)
     
      OtherProfiles = False
      Dim rsCount As Integer
      rsCount = 0
      On Error Resume Next
        While rsCount < 20
          If ((RS.EOF)) Then
            rsCount = 25
          Else
            If LCase$(RS("user")) <> "average" Then OtherProfiles = True
            
          End If
          RS.MoveNext
          rsCount = rsCount + 1
        Wend
         
        If OtherProfiles Then
          On Error Resume Next
          DoUpdateScript = True
         
          
         End If
         Set RS = Nothing
         Set DBOut = Nothing
    End If
   
  
    Dim ffb As Long, junk As String, junks() As String
    ffb = FreeFile
    Open App.path & "\resources\branding.txt" For Input As #ffb
    While (Not EOF(ffb))
      Line Input #ffb, junk
      junks = Split(junk, "==")
      Branding.Add Trim$(junks(1)), Trim$(junks(0))
      If Err.Number <> 0 Then GoTo jumpout
    Wend
    Close #ffb
    
jumpout:
    
    Dim DBPath As String
    
    If Right(App.path, 1) = "\" Then
       DBPath = App.path & "resources\sr16-2.mdb"
    Else
       DBPath = App.path & "\resources\sr16-2.mdb"
    End If
    App.HelpFile = App.path & "\Resources\Help\CalorietrackerHelp.chm"
    HelpPath = App.HelpFile & ">Main"
    Dim n As Long, i As Long
    Installdate = GetSetting(App.Title, "Settings", "FD", "2000-01-01")

    If True Then
       Dim ProtoB As Boolean, Sr16B As Boolean
CheckDB:
       On Error Resume Next
       Err.Clear
       Open App.path & "\resources\proto.mdb" For Input As #1
       If Err.Number = 0 Then ProtoB = True
       Close #1
       
       FirstRun = ProtoB
       Err.Clear
       Open App.path & "\resources\sr16-2.mdb" For Input As #1
       If Err.Number = 0 Then Sr16B = True
       Close #1
       
       If ProtoB And Sr16B Then
       Dim ret As VbMsgBoxResult
        

       
          
          ret = MsgBox("You had an old version of the program installed." & vbCrLf & "Would you like to attempt to import the old data? (this includes recipes, meals, and your diet history)", vbYesNo, "")
          If ret = vbYes Then
            Call Dialog.ShowIt(0, "Please wait", "Upgrading Database")
            Dialog.Show
            DoEvents
            'clear out the average and other stuff from the basic database
            Set DB = OpenDatabase(App.path & "\resources\proto.mdb")
           
            Set RS = DB.OpenRecordset("select * from profiles;", dbOpenDynaset)
            While Not RS.EOF
              RS.Delete
              RS.MoveNext
            Wend
            Set RS = Nothing
            Set DB = Nothing
            
            'until I get the update script fixed, it makes sense to just use the old database
            'this saves me from the problems I have been having with the update plan script
            
            
            'Name App.path & "\resources\sr16-2.mdb" As App.path & "\resources\plan.mdb"
            'Name App.path & "\resources\proto.mdb" As App.path & "\resources\sr16-2.mdb"
            'Set DB = OpenDatabase(DBPath)
            
        '    On Error GoTo 0
            'Call REadScriptMod.UpdateScript(App.path & "\resources\plan.mdb")
            'Kill App.path & "\resources\plan.mdb"
            
            'this just gets rid of the new database
            'msgbox "3.5"
            Kill App.path & "\resources\proto.mdb"
            
            Unload Dialog
          Else
            Kill App.path & "\resources\sr16-2.mdb"
            Name App.path & "\resources\proto.mdb" As App.path & "\resources\sr16-2.mdb"
          End If
       ElseIf ProtoB And Not Sr16B Then
           Name App.path & "\resources\proto.mdb" As App.path & "\resources\sr16-2.mdb"
       End If
    
    End If
    Installdate = GetSetting(App.Title, "Settings", "FD", "2000-01-01")
    If NoTrial Then
       Call SaveSetting(App.Title, "Settings", "PD", True)
       If Installdate = "2000-01-01" Or Installdate = "2000-09-01" Then Installdate = "2006-12-04"
       Call SaveSetting(App.Title, "Settings", "pd", True)
       Call SaveSetting(App.Title, "Settings", "fd", Installdate)
    
    End If
    If RestartTrial Then Installdate = "2001-01-01"
    Err.Clear
    Set DB = OpenDatabase(DBPath)
    If Err.Number <> 0 Then
       Err.Clear
       Set DB = OpenDatabase(App.path & "\resources\proto.mdb")
       If Err.Number <> 0 Then
           MsgBox "Something has gone wrong.  The database is missing. Please contact Brian at accounts@caloriebalancediet.com" & vbCrLf & "Sorry!", vbOKOnly
           OpenURL "http://www.caloriebalancediet.com/BadDatabase.asp?DB=bad&dbpath=" & Err.Number & Replace(Err.Description, " ", "_")
       
           End
       End If
    End If
   Set frmMain = New frmMainO
   
    DisplayDate = Today
    FirstChanged = DisplayDate
    
    If Installdate <= "2001-11-04" Then
     
      SaveSetting App.Title, "Settings", "FD", Today
      SaveSetting App.Title, "Settings", "PD", False
      frmLogin.RemainingDays = TrialDays
      
      FirstRun = True
      Paid = GetSetting(App.Title, "Settings", "PD", False)
      Paid = True
      frmLogin.SeriesDay = Day(Today) Mod 10
      Paid = True
      Call SaveSetting(App.Title, "settings", "pd", True)
    Else
      FirstRun = False
      Paid = GetSetting(App.Title, "Settings", "PD", False)
      frmLogin.RemainingDays = TrialDays
         
      frmLogin.SeriesDay = Day(Installdate) Mod 10
      Paid = True
      Call SaveSetting(App.Title, "settings", "pd", True)
    End If
    
    If Trim$(Interaction.Command$) <> "" Then
       MsgBox "Please sign in and then the plan will be loaded.", vbOKOnly, ""
    End If
   
    ReDim WatchHeads(6)
    WatchHeads(0) = "Calories"
    WatchHeads(1) = "Carbs"
    WatchHeads(2) = "Fat"
    WatchHeads(3) = "Protein"
    WatchHeads(4) = "Sugar"
    WatchHeads(5) = "Fiber"
    WatchHeads(6) = "Calories Net"
    
    Call loadConverts(Converts, ConvertsW)
    
    Dim T As String
    T = frmMain.LastUser
    frmMain.LastUser = "average"

    Call OpenUser(False, True)
    frmMain.LastUser = T
    frmMain.Show

    If FirstRun Then
       DoEvents
       Err.Clear
     
       If Err.Number <> 0 And DoDebug Then Stop
       DoEvents
       OpenURL "http://www.caloriebalancediet.com/HelpMovies/HelpFiles.asp?fi=300"
    End If
    
    Call OpenUser
    
        junk = ""
        junk = LCase(Trim(Branding("openOnWebsite")))
        If junk = "yes" Or junk = "true" Then
           Call frmMain.EasyHover1_Click(1)
        End If
    
    DoEvents
    frmMain.FlexDiet.Changed = False
   
    If Trim$(Interaction.Command$) <> "" And AlreadyLoaded = False Then
      MsgBox "Importing a plan can take up to 20 minutes depending on the size.", vbOKOnly, ""
      
      If DoUpdateScript Then
'          DBOut.Close
 '         RS.Close
  '        Set DBOut = Nothing
   '       Set RS = Nothing
          
          Call UpdateScript(Interaction.Command$)
          'AlreadyLoaded = True
      Else
          Call ReadScript(Replace(Interaction.Command$, """", ""), CurrentUser.Username, True)
      End If
    End If
    Call frmMain.LoadDeepSearch

End Sub

Public Sub OpenUser(Optional NewUSer As Boolean = False, Optional ReLoad As Boolean = False)
   On Error Resume Next
    Dim i As Long
   
    If ReLoad Then GoTo ReloadIt

noUser:
      CloseProgram = False
      If FirstRun Then frmLogin.LaunchNewUser = True
      frmLogin.Show vbModal, frmMain
      If CloseProgram Then End
      If Not frmLogin.ok Then End
    

    
    Unload frmLogin
    Unload FNewUserD
    
    CloseProgram = False
ReloadIt:
    

    Dim ret As Boolean, junk As String
    ret = LoadUser(frmMain.LastUser)
    
        junk = Branding("LoginWebsite")
        junk = Replace(junk, "#Uname#", CurrentUser.Username, , , vbTextCompare)
        junk = Replace(junk, "#Upassword#", CurrentUser.Password, , , vbTextCompare)
        
        DoEvents
        Branding.Add junk, "LogedWebsite"
   
    DoEvents
    
    If ret Then
    
        'these have to be here to allow nutmaxes to load the correct user profile and max values
        Set Nutmaxes = New Calories
        Call Nutmaxes.Init(CurrentUser.Username, Today, frmMain.SC)

        Call frmMain.FlexDiet.AddDataBase(DB, CurrentUser.Username, Today, Nutmaxes)
        Call frmMain.FlexDiet.SetHeads(WatchHeads)
        Call frmMain.Exercise.AddDataBase(DB)
        Call FNewExercise.SetDB
     
        Dim temp As Recordset
        Dim PlanDate As Date, vars As New Collection
    
        Set temp = DB.OpenRecordset("Select * from profiles where user = '" & CurrentUser.Username & "';", dbOpenDynaset)
        On Error Resume Next
        temp.Close
        Set temp = Nothing
        frmMain.Show 'todo: this is just temp, remove
        Call DisplayDay(DisplayDate)
        Call frmMain.MakeMealList
        
        junk = GetSetting(App.Title, "Settings", "IntroMain", "")
        StartingProgram = True
        If InStr(1, junk, CurrentUser.Username, vbTextCompare) = 0 Then
           junk = junk & "," & CurrentUser.Username
           SaveSetting App.Title, "Settings", "IntroMain", junk
           DoEvents
           frmMain.ZOrder
        End If
           Call frmMain.EasyHover1_Click(2)
            frmMain.Caption = Branding("caption") & " for " & CurrentUser.Username

        
    Else
        MsgBox "Login Failed", vbOKOnly, ""
        GoTo noUser
    End If
End Sub
Public Function LoadUser(USER As String) As Boolean
On Error Resume Next
  Dim i As Long
  Dim temp As Recordset
  
  Set temp = DB.OpenRecordset("Select * From Profiles where (((Profiles.User)='" & USER & "'));", dbOpenDynaset)
  
  If Not temp.RecordCount = 0 Then
    temp.MoveFirst
    On Error Resume Next
    With CurrentUser
       .Username = USER
       .Weight = temp.Fields("Weight")
       .Height = temp.Fields("Height")
       .BMR = temp("BMR")
       .CalPound = Val(temp("CalPound"))
       .Password = temp("password")
       .HashNumber = temp("hashnumber")
    End With
    Dim junk As String, Parts() As String
    Dim j As Long
    junk = temp("OtherWatches")
    Parts = Split(junk, ",")
    j = 6 ' UBound(WatchHeads)
    For i = 0 To UBound(Parts)
      If Parts(i) <> "" Then
         j = j + 1
         ReDim Preserve WatchHeads(j)
         WatchHeads(j) = Trim$(Parts(i))
      End If
    Next i
    LoadUser = True
  
  Else
    LoadUser = False
  End If
  temp.Close
  Set temp = Nothing
End Function


Public Function FixDate(Daydate As Date) As String
    'FixDate = Month(Daydate) & "/" & Day(Daydate) & "/" & Year(Daydate)
    FixDate = DateHandler.IsoDate(Daydate)
End Function

Public Sub DisplayDay(Daydate As Date)
On Error GoTo errhandl
   Call Nutmaxes.Update(Daydate)
   frmMain.FlexDiet.OpenDay Daydate
  ' frmMain.FlexDiet.DisplayDate = DayDate
   frmMain.Exercise.OpenWeek Daydate
   frmMain.MP.OpenWeek
'   frmMain.MP.OpenWeek
   'Call frmJournal.ResetMe(DayDate)
   Exit Sub
errhandl:
  
End Sub
Public Sub SaveDay(Today As Date)
  On Error Resume Next
 ' Call frmMain.MP.SaveWeek(today)
  Call frmMain.FlexDiet.SaveDay(Today)
  Call frmMain.Exercise.SaveWeek(Today)
  
End Sub
Public Function ConvertFractionToDecimal(ByVal STR As String) As String
On Error Resume Next
Dim strFraction As String
Dim strWholeNumber As String
Dim strNumerator As String
Dim StrDenominator As String
Dim intFirst As Integer
Dim intLength As Long
STR = Trim$(STR)

If InStr(1, STR, "/") = 0 Then
   ConvertFractionToDecimal = Val(STR)
   Exit Function
End If

If InStr(1, STR, " ") Then
    strWholeNumber = Mid(STR, 1, InStr(1, STR, " ") - 1)
Else
    strWholeNumber = "0"
End If

If strWholeNumber <> "0" Then
    STR = Trim(Mid(STR, InStr(1, STR, " ")))
    intLength = InStr(1, STR, "/") - 1
    strNumerator = Mid(STR, 1, intLength)
    StrDenominator = Mid(STR, InStr(1, STR, "/") + 1)
Else
    STR = Trim(STR)
    intLength = InStr(1, STR, "/") - 1
    strNumerator = Mid(STR, 1, intLength)
    StrDenominator = Mid(STR, InStr(1, STR, "/") + 1)
End If
If Val(StrDenominator) = 0 Then
  ConvertFractionToDecimal = Val(strWholeNumber) + Val(strNumerator)
Else
  ConvertFractionToDecimal = Val(strWholeNumber) + Val(strNumerator) / Val(StrDenominator)
End If


End Function


Public Function ConvertDecimalToFraction(ByVal STR As String) As String
On Error Resume Next
Dim intCountDecimalPoints As Integer
Dim strWholeNumber As String
Dim strDecimal As String
Dim StrDenominator As String
Dim intDecimalMarker As Integer
Dim intWholeNumberLength As Integer
Dim i As Long
Dim temp As String
STR = Trim$(STR)

If InStr(1, STR, ".") = 0 Then
  ConvertDecimalToFraction = STR
  Exit Function
End If

intDecimalMarker = InStr(1, STR, ".")
intCountDecimalPoints = Len(Mid(STR, InStr(1, STR, ".") + 1))

intWholeNumberLength = intDecimalMarker - 1
strWholeNumber = Mid(STR, 1, intWholeNumberLength)
strDecimal = Mid(STR, intDecimalMarker + 1)

temp = CheckForRepeatingDecimal(Val("." & strDecimal), strDecimal, StrDenominator)

If temp = "0" Then
    StrDenominator = "1"
    For i = 1 To intCountDecimalPoints
        StrDenominator = StrDenominator & "0"
    Next
End If

If strWholeNumber = "0" Then
    ConvertDecimalToFraction = Trim(ReduceToLCD(Val(strDecimal), Val(StrDenominator)))
Else
    ConvertDecimalToFraction = Trim(strWholeNumber & " " & ReduceToLCD(Val(strDecimal), Val(StrDenominator)))
End If
End Function


Private Function ReduceToLCD(dblNumerator, dblDenominator As Double) As String
Dim i As Long
On Error Resume Next


For i = 2 To dblDenominator
    If i > dblDenominator Then GoTo fini
    If dblNumerator Mod i = 0 And dblDenominator Mod i = 0 Then
        dblNumerator = dblNumerator / i
        dblDenominator = dblDenominator / i
        i = 1
    End If
Next

fini:
ReduceToLCD = dblNumerator & "/" & dblDenominator

End Function


Private Function CheckForRepeatingDecimal(sDecimal As Single, Num As String, Denom As String) As String
On Error Resume Next
Dim i As Long
Dim j As Single, J2 As Single
Dim ST As String
Dim L As Long
ST = STR$(sDecimal)
L = Len(STR$(ST)) - InStr(1, ST, ".") - 1

For i = 3 To 25 Step 2
   j = Round(sDecimal * i * 10000, L) / 10000
   J2 = 1 / j
   If Abs(J2 - Int(J2)) < 0.001 Or Abs(Int(J2) - J2) > 0.999 Then
     ST = Round(j, 4)
     L = Len(ST) - InStr(1, ST, ".", vbBinaryCompare)
     Num = 10 ^ L * Round(j, 4)
     Denom = 10 ^ L * i
     Exit Function
   End If
   If Abs(j - Int(j)) < 0.001 Then
     ST = Round(j, 4)
     L = Len(ST) - InStr(1, ST, ".", vbBinaryCompare)
     Num = Round(j, 4)
     Denom = i
     Exit Function
   End If
Next i
CheckForRepeatingDecimal = "0"
End Function

Public Function WrapMonth(ByVal X)
On Error Resume Next
  While X > 12
   X = X - 12
  Wend
  While X < 1
   X = X + 12
  Wend
  WrapMonth = X
End Function


Public Function Log10(X)
On Error Resume Next
  If X = 0 Then
    Log10 = -10000
  Else
    Log10 = Log(Abs(X)) / Log(10)
  End If
End Function

Public Function FindFirstDay(Week As Date) As Date
On Error Resume Next
Dim i As Long
Dim T As Date
Dim TT As Date
TT = Date
i = Weekday(Week, vbSunday)
 T = DateAdd("w", -1 * i + 1, Week)
 TT = T
FindFirstDay = TT
End Function

Public Sub LinearFit(Data() As Single, slope As Double, B As Double)
On Error GoTo errhandl
Dim n As Long, i As Long
Static nn As Long
Static SX As Double, Sy As Double
Static SXY As Double, Sx2 As Double
Dim tSX As Double, tSy As Double

Dim T As Double
Dim d As Double
Dim startN As Long
  n = UBound(Data, 2)
  SX = 0
  Sy = 0
  SXY = 0
  Sx2 = 0
  nn = 0
  startN = 0
  Dim X As Double
  For i = startN To n
    T = Data(1, i)
    X = Data(0, i)
    Sy = Sy + T
    SXY = SXY + T * X
    Sx2 = Sx2 + X * X
    SX = SX + X
    nn = nn + 1
  Next i

i = nn
tSX = SX / i
tSy = Sy / i
d = Sx2 - i * tSX ^ 2
If d <> 0 Then
  B = (tSy * Sx2 - tSX * SXY) / d
  slope = (SXY - SX * tSy) / d
Else
  B = 0
  slope = 10000000
End If
errhandl:
End Sub


Public Function WebHexColor(clr As Long)
Dim r As Long, g As Long, B As Long
Dim oClr As OLE_COLOR
On Error Resume Next
    Call LongToRGB(clr, r, g, B)
    oClr = B + 256 * (g + 256 * r)
    WebHexColor = Right$("000000" & Hex$(oClr), 6)
End Function
Private Function LongToRGB(lColor As Long, iRed As Long, iGreen As Long, iBlue As Long) As String
On Error Resume Next
   iRed = lColor Mod 256
   iGreen = ((lColor And &HFF00) / 256&) Mod 256&
   iBlue = (lColor And &HFF0000) / 65536
   LongToRGB = STR$(iRed) & STR$(iGreen) & STR$(iBlue)
End Function




Public Function G_Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    '* Purpose: Module scope error handling function
  If DoDebug Then
    MsgBox "Error occured:" & vbNewLine & _
           "Module: " & ModuleName & vbNewLine & _
           "Function: " & ProcName & vbNewLine & _
           "Description: " & ErrorDesc
  End If

    G_Err_Handler = True

End Function
