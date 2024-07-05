VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Dim Defaults As Collection
Dim mNames As Collection
Dim mUSerName As String
Dim mToday As Date
Dim Script As ScriptControl
Dim Alpha As Single
Dim Planscript As String
Dim StartDate As Date

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function IdealsNames(ByVal Key) As String


    On Error GoTo Err_Proc
  IdealsNames = mNames(Key)
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "Calories", "IdealsNames", Err.Description
    Resume Exit_Proc


End Function
Public Function ideals(ByVal Key As String) As Variant


    On Error GoTo Err_Proc
  ideals = Defaults(Key)
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "Calories", "Ideals", Err.Description
    Resume Exit_Proc


End Function
Public Function Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
On Error Resume Next
  If LCase$(Key) = "carbohydrate" Then Key = "carbs"
  Item = Defaults(Key) * Alpha
End Function

Public Function Update(ByVal NewDate As Date, Optional SetAlpha As Single = 0) As Single
Dim sTemp As Recordset


On Error Resume Next
  Dim AddNew As Boolean

   Set sTemp = DB.OpenRecordset("SELECT * FROM DailyLog WHERE ((Date=#" & FixDate(mToday) & "#) AND (user='" & mUSerName & "'));", dbOpenDynaset)
   If SetAlpha = 0 Then Alpha = sTemp.Fields("Alpha") Else Alpha = SetAlpha
   If Err.Number = 3021 Then AddNew = True Else AddNew = False
   Err.Clear
   
   If Alpha = 0 Or AddNew Then
        Dim Period As Long
        Period = DateDiff("d", NewDate, "1/1/" & Year(Today))
        Script.AddCode "DayofYear = " & Period
        Script.AddCode "DietDay = " & DateDiff("d", StartDate, NewDate)
        Script.AddCode Planscript
        Alpha = Script.Eval("alpha")
        If Alpha = 0 Then Alpha = 1
        Update = Alpha
   
        If AddNew Then
            sTemp.AddNew
            sTemp.Fields("user") = mUSerName
            sTemp.Fields("date") = mToday 'FixDate(mToday)
        Else
            sTemp.Edit
        End If
        sTemp.Fields("Alpha") = Alpha
        sTemp.Update
   End If
   sTemp.Close
   Set sTemp = Nothing
End Function

Public Sub Add(ByVal Item As Variant, ByVal Key As String)
  On Error Resume Next
  
  Defaults.Remove Key
  Defaults.Add Item, Key
  mNames.Remove Key
  mNames.Add Key, Key
End Sub

Private Sub Class_Initialize()


    On Error GoTo Err_Proc
Set Defaults = New Collection
Set mNames = New Collection

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "Calories", "Class_Initialize", Err.Description
    Resume Exit_Proc


End Sub

Public Sub Init(Username As String, Today As Date, SC As ScriptControl)
mToday = Today
mUSerName = Username
Dim Help As Helpers
Set Script = SC
Set Help = New Helpers
On Error Resume Next
Script.AddObject "Help", Help, True
'On Error
Call GetMaxes(Username)
End Sub

Private Function GetMaxes(Username As String)
   Dim i As Long, junk As String, n As Long, junk2, junk3 As String
   Dim sTemp As Recordset
   Dim Cals As Single
   Dim Skip As Boolean
   
   
   Set sTemp = DB.OpenRecordset("SELECT * " _
                               & "From Profiles " _
                               & "WHERE (((Profiles.USer)=""" & Username & """));", dbOpenDynaset)
 
   Cals = sTemp("calories") 'sTemp("bmr") + sTemp("Weightlossrate") * sTemp("calpound") / 7
   
   junk2 = sTemp("protein") * Cals / 4
   Call Add(junk2, "Protein")
   junk2 = sTemp("fat") * Cals / 9
   Call Add(junk2, "Fat")
   junk2 = (sTemp("Carbs") * Cals / 4) ' * (1 - sTemp("sugar"))
   junk3 = (sTemp("sugar") * Cals / 4) ' * (sTemp("Carbs") * Cals / 4)
   Call Add(junk2, "Carbs")
   Call Add(Val(junk3), "Sugar")
   
   Call Add(sTemp("fiber"), "fiber")
   
   Call Add(Cals, "Calories")
   On Error Resume Next
   StartDate = sTemp("StartDietPlanDate")
   Planscript = sTemp("PlanScript")
   Script.AddCode "ExtraInfo = " & sTemp.Fields("ExtraInfo")
   Script.AddCode "Weightlossrate = " & sTemp("weightlossrate")
   Script.AddCode "CalPerPound = " & sTemp("calpound")
   Script.AddCode "Bmr = " & sTemp("bmr")
   Script.AddCode "Calories = " & Cals
 
 '  On Error
   
   sTemp.Close
   Set sTemp = Nothing
   
   
 
   Set sTemp = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
   n = sTemp.Fields.Count
   For i = 0 To n - 1
      Skip = False
      junk = sTemp.Fields(i).Name
      junk = Replace(junk, "Ideals.", "")
      junk2 = sTemp.Fields(i) + 0 '* Cals / 2000
      junk3 = LCase$(junk)
      If junk3 <> "sugar" And junk3 <> "fiber" And _
         junk3 <> "calories" And junk3 <> "carbs" And _
         junk3 <> "fat" And junk3 <> "protein" And _
         junk3 <> "carbohydrates" Then
        Call Add(junk2, junk)
      End If
   Next i
   
   Call Update(mToday)
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
