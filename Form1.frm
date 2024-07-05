VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Save MealPlan"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox URL 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   4875
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5715
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":57E2
   End
   Begin VB.TextBox PlanName 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   4935
   End
   Begin CalorieBalance.MonthDayPicker MD1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4683
   End
   Begin CalorieBalance.MonthDayPicker MD2 
      Height          =   2655
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4683
   End
   Begin VB.Label Label10 
      Caption         =   "Instructions URL"
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   1560
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   3945
   End
   Begin VB.Label Description 
      Caption         =   "Plan Description"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Plan Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "End Day"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Start Day"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public PlanMode As Integer
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub Command1_Click()
On Error GoTo errhandl
CD.CancelError = True
CD.Filter = "Meal Plans (*.cbm) | *.cbm"
CD.InitDir = App.path & "\resources\plans"
CD.ShowSave

If PlanMode = 0 Then
   SaveTempFile CD.Filename
   Unload Me
Else
  If SaveExerciseTempFile(CD.Filename) Then Call Unload(Me)
End If

Exit Sub
errhandl:

End Sub
Private Function SaveTempFile(Filename As String) As Boolean
On Error Resume Next
Dim DBOut As Database, i As Long
Dim temp As Recordset, temp2 As Recordset
Dim temp4 As Recordset
Dim mStart As Date, mEnd As Date
mStart = MD1.GetDate
mEnd = MD2.GetDate
SaveTempFile = True
If Trim$(PlanName.Text) = "" Then
  MsgBox "Please enter a plan name.", vbOKOnly, ""
  SaveTempFile = False
  Exit Function
End If

CopyFile App.path & "\resources\tempscript.mdb", App.path & "\resources\temp\temp.mdb", 0
Set DBOut = OpenDatabase(App.path & "\resources\temp\temp.mdb")

Set temp = DB.OpenRecordset("select * from foodgroups;", dbOpenDynaset)
Set temp2 = DBOut.OpenRecordset("select * from foodgroups;", dbOpenDynaset)
While Not temp.EOF
   temp2.AddNew
   For i = 0 To temp2.Fields.Count - 1
      temp2(i) = temp(temp2(i).Name)
   Next i
   temp2.Update
   temp.MoveNext
Wend

Set temp = DBOut.OpenRecordset("Select * from mealplanner;", dbOpenDynaset)
temp.AddNew
temp("Mealname") = PlanName.Text & " "
temp("description") = RTB1.Text & " "
temp("url") = URL.Text & " "
temp("mealid") = -1
temp.Update


'Set temp = DB.OpenRecordset("SELECT meals.id, Meals.MealNumber, Meals.EntryDate, Meals.EntryDate, MealDefinition.Serving, MealDefinition.Unit, WEIGHT.Amount, WEIGHT.Gm_Wgt, ABBREV.* " _
'& "FROM Meals LEFT JOIN ((MealDefinition LEFT JOIN ABBREV ON MealDefinition.AbbrevID=ABBREV.Index) LEFT JOIN WEIGHT ON (MealDefinition.AbbrevID=WEIGHT.Index) AND (MealDefinition.Unit=WEIGHT.Msre_Desc)) ON Meals.MealId=MealDefinition.MealID " _
'& "Where ((user ='" & CurrentUser.Username & "') and ((Meals.EntryDate) >= #" & mStart & "# And (Meals.EntryDate) <= #" & mEnd & "#)) " _
'& "ORDER BY ABBREV.Foodname;", dbOpenDynaset)

Dim SQL As String
SQL = "SELECT Meals.User, Meals.EntryDate, Abbrev.* " _
 & "FROM Meals INNER JOIN (MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID = Abbrev.Index) ON Meals.MealId = MealDefinition.MealID " _
 & "WHERE (((Meals.User)='" & CurrentUser.Username & "') AND ((Meals.EntryDate) >= #" & FixDate(mStart) & "# And (Meals.EntryDate) <= #" & FixDate(mEnd) & "#)) " _
 & "ORDER BY Abbrev.Foodname;"
Set temp = DB.OpenRecordset(SQL, dbOpenDynaset)

Set temp2 = DBOut.OpenRecordset("select * from abbrev;", dbOpenDynaset)

While Not temp.EOF

  Set temp4 = DBOut.OpenRecordset("select * from abbrev where index=" & temp("index") & ";", dbOpenDynaset)
  
  If temp4.EOF And temp4.BOF Then
      temp2.AddNew
      
      For i = 0 To temp2.Fields.Count - 1
      
         temp2(i) = temp(temp2.Fields(i).Name)
      Next i
      temp2.Update
  End If
  temp.MoveNext
Wend
Dim MealPlanID As Long
MealPlanID = -1
'For i = 0 To LMeals.ListCount - 1
'   If LMeals.Selected(i) Then
'      MealPlanID = LMeals.ItemData(i)
'   End If
'Next i

If MealPlanID > -1 Then
Set temp = DB.OpenRecordset("SELECT Abbrev.*, MealPlanner.PlanID " _
 & "FROM (MealPlanner INNER JOIN MealDefinition ON MealPlanner.MealID = MealDefinition.MealID) INNER JOIN Abbrev ON MealDefinition.AbbrevID = Abbrev.Index " _
 & "WHERE (((MealPlanner.PlanID)=" & MealPlanID & "));", dbOpenDynaset)

While Not temp.EOF
  Set temp4 = DBOut.OpenRecordset("select * from abbrev where index=" & temp("index") & ";", dbOpenDynaset)
  
  If temp4.EOF And temp4.BOF Then
      temp2.AddNew
      For i = 0 To temp2.Fields.Count - 1
         temp2(i) = temp(temp2.Fields(i).Name)
      Next i
      temp2.Update
  End If
  temp.MoveNext
Wend
End If

temp2.Close

Set temp2 = DBOut.OpenRecordset("select * from abbrev where ndb_no='-100';", dbOpenDynaset)
Dim indexs()
i = 0
While Not temp2.EOF
   ReDim Preserve indexs(i)
   indexs(i) = temp2("index")
   i = i + 1
   temp2.MoveNext
Wend
temp2.Close
Set temp2 = DBOut.OpenRecordset("select * from abbrev;", dbOpenDynaset)
Dim j As Integer
For j = 0 To i - 1
   Set temp = DB.OpenRecordset("SELECT Abbrev.*, RecipesIndex.AbbrevID " _
     & "FROM (RecipesIndex INNER JOIN Recipes ON RecipesIndex.RecipeID = Recipes.RecipeID) INNER JOIN Abbrev ON Recipes.ItemID = Abbrev.Index " _
     & "WHERE (((RecipesIndex.AbbrevID)=" & indexs(j) & " ));", dbOpenDynaset)
     
  While Not temp.EOF
     Set temp4 = DBOut.OpenRecordset("select * from abbrev where index=" & temp("index") & ";", dbOpenDynaset)
     If temp4.EOF And temp4.BOF Then
      temp2.AddNew
      For i = 0 To temp2.Fields.Count - 1
         temp2(i) = temp(temp2.Fields(i).Name)
      Next i
      temp2.Update
     End If
     temp.MoveNext
  Wend
Next j



Set temp2 = DBOut.OpenRecordset("select * from abbrev;", dbOpenDynaset)
Dim temp3 As Recordset
Set temp3 = DBOut.OpenRecordset("select * from weight;", dbOpenDynaset)

temp2.MoveFirst
While Not temp2.EOF
   Set temp = DB.OpenRecordset("select * from weight where index = " & temp2("index") & ";", dbOpenDynaset)
   While Not temp.EOF
     temp3.AddNew
     For i = 0 To temp3.Fields.Count - 1
        temp3(i) = temp(temp3.Fields(i).Name)
     Next i
     
     temp3.Update
     temp.MoveNext
   Wend
   temp2.MoveNext
Wend


Set temp = DBOut.OpenRecordset("select * from abbrev where ndb_no='-100'", dbOpenDynaset)
Set temp2 = DBOut.OpenRecordset("select * from recipesindex;", dbOpenDynaset)
On Error Resume Next
  While Not temp.EOF
     Set temp3 = DB.OpenRecordset("SELECT  RecipesIndex.* " _
       & "FROM RecipesIndex INNER JOIN Recipes ON RecipesIndex.RecipeID = Recipes.RecipeID " _
       & "WHERE (((RecipesIndex.AbbrevID)=" & temp("index") & "));", dbOpenDynaset)
  
       temp2.AddNew
       For i = 0 To temp2.Fields.Count - 1
          temp2(i) = temp3(temp2.Fields(i).Name)
       Next i
       temp2.Update
     temp.MoveNext
  Wend

Set temp = DBOut.OpenRecordset("select * from abbrev where ndb_no='-100'", dbOpenDynaset)
Set temp2 = DBOut.OpenRecordset("select * from recipes;", dbOpenDynaset)
  While Not temp.EOF
  
     Set temp3 = DB.OpenRecordset("SELECT Recipes.* " _
     & "FROM Recipes INNER JOIN RecipesIndex ON Recipes.RecipeID = RecipesIndex.RecipeID " _
     & "WHERE (((RecipesIndex.AbbrevID)=" & temp("index") & "));", dbOpenDynaset)
     
  
       While Not temp3.EOF
          temp2.AddNew
          For i = 0 To temp2.Fields.Count - 1
            temp2(i) = temp3(temp2.Fields(i).Name)
          Next i
          temp2.Update
          temp3.MoveNext
       Wend
     
     temp.MoveNext
  Wend

Set temp = DB.OpenRecordset("SELECT MealPlanner.*, Meals.EntryDate, Meals.User " _
   & "FROM Meals INNER JOIN MealPlanner ON Meals.MealId = MealPlanner.MealID " _
   & " WHERE (((Meals.EntryDate)>=#" & FixDate(mStart) & "#) AND ((Meals.EntryDate)<=#" & FixDate(mEnd) & "#) " _
   & " AND ((Meals.User)='" & CurrentUser.Username & "'));", dbOpenDynaset)

Set temp2 = DBOut.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
On Error Resume Next
While Not temp.EOF
     temp2.AddNew
     For i = 0 To temp2.Fields.Count - 1
        temp2(i) = temp(temp2.Fields(i).Name)
     Next i
     temp2.Update
     If Err.Number <> 3022 And Err.Number <> 0 Then
     
     End If
     temp.MoveNext
Wend

If MealPlanID > -1 Then

Set temp = DB.OpenRecordset("SELECT * " _
   & "FROM MealPlanner " _
   & " WHERE planid=" & MealPlanID & ";", dbOpenDynaset)

Set temp2 = DBOut.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
On Error Resume Next
While Not temp.EOF
     temp2.AddNew
     For i = 0 To temp2.Fields.Count - 1
        temp2(i) = temp(temp2.Fields(i).Name)
     Next i
     temp2.Update
     If Err.Number <> 3022 Then
    
     End If
     temp.MoveNext
Wend
  


End If
'On Error
Set temp = DBOut.OpenRecordset("select * from mealplanner;", dbOpenDynaset)

Set temp2 = DBOut.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
While Not temp.EOF
     Set temp3 = DB.OpenRecordset("Select * from mealdefinition where mealid=" & temp("mealid") & ";", dbOpenDynaset)
     While Not temp3.EOF
         temp2.AddNew
         For i = 0 To temp2.Fields.Count - 1
            temp2(i) = temp3(temp2.Fields(i).Name)
         Next i
         temp2("user") = "blank"
         temp2.Update
         temp3.MoveNext
     Wend
     temp.MoveNext
Wend

Set temp = DB.OpenRecordset("SELECT * FROM Meals " _
   & " WHERE (((Meals.EntryDate)>=#" & FixDate(mStart) & "#) AND ((Meals.EntryDate)<=#" & FixDate(mEnd) & "#) " _
   & " AND ((Meals.User)='" & CurrentUser.Username & "'));", dbOpenDynaset)

Set temp2 = DBOut.OpenRecordset("select * from meals;", dbOpenDynaset)
While Not temp.EOF
     temp2.AddNew
     For i = 0 To temp2.Fields.Count - 1
        temp2(i) = temp(temp2.Fields(i).Name)
     Next i
     temp2.Update
     temp.MoveNext
Wend

Set temp = Nothing
Set temp2 = Nothing
Set temp3 = Nothing
Set temp4 = Nothing
DBOut.Close
Set DBOut = Nothing

CopyFile App.path & "\resources\temp\temp.mdb", Filename, 0
Exit Function
errhandl:
MsgBox Err.Description, vbCritical
SaveTempFile = False
End Function


Private Sub Command2_Click()


    On Error GoTo Err_Proc
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "Form1", "Command2_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_Load()
On Error Resume Next
Dim j As Date
j = firstSunday(Today)
MD2.SetDate DateAdd("d", 6, j)
MD1.SetDate j
If PlanMode = 0 Then
    Label4.Caption = "This dialog allows you to save the meals that have already been added to the meal planner.  To use this dialog: " & vbCrLf _
    & "   1.  Choose a useful name for the plan.  This will be displayed in the toolboxes." & vbCrLf _
    & "   2.  Give a short description of the plan.  This should help people to load the plan. " & vbCrLf _
    & "   3.  If your plan has a helpful webpage, you can include a link to it here.  " & vbCrLf _
    & "   4.  Finally, you need to select the start and end dates of your plan. "
    
    'LMeals.Visible = True
    'LMeals.Clear
    'Dim rs As Recordset, i As Long
    'Set rs = DB.OpenRecordset("Select * from mealplanner where mealid=-1 and calories=0;", dbOpenDynaset)
    'i = 1
    'LMeals.AddItem "Favorites"
    'LMeals.ItemData(0) = 1
    'While Not rs.EOF
    '  If rs("planid") > 1 Then
    '    LMeals.AddItem rs("Mealname")
    '    LMeals.ItemData(i) = rs("planid")
    '    i = i + 1
    '  End If
    '  rs.MoveNext
    'Wend
    'LMeals.Selected(0) = True
    'Set rs = Nothing
Else
   ' LMeals.Visible = False
    Label4.Caption = "Choose the week that your exercise plan will run.  Enter a good description and name for this workout and then hit the save button."
End If
End Sub

Private Function SaveExerciseTempFile(Filename As String) As Boolean


    On Error GoTo Err_Proc

Dim DBOut As Database, i As Long
Dim temp As Recordset, temp2 As Recordset

Dim mStart As Date, mEnd As Date
mStart = MD1.GetDate
mEnd = MD2.GetDate
SaveExerciseTempFile = True
If Trim$(PlanName.Text) = "" Then
  MsgBox "Please enter a plan name.", vbOKOnly, ""
  SaveExerciseTempFile = False
  Exit Function
End If

CopyFile App.path & "\resources\tempscript.mdb", App.path & "\resources\temp\temp.mdb", 0
Set DBOut = OpenDatabase(App.path & "\resources\temp\temp.mdb")

Set temp = DBOut.OpenRecordset("Select * from mealplanner;", dbOpenDynaset)
temp.AddNew
temp("Mealname") = PlanName.Text & " "
temp("description") = RTB1.Text & " "
temp("url") = URL.Text & " "
temp("mealid") = -1
temp("calories") = 1
temp.Update

Dim mStartWeek As Date, mEndWeek As Date


mStartWeek = firstSunday(MD1.GetDate)
mEndWeek = DateAdd("d", 6, firstSunday(MD2.GetDate))


Set temp = DB.OpenRecordset(" SELECT AbbrevExercise.* " _
& "FROM ExerciseLog INNER JOIN AbbrevExercise ON ExerciseLog.exerciseID = AbbrevExercise.Index " _
& "Where (" _
& "((ExerciseLog.Week) >= #" & FixDate(mStartWeek) & "#) and " _
& "((ExerciseLog.Week) <= #" & FixDate(mEndWeek) & "#) and " _
& "((user) = '" & CurrentUser.Username & "')" _
& " ) " _
& "ORDER BY AbbrevExercise.Index;", dbOpenDynaset)

Dim AbbrevOut As Recordset, Abbrev As Recordset, Lastfood As Long
Set AbbrevOut = DBOut.OpenRecordset("select * from abbrevexercise;", dbOpenDynaset)
While Not temp.EOF
  If temp(0) <> Lastfood Then
       AbbrevOut.AddNew
       For i = 0 To AbbrevOut.Fields.Count - 1
          AbbrevOut(i) = temp(AbbrevOut.Fields(i).Name)
       Next i
       AbbrevOut.Update
  End If
  If temp(0) <> vbNull Then Lastfood = temp(0)
  temp.MoveNext
Wend
temp.Close
Set temp = Nothing
AbbrevOut.Close
Set AbbrevOut = Nothing


Set temp = DB.OpenRecordset(" SELECT ExerciseLog.* " _
& "FROM ExerciseLog INNER JOIN AbbrevExercise ON ExerciseLog.exerciseID = AbbrevExercise.Index " _
& "Where (" _
& "((ExerciseLog.Week) >= #" & FixDate(mStartWeek) & "#) and " _
& "((ExerciseLog.Week) <= #" & FixDate(mEndWeek) & "#) and " _
& "((user) = '" & CurrentUser.Username & "')" _
& " ) " _
& "ORDER BY ExerciseLog.Week, ExerciseLog.Order;", dbOpenDynaset)
Dim LogOut As Recordset, junk As String, junks() As String, junkss As String, j As Long
Set LogOut = DBOut.OpenRecordset("select * from exerciselog;", dbOpenDynaset)
While Not temp.EOF
   LogOut.AddNew
   For i = 0 To LogOut.Fields.Count - 1
     If LCase$(LogOut.Fields(i).Name) = "weekinfo" Then
        junk = temp(LogOut.Fields(i).Name)
        junks = Split(junk, "~")
        junkss = ""
        For j = 0 To UBound(junks) - 1
           If InStr(1, junks(j), "*") = 0 Then
              junkss = junkss & "*" & Trim$(junks(j)) & " ~ "
           Else
              junkss = junkss & Trim$(junks(j)) & " ~ "
           End If
        Next
        
        If Trim$(junks(j)) <> "" Then
        If InStr(1, junks(j), "*") = 0 Then
           junkss = junkss & "*" & junks(j)
        Else
           junkss = junkss & junks(j)
        End If
        End If
        LogOut(i) = junkss
     Else
        LogOut(i) = temp(LogOut.Fields(i).Name)
     End If
   Next i
   LogOut("user") = "blank"
   LogOut.Update
   temp.MoveNext
Wend
temp.Close
Set temp = Nothing
LogOut.Close
Set LogOut = Nothing

DBOut.Close
Set DBOut = Nothing
CopyFile App.path & "\resources\temp\temp.mdb", Filename, 0

Exit Function
errhandl:
MsgBox "Unable to save plan. Please check all fields.", vbOKOnly, ""
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "Form1", "SaveExerciseTempFile", Err.Description
    Resume Exit_Proc


End Function


Private Sub LMeals_Click()


    On Error GoTo Err_Proc

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "Form1", "LMeals_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
