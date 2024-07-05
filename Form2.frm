VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   8760
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   2535
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Foodname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Calories"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Serving"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "grams"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Calories"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BPause As Boolean
Dim BPause2 As Boolean

Private Sub LoadMenu()



'On Error GoTo errhandl
Dim temp As Recordset, rs As Recordset
Dim DD As Date, ddi As Long, ddd As Date
Dim ff As Long, junk As String
openfile:
'On Error Resume Next
DD = firstSunday(DisplayDate)
For ddi = 0 To 6
  ddd = DateAdd("d", ddi, DD)
  Set temp = DB.OpenRecordset("SELECT Meals.MealId, Meals.User, Meals.EntryDate, MealPlanner.MealName, MealPlanner.Description, MealPlanner.Instructions, Meals.EntryDate, Meals.MealNumber " _
   & "FROM Meals INNER JOIN MealPlanner ON Meals.MealId=MealPlanner.MealID " _
   & "WHERE (((Meals.User)='" & CurrentUser.Username & "') AND ((Meals.EntryDate)=#" & FixDate(ddd) & "#)) " _
   & "ORDER BY Meals.EntryDate, Meals.MealNumber;", dbOpenDynaset)



 ' On Error GoTo errhandl
  While Not temp.EOF
            LV2.ListItems.Clear
            junk = temp("mealname") & vbCrLf
            Dim rs2 As Recordset
            Set rs = DB.OpenRecordset("SELECT MealDefinition.*, Abbrev.* " _
              & "FROM MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID=Abbrev.Index " _
              & "WHERE (((MealDefinition.MealID)=" & temp("mealid") & "));", dbOpenDynaset)
            Dim cc As Single
            
'            Print #1, "<table><tr><td>Servings</td><td>Units</td><td>Foodname</td>"
'            Print #1, "<td>Calories</td><td>Fat (gm)</td><td>Carbs (gm)</td><td>Protein (gm)</td></tr>"
            Dim Cals As Single, fff As Single, cbs As Single, pro As Single
            Dim ll As Long
            Cals = 0
            fff = 0
            cbs = 0
            pro = 0
            ll = 0
            While Not rs.EOF
              Dim lv As ListItem
              Set lv = LV2.ListItems.Add(, , rs("foodname"))
              lv.SubItems(1) = rs("unit")
              lv.SubItems(2) = rs("serving")
              Dim IsGrams() As Boolean
              ReDim Preserve IsGrams(ll)
              If LCase$(Trim$(rs("unit"))) = "grams" Then IsGrams(ll) = True
              'junk = junk & rs("serving") & " " & rs("unit") & " " & rs("foodname")
              Dim GMMs As Single
              GMMs = TranslateUnitToGrams(rs("abbrev.index"), rs("unit")) / 100 * rs("serving")
              cc = GMMs
                  If Not IsNull(rs("calories")) Then Cals = Cals + rs("calories") * cc
                  If Not IsNull(rs("fat")) Then fff = fff + rs("fat") * cc
                  If Not IsNull(rs("carbs")) Then cbs = cbs + rs("carbs") * cc
                  If Not IsNull(rs("protein")) Then pro = pro + rs("protein") * cc
               lv.SubItems(3) = Round(rs("calories") * cc)
               Dim Calss() As Single
               ReDim Preserve Calss(ll)
               Calss(ll) = GMMs * 100
               ll = ll + 1
               lv.SubItems(4) = Round(rs("fat") * cc) & "  "
               lv.SubItems(5) = Round(rs("carbs") * cc) & "  "
               lv.SubItems(6) = Round(rs("protein") * cc) & "  "
               rs.MoveNext
            Wend
            
            
        rs.MoveFirst
        ll = 0
        Dim NewMeal As Collection
        Set NewMeal = New Collection

        While Not rs.EOF
        
          If IsGrams(ll) Then
            LV1.ListItems.Clear
            Label1.Caption = rs("foodname")
            Set rs2 = DB.OpenRecordset("select * from weight where index=" & rs("abbrev.index") & ";", dbOpenDynaset)
            Dim lvalue As Single
            While Not rs2.EOF
              lvalue = rs2("gm_wgt") / rs2("amount")
              lvalue = Round(Calss(ll) / lvalue * 8) / 8
              Set lv = LV1.ListItems.Add(, , lvalue)
              lv.SubItems(1) = rs2("msre_desc")
              lv.SubItems(2) = rs2("gm_wgt")
              lv.SubItems(3) = rs("calories") / 100 * lvalue / rs2("amount") * rs2("gm_wgt")
              rs2.MoveNext
            Wend
              Set lv = LV1.ListItems.Add(, , Calss(ll))
              lv.SubItems(1) = "Grams"
              lv.SubItems(2) = Calss(ll)
              lv.SubItems(3) = rs("calories") * Calss(ll) / 100
            
            BPause2 = False
            Do
              DoEvents
            Loop Until BPause2
            Set lv = LV1.SelectedItem
            If lv Is Nothing Then Set lv = LV1.ListItems(0)
            
            'rs.Edit
            junk = rs("abbrev.index") & ";;;" & lv.Text & ";;;" & lv.SubItems(1)
            NewMeal.Add junk
            
            'rs.Update
          Else
            NewMeal.Add ""
          End If
            rs.MoveNext
            ll = ll + 1
        Wend
        junk = temp("mealid")
        For i = 1 To NewMeal.Count
         If NewMeal(i) <> "" Then
          Dim junks() As String
          junks = Split(NewMeal(i), ";;;")
          Set rs = DB.OpenRecordset("select * from mealdefinition where mealid=" _
            & junk & " and abbrevid=" & junks(0) & ";", dbOpenDynaset)
          
          rs.Edit
            rs("serving") = junks(1)
            rs("unit") = junks(2)
          rs.Update
         End If
        Next i
        
            temp.MoveNext
'            RTB.Text = junk
            BPause = True
            Do
              DoEvents
            Loop Until BPause
            
 Wend

 
 Next ddi
 
Exit Sub
errhandl:
On Error Resume Next
 MsgBox "Unable to make print preview." & vbCrLf & Err.Description, vbOKOnly, ""
 If DoDebug Then Resume

End Sub


Private Sub Command1_Click()
BPause = True
End Sub

Private Sub Command2_Click()
BPause2 = True
End Sub

Private Sub Form_Click()
LoadMenu
End Sub

Private Sub Form_Load()
Me.Show
DoEvents
Call LoadMenu
End Sub
