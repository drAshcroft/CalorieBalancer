VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMeals 
   Caption         =   "New Meal"
   ClientHeight    =   10440
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   12750
   Icon            =   "frmMeals2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10440
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   13095
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   6960
         ScaleHeight     =   6255
         ScaleWidth      =   3735
         TabIndex        =   17
         Top             =   0
         Width           =   3735
         Begin VB.VScrollBar VScroll1 
            Height          =   6255
            Left            =   3240
            TabIndex        =   19
            Top             =   0
            Width           =   375
         End
         Begin CalorieBalance.RDADisplay RDADisplay1 
            Height          =   13620
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   24024
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enter Ingredients"
         Height          =   495
         Left            =   10800
         TabIndex        =   14
         Top             =   5400
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   -1320
         ScaleHeight     =   3465
         ScaleWidth      =   7185
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   495
            Left            =   2640
            TabIndex        =   13
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   480
            TabIndex        =   12
            Top             =   240
            Width           =   6375
         End
      End
      Begin RichTextLib.RichTextBox RTBDescription 
         Height          =   2970
         Left            =   135
         TabIndex        =   3
         Top             =   6510
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   5239
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMeals2.frx":57E2
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
      Begin VB.TextBox TMealName 
         Height          =   285
         Left            =   105
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.ListBox MealTime 
         Height          =   1230
         ItemData        =   "frmMeals2.frx":5864
         Left            =   3840
         List            =   "frmMeals2.frx":587A
         TabIndex        =   4
         Top             =   720
         Width           =   3015
      End
      Begin RichTextLib.RichTextBox Description 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2160
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7011
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMeals2.frx":58AE
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
      Begin VB.Label Label6 
         Caption         =   $"frmMeals2.frx":5930
         Height          =   3015
         Left            =   7200
         TabIndex        =   15
         Top             =   6480
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   6285
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Meal Time"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Meal Name"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Instructions"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
   End
   Begin CalorieBalance.AdvancedFlex Meal 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12515
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Name and Info"
            Key             =   "N"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ingredients"
            Key             =   "I"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   11040
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuNewFood 
      Caption         =   "New Food or Recipe"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuPOpup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutRow 
         Caption         =   "C&ut Row"
      End
      Begin VB.Menu mnuCopyRow 
         Caption         =   "&Copy Row"
      End
      Begin VB.Menu mnuPasteRow 
         Caption         =   "&Paste Row"
      End
   End
End
Attribute VB_Name = "frmMeals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public PlanID As Long
Public mMealID As Long
Public AllDone As Boolean
Dim AllTotCollection As exCollection
Dim NoUpdate As Boolean
Dim Titles
Dim Changed As Boolean

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function ViewMeal(mealID As Long, CPlanID As Long) As Long
On Error GoTo errhandl
PlanID = CPlanID
mMealID = mealID
Dim temp As Recordset
Dim i As Long, URL As String
Set temp = DB.OpenRecordset( _
"SELECT MealPlanner.Meal, mealplanner.instructions, MealPlanner.MealName, mealplanner.url, MealDefinition.AbbrevID, MealDefinition.Unit, MealDefinition.Serving, MealPlanner.Description, MealPlanner.MealID" & _
" FROM MealPlanner INNER JOIN MealDefinition ON MealPlanner.MealID = MealDefinition.MealID" & _
" WHERE (((MealPlanner.MealID)=" & mealID & "));", dbOpenDynaset)
If temp("mealname") <> "blank" Then
    TMealName.Text = temp("MealName")
End If
URL = temp("url") & ""
'If URL <> "" Then WebBrowser1.Navigate2 URL
Dim junk As String
junk = LCase$(temp("Meal"))
For i = 0 To MealTime.ListCount - 1
  If junk = LCase$(MealTime.List(i)) Then
     MealTime.Selected(i) = True
     Exit For
  End If
Next i
Dim RS As Recordset
Set RS = DB.OpenRecordset("select * from daysinfo where date=#1975-01-01#;", dbOpenDynaset)
While Not RS.EOF
  RS.Delete
  RS.MoveNext
Wend
RS.Close
Set RS = Nothing
Call Meal.OpenDay("1975-01-01")

Description.Text = temp("instructions") & ""
RTBDescription.Text = temp("Description") & ""
i = 1
NoUpdate = True
While Not temp.EOF
   Call Meal.SetRow(i, temp("abbrevid"), temp("serving"), temp("unit"), "")
   i = i + 1
   temp.MoveNext
Wend

NoUpdate = False
Call Meal_RowUpdated


Set temp = Nothing
Changed = False
ViewMeal = 0
Exit Function
errhandl:
MsgBox "Unable to open meal." & vbCrLf & Err.Description, vbOKOnly, ""
ViewMeal = Err.Number
'If Err.Number = "3021" Then Call Unload(Me)

End Function

Private Sub Command1_Click()
Picture1.Visible = False
End Sub

Private Sub Command2_Click()
        Meal.Visible = True
        Frame1.Visible = False
        DoEvents
        Meal.Refresh
End Sub

Private Sub CommandSave_Click()
mnuSave_Click
End Sub

Private Sub Description_Change()


    On Error GoTo Err_Proc
Changed = True
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMeals", "Description_Change", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_Load()
On Error Resume Next
Meal.BackColor = &H800000
Set Meal.Font = frmMain.FlexDiet.Font
'Me.Width = (RecipeNuts.Left + RecipeNuts.Width + 100)
Meal.ShowAsPercent = False
  Meal.AddDataBase DB, CurrentUser.Username, DisplayDate, Nutmaxes
  Meal.SetHeads WatchHeads
  
  ReDim Titles(7)
  Titles(0) = "Suggestions"
  Titles(1) = "Main Course"
  Titles(2) = "Vegetable"
  Titles(3) = "Vegetable"
  Titles(4) = "Fruit"
  Titles(5) = "Fruit"
  Titles(6) = "Grain"
  Titles(7) = "Others"
  'Call Meal.SetYFixed(Titles)
  Meal.SetBackGround Me.BackColor
 
 Call Meal.SetPopUpMenu(MnuPopUp)
 Changed = False
 Call Form_Resize
 If GetSetting(App.Title, "Settings", "MealFirst", True) = True Then
    Call SaveSetting(App.Title, "Settings", "MealFirst", False)
    Picture1.Visible = True
    Label5.Caption = "Making a meal takes two steps.  On this page please enter all the general information about the meal that is needed.  " & vbCrLf _
    & "This requires a descriptive mealname and the default mealtime, then you can add instructions and a nice description if you wish." & vbCrLf & vbCrLf _
    & "Next click on the ingredients tab at the top of the page to enter the foods in the meal.  Last click the save at the top of the page to finish."
    Picture1.ZOrder
 End If
 
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
If Changed And (Not NoQuestions) Then
   Dim ret As VbMsgBoxResult
   ret = MsgBox("Do you wish to save your work?", vbYesNoCancel, "")
   If ret = vbYes Then
     Call mnuSave_Click
     Call frmMain.MakeMealList
   ElseIf ret = vbNo Then
     
     Set RS = DB.OpenRecordset("select * from daysinfo where date=#1975-01-01#;", dbOpenDynaset)
     While Not RS.EOF
       RS.Delete
       RS.MoveNext
     Wend
     RS.Close
     Set RS = Nothing
   
   Else
     Cancel = 1
   End If
   
   Changed = False
End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
   TabStrip1.Move 0, 0, Me.ScaleWidth - 300
   CommandSave.Left = ScaleWidth - CommandSave.Width
   Meal.Move 0, TabStrip1.Height, Me.ScaleWidth - CommandSave.Width, Me.ScaleHeight - 100 - TabStrip1.Height
   Frame1.Move 0, TabStrip1.Height, Me.ScaleWidth, Me.ScaleHeight - 100
  ' WebBrowser1.Width = Me.ScaleWidth - (RecipeNuts.Left + RecipeNuts.Width + 100)
   RTBDescription.Move Description.Left, RTBDescription.Top, Me.ScaleWidth - 100 - Description.Left, Me.ScaleHeight - 200 - (Label4.Top + Label4.Height)
   VScroll1.Move Picture2.Width - VScroll1.Width, 0, VScroll1.Width, Picture2.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
mMealID = 0
AllDone = True
End Sub

Private Sub NameChange()
On Error GoTo errhandl
 Dim NServe As Single, Grams As Single
  Dim Elements, Element, SpanName As String
  
For Each Element In Elements
    On Error Resume Next
    SpanName = Element.Name
    If LCase$(SpanName) = "servingsize" Then
        Element.innerHTML = TMealName.Text & "<br>Serves 1 <br>"
    End If
  Next
errhandl:
End Sub







Private Sub Meal_LostFocus()
Meal.ForceLoseFocus
End Sub

Private Sub Meal_PresentPopup(PopUp As Object, Un As Variant, X As Single, Y As Single, MealRow As Boolean, MealName As String)
On Error Resume Next
Call Me.PopUpMenu(MnuPopUp, , X, Y)
DoEvents
End Sub

Private Sub Meal_RowUpdated()
On Error Resume Next
  Dim NServe As Single
  Dim document, i As Long
  Dim Elements, Element, g
  Dim SpanName As String
  Dim ideals As Recordset
  Dim NewText As String
  
  NServe = 1
  If NServe = 0 Then NServe = 1
  Set AllTotCollection = Meal.GetAllTotals
  Set ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
  
  Call RDADisplay1.DisplayFoodCollection(AllTotCollection, ideals, "1 meal")
  
ideals.Close
Set ideals = Nothing
Set Elements = Nothing
Set document = Nothing
Changed = True
End Sub

Private Sub MealTime_Click()


    On Error GoTo Err_Proc
Changed = True
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMeals", "MealTime_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuCopyRow_Click()
On Error Resume Next
Call Meal.Copy
End Sub

Private Sub mnuCutRow_Click()
On Error Resume Next
If Meal.Visible Then Call Meal.Copy
If Meal.Visible Then Call Meal.DeleteRows
End Sub


Private Sub mnuDelete_Click()
On Error Resume Next
 Call Meal.DeleteRows
' Call Meal.SetYFixed(Titles)
 
End Sub








Private Sub mnuExit_Click()
On Error Resume Next
mMealID = 0
  Meal.Clear
  TMealName.Text = ""
  Description.Text = ""
  RTBDescription.Text = ""
  MealTime.Selected(0) = True
  MealTime.Selected(0) = False
  Meal.Visible = False
  Frame1.Visible = True
  Unload Me
  'Me.Hide
End Sub

Private Sub mnuHelp_Click()
On Error GoTo errhandl
HelpWindowHandle = htmlHelpTopic(Me.hWnd, HelpPath, _
         0, "newHTML/NewRecipeMeal.htm#NewMeal")
Exit Sub
errhandl:
MsgBox "Unable to open help file." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub mnuInsert_Click()
On Error Resume Next
Call Meal.InsertRows
'Call Meal.SetYFixed(Titles)
End Sub

Private Sub mnuNewFood_Click()
   On Error Resume Next
   frmMain.EasyHover1_Click 5
End Sub

Private Sub mnuPasteRow_Click()
On Error Resume Next
Meal.Paste
End Sub

Private Sub mnuSave_Click()
 On Error GoTo errhandl
  Dim i As Long
  Dim MealIndex As Long
  Dim ID As Long, Serving As Single, Unit As String, Adjustable As String
  Dim sTemp As Recordset
  
  Meal.ForceLoseFocus
  If TMealName.Text = "" Then
    MsgBox "Please enter a meal name", vbOKOnly, ""
    TabStrip1.Tabs(1).Selected = True
    TabStrip1_Click
    Exit Sub
  End If
  If SelectedList(MealTime) = -1 Then
    MsgBox "Please Enter the meal time", vbOKOnly, ""
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Exit Sub
  End If
  Dim AServing
  For i = 1 To Meal.Rows - 1
    Call Meal.GetRow(i, ID, Serving, Unit, Adjustable)
    AServing = AServing + Serving
  Next i
  If AServing = 0 Then
    MsgBox "Please enter ingredients into meal.", vbOKOnly, ""
    TabStrip1.Tabs(2).Selected = True
    
    TabStrip1_Click
    Exit Sub
  End If
  
  
  If mMealID = 0 Then
      Set sTemp = DB.OpenRecordset("Select * from MealPlanner;", dbOpenDynaset)
      If sTemp.RecordCount <> 0 Then
        sTemp.MoveLast
        MealIndex = sTemp.Fields("Index") + 1
      Else
        MealIndex = 1
      End If
      sTemp.AddNew
  Else
      Set sTemp = DB.OpenRecordset("Select * from MealPlanner where mealid = " & mMealID & ";", dbOpenDynaset)
      MealIndex = mMealID
      sTemp.Edit
  End If
  
  If PlanID = 0 Then PlanID = 1
  sTemp.Fields("MealName") = TMealName.Text
  sTemp.Fields("MealId") = MealIndex
  sTemp.Fields("PlanId") = PlanID
  sTemp.Fields("Meal") = MealTime.Text
  sTemp("user") = CurrentUser.Username
  
  If Description.Text <> "" Then sTemp.Fields("instructions") = Description.Text Else sTemp("instructions") = " "
  If RTBDescription.Text <> "" Then sTemp.Fields("description") = RTBDescription.Text Else sTemp("description") = " "
  
  Dim Macs() As Single
  sTemp.Fields("calories") = Meal.GetTotalCalories(Macs)
  sTemp("Fat") = Macs(0)
  sTemp("sugar") = Macs(1)
  sTemp("Carbs") = Macs(2)
  sTemp("Protein") = Macs(3)
  sTemp("Fiber") = Macs(4)
   

  
  sTemp.Update
  
  sTemp.Close
  Set sTemp = Nothing
  
  Set sTemp = DB.OpenRecordset("Select * from MealDefinition where mealid=" & MealIndex & ";", dbOpenDynaset)
  If sTemp.EOF = False Then
    While Not sTemp.EOF
     sTemp.Delete
     sTemp.MoveNext
    Wend
  End If
  
  Dim junk As Boolean
  For i = 1 To Meal.Rows - 1
    junk = True
    Call Meal.GetRow(i, ID, Serving, Unit, Adjustable)
    Adjustable = LCase$(Adjustable)
    Dim AJ As Long
    If Adjustable = "yes" Then AJ = -1 Else AJ = 0
    
    If ID = 0 Or ID = -1111 Then junk = False
    If Serving = 0 Then junk = False
    If Unit = "" Then junk = False
    If junk Then
      sTemp.AddNew
      sTemp.Fields("MealID") = MealIndex
      sTemp.Fields("abbrevID") = ID
      sTemp.Fields("unit") = Unit
      sTemp.Fields("serving") = Serving
      sTemp.Fields("Adjustable") = AJ
      sTemp.Update
    End If
  Next i
  sTemp.Close
  Set sTemp = Nothing
  mMealID = 0
  
  Meal.Clear
  TMealName.Text = ""
  Description.Text = ""
  RTBDescription.Text = ""
  MealTime.Selected(0) = True
  MealTime.Selected(0) = False
  
  Meal.Visible = False
  Frame1.Visible = True
  Changed = False
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("select * from daysinfo where date=#1975-01-01#;", dbOpenDynaset)
  While Not RS.EOF
    RS.Delete
    RS.MoveNext
  Wend
  RS.Close
  Set RS = Nothing

  'Call frmMenuPlanner.LoadPlan(frmMenuPlanner.CurrentPlanID)
  Call frmMain.MakeMealList
  frmMain.MealSaved = True
  Unload Me
     
  Exit Sub
errhandl:
  MsgBox "Cannot save food. Please check your entries.", vbOKOnly, ""
End Sub

Private Sub RecipeNuts_DownloadComplete()
On Error Resume Next
  Call Meal_RowUpdated
  Changed = False
End Sub

Private Sub RTBDescription_Change()
On Error Resume Next
Changed = True
End Sub

Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem.Index = 1 Then
        Frame1.Visible = True
        Meal.Visible = False
    Else
        Meal.Visible = True
        Frame1.Visible = False
        DoEvents
        Meal.Refresh
    End If
End Sub

Private Sub TMealName_Change()
On Error Resume Next
Call NameChange
Changed = True
End Sub



Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function

Private Sub VScroll1_Change()
RDADisplay1.Top = (Height - RDADisplay1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
RDADisplay1.Top = (Height - RDADisplay1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
