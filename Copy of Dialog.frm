VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMeals 
   Caption         =   "New Meal"
   ClientHeight    =   10530
   ClientLeft      =   2775
   ClientTop       =   4065
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin CalorieTracker.AdvancedFlex Meal 
      Height          =   7095
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12515
      Yfixed          =   -1  'True
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   13095
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3615
         Left            =   120
         TabIndex        =   0
         Top             =   6240
         Width           =   11775
         ExtentX         =   20770
         ExtentY         =   6376
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
      Begin VB.ListBox MealTime 
         Height          =   1230
         ItemData        =   "Dialog.frx":0000
         Left            =   3840
         List            =   "Dialog.frx":0013
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
      Begin RichTextLib.RichTextBox Description 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7011
         _Version        =   393217
         TextRTF         =   $"Dialog.frx":003F
      End
      Begin VB.TextBox TMealName 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin SHDocVwCtl.WebBrowser RecipeNuts 
         Height          =   6015
         Left            =   7200
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   4695
         ExtentX         =   8281
         ExtentY         =   10610
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
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
      Begin VB.Label Label3 
         Caption         =   "Instructions"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Meal Name"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Meal Time"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuNewFood 
      Caption         =   "New Recipe or Food"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMeals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PlanID As Long
Public mMealID As Long
Dim AllTotCollection As exCollection
Dim NoUpdate As Boolean
Public Sub ViewMeal(MealID As Long, cPlanID As Long)
PlanID = cPlanID
mMealID = MealID
Dim temp As Recordset
Dim I As Long, url As String
Set temp = DB.OpenRecordset( _
"SELECT MealPlanner.Meal, MealPlanner.MealName, mealplanner.url, MealDefinition.AbbrevID, MealDefinition.Unit, MealDefinition.Serving, MealPlanner.Description, MealPlanner.MealID" & _
" FROM MealPlanner INNER JOIN MealDefinition ON MealPlanner.MealID = MealDefinition.MealID" & _
" WHERE (((MealPlanner.MealID)=" & MealID & "));", dbOpenDynaset)
TMealName.Text = temp("MealName")
On Error Resume Next
url = temp("url")
If url <> "" Then WebBrowser1.Navigate2 url
Dim Junk As String
Junk = LCase$(temp("Meal"))
For I = 0 To MealTime.ListCount - 1
  If Junk = LCase$(MealTime.List(I)) Then
     MealTime.Selected(I) = True
     Exit For
  End If
Next I
Description.Text = temp("Description")
I = 1
NoUpdate = True
While Not temp.EOF
   Call Meal.SetRow(I, temp("abbrevid"), temp("serving"), temp("unit"), "")
   I = I + 1
   temp.MoveNext
Wend

NoUpdate = False
Call Meal_RowUpdated


Set temp = Nothing
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()


End Sub





Private Sub Form_Load()
Dim MealMaxes As Calories
Set MealMaxes = New Calories
Call MealMaxes.Init("Average", Date, frmMain.SC)
Call MealMaxes.Update(Date, 0.333)

  Meal.AddDataBase DB, CurrentUser.Username, DisplayDate, MealMaxes
  Dim Heads(3) As String
  Heads(0) = "Calories"
  Heads(1) = "Protein"
  Heads(2) = "Fat"
  Heads(3) = "Carbs"
  Meal.SetHeads Heads
  Dim Titles
  ReDim Titles(7)
  Titles(0) = "Suggestions"
  Titles(1) = "Main Course"
  Titles(2) = "Vegetable"
  Titles(3) = "Vegetable"
  Titles(4) = "Fruit"
  Titles(5) = "Fruit"
  Titles(6) = "Grain"
  Titles(7) = "Others"
  Call Meal.SetYFixed(Titles)
  Meal.SetBackGround Me.BackColor
  RecipeNuts.Navigate2 App.path & "\Resources\daily\NutritionFacts.htm"
  WebBrowser1.Navigate2 App.path & "\resources\help\meal_instructions.htm"
End Sub

Private Sub Form_Resize()
  
   TabStrip1.Move 0, 0, Me.ScaleWidth - 300
   Meal.Move 0, TabStrip1.Height, Me.ScaleWidth, Me.ScaleHeight - 100
   Frame1.Move 0, TabStrip1.Height, Me.ScaleWidth, Me.ScaleHeight - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
mMealID = 0
End Sub

Private Sub NameChange()
 Dim NServe As Single, Grams As Single
  Dim Elements, Element, SpanName As String
  Set Elements = RecipeNuts.Document.getElementsByTagName("Div")
  
For Each Element In Elements
    On Error Resume Next
    SpanName = Element.name
    If LCase$(SpanName) = "servingsize" Then
        Element.innerHTML = TMealName.Text & "<br>Serves 1 <br>"
    End If
  Next
End Sub

Private Sub Meal_RowUpdated()
  Dim NServe As Single
  Dim Document As HTMLDocument, I As Long
  Dim Elements, Element, G As HTMLSpanElement
  Dim SpanName As String
  Dim Ideals As Recordset
  Dim NewText As String
  
  NServe = 1
  If NServe = 0 Then NServe = 1
  Set AllTotCollection = Meal.GetAllTotals
  Set Ideals = DB.OpenRecordset("Select * from ideals where index=1;", dbOpenDynaset)
  Set Document = RecipeNuts.Document

  
  Set Elements = Document.getElementsByTagName("span")
  
For Each Element In Elements
  SpanName = ""
  On Error Resume Next
  SpanName = Element.name
  If SpanName <> "" Then
     NewText = ""
     SpanName = Replace(SpanName, "carbohydrates", "carbs")
     If Right$(SpanName, 3) = "__P" Then
       SpanName = Replace(Left$(SpanName, Len(SpanName) - 3), "_", " ")
       NewText = Round(AllTotCollection(SpanName) / Ideals(SpanName) / NServe * 100)
     Else
       SpanName = Replace(SpanName, "_", " ")
       NewText = Round(AllTotCollection(SpanName) / NServe, 1)
     End If
     If NewText = "" Then NewText = "0"
     Element.innerText = " " & NewText
  End If
Next
'Call NameChange
Debug.Print AllTotCollection("Total grams") / NServe
Ideals.Close
Set Ideals = Nothing
Set Elements = Nothing
Set Document = Nothing
End Sub

Private Sub mnuExit_Click()
mMealID = 0
  Meal.Clear
  TMealName.Text = ""
  Description.Text = ""
  MealTime.Selected(0) = True
  MealTime.Selected(0) = False
  Meal.Visible = False
  Frame1.Visible = True
  
  Me.Hide
End Sub

Private Sub mnuHelp_Click()
Call HTMLHelp(frmMain.hWnd, App.HelpFile, HH_DISPLAY_TOPIC, ByVal "how_to_make_a_meal.htm")
End Sub

Private Sub mnuNewFood_Click()
   Dim newfood As FNewFood
   Set newfood = New FNewFood
   newfood.Show vbModal, Me
   Unload newfood
   Set newfood = Nothing
   
End Sub

Private Sub mnuSave_Click()
  If TMealName.Text = "" Then
    MsgBox "Please enter a meal name", vbOKOnly, ""
  End If
  If SelectedList(MealTime) = -1 Then
    MsgBox "Please Enter the meal time", vbOKOnly, ""
    Exit Sub
  End If
  Dim I As Long
  Dim MealIndex As Long
  Dim ID As Long, serving As Single, Unit As String, Adjustable As String
  Dim sTemp As Recordset
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
  

  sTemp.Fields("MealName") = TMealName.Text
  sTemp.Fields("MealId") = MealIndex
  sTemp.Fields("PlanId") = PlanID
  sTemp.Fields("Meal") = MealTime.Text
  
  If Description.Text <> "" Then sTemp.Fields("description") = Description.Text Else sTemp("description") = " "
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
  
  Dim Junk As Boolean
  For I = 1 To Meal.Rows - 1
    Junk = True
    Call Meal.GetRow(I, ID, serving, Unit, Adjustable)
    Adjustable = LCase$(Adjustable)
    Dim AJ As Long
    If Adjustable = "yes" Then AJ = -1 Else AJ = 0
    
    If ID = 0 Or ID = -1111 Then Junk = False
    If serving = 0 Then Junk = False
    If Unit = "" Then Junk = False
    If Junk Then
      sTemp.AddNew
      sTemp.Fields("MealID") = MealIndex
      sTemp.Fields("abbrevID") = ID
      sTemp.Fields("unit") = Unit
      sTemp.Fields("serving") = serving
      sTemp.Fields("Adjustable") = AJ
      sTemp.Update
    End If
  Next I
  sTemp.Close
  Set sTemp = Nothing
  mMealID = 0
  
  Meal.Clear
  TMealName.Text = ""
  Description.Text = ""
  MealTime.Selected(0) = True
  MealTime.Selected(0) = False
  
  Meal.Visible = False
  Frame1.Visible = True
  Me.Hide
End Sub

Private Sub RecipeNuts_DownloadComplete()
  Call Meal_RowUpdated
End Sub

Private Sub TabStrip1_Click()
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
Call NameChange
End Sub
