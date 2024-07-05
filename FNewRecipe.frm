VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FNewRecipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Recipe"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   900
   ClientWidth     =   11670
   Icon            =   "FNewRecipe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10320
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete Recipe"
      Height          =   495
      Left            =   10320
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   10320
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8175
      LargeChange     =   100
      Left            =   9720
      Max             =   1000
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   13620
      Left            =   120
      ScaleHeight     =   13620
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4095
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   3975
         Begin VB.Label Label9 
            Caption         =   $"FNewRecipe.frx":57E2
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   3975
            WordWrap        =   -1  'True
         End
      End
      Begin CalorieBalance.AutoComplete Foodname 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
         font            =   "FNewRecipe.frx":5956
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   9615
         Begin VB.TextBox NumberOfServings 
            Height          =   285
            Left            =   0
            TabIndex        =   7
            Text            =   "1"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TServeAmount 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox TUnit 
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ListBox FoodGroup 
            Height          =   1035
            ItemData        =   "FNewRecipe.frx":5982
            Left            =   120
            List            =   "FNewRecipe.frx":599B
            TabIndex        =   4
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox TGrams 
            Height          =   285
            Left            =   3600
            TabIndex        =   3
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Number of Servings in Recipe:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   5295
         End
         Begin VB.Label Label4 
            Caption         =   "Serving Size:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Unit"
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Foodgroup:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label label25 
            Caption         =   "Grams Per Serving"
            Height          =   255
            Left            =   3600
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
      End
      Begin CalorieBalance.AdvancedFlex Ingred 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4320
         Width           =   9015
         _extentx        =   13150
         _extenty        =   7223
         backcolor       =   8388608
         font            =   "FNewRecipe.frx":5A20
      End
      Begin RichTextLib.RichTextBox Instructions 
         Height          =   2775
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   10680
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4895
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"FNewRecipe.frx":5A4C
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
      Begin VB.Label Label3 
         Caption         =   "Recipe Instructions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   10320
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Ingredients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Recipe Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut Row"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Row"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste "
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save"
   End
   Begin VB.Menu mnuRename 
      Caption         =   "&Rename"
   End
   Begin VB.Menu mnuDeleteItem 
      Caption         =   "&Delete Item"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu munHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "FNewRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit

  Dim document
  Dim Nutrients
  Dim General
  Dim Elements, Element
Dim AbbrevID As Long
'Default Property Values:
'Property Variables:
Dim m_RecipeBound As Boolean
Dim FormLoadRedo As Boolean
'Dim Undivided As Collection
 
Dim AllTotCollection As exCollection
Dim Changed As Boolean
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub DisplayRecipe(AbbrevID As Long)

On Error GoTo errhandl
If AbbrevID = -1 Then
 
 
Else
 Foodname.SelectedID = AbbrevID
 Call FoodName_ItemSelected(AbbrevID)
End If
 Changed = False
 Exit Sub
errhandl:
 MsgBox "Cannot display this recipe." & vbCrLf & Err.Description, vbOKOnly, ""
End Sub
Public Sub ShowRecipeUpLoad(AbbrevID As Long)
On Error GoTo errhandl
 'frmMain
 
 Call DisplayRecipe(AbbrevID)

 Changed = False
 
 Exit Sub
errhandl:
 MsgBox "Cannot display this recipe for upload." & vbCrLf & Err.Description, vbOKOnly, ""
 
End Sub

Private Sub CSend_Click()
On Error Resume Next

DoEvents
Call ShowUpLoad
Foodname.RecipeOnly = False
'Unload Me
End Sub

Private Sub CDelete_Click()
mnuDeleteItem_Click
End Sub

Private Sub Command1_Click()


    On Error GoTo Err_Proc
  Call ClearGrid
  
  Unload Me
   'vbModal, frmMain
  
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "FNewFood", "Command1_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub CSave_Click()
  Call mnuSave_Click
End Sub

Private Sub FoodGroup_Click()
On Error Resume Next
Changed = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Changed And Not NoQuestions Then
  Dim ret As VbMsgBoxResult
  ret = MsgBox("Do you wish to save your work?", vbYesNoCancel, "")
  If ret = vbYes Then
     mnuSave_Click
     Cancel = 1
  ElseIf ret = vbCancel Then
     Cancel = 1
  Else
     Call ClearGrid
  End If
End If

End Sub




Private Sub Form_Resize()
VScroll1.Height = ScaleHeight
End Sub

Private Sub Ingred_LostFocus()
Ingred.ForceLoseFocus
End Sub

Private Sub Ingred_PresentPopup(PopUp As Object, Un As Variant, X As Single, Y As Single, MealRow As Boolean, MealName As String)
On Error Resume Next
PopUpMenu PopUp, , X, Y
End Sub

Private Sub mnuPrint_Click()
On Error GoTo errhandl

Call ShowSummary(Foodname.SelectedID)
Call OpenURL(App.path & "\Resources\temp\RecipeView.htm")


 Exit Sub
errhandl:
 MsgBox "Cannot print." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub



Private Function ValidateRecipe() As Boolean
On Error Resume Next
Dim Error As String

If Val(NumberOfServings.Text) <= 0 Then Error = "Please enter the number of servings in recipe"
If Ingred.Rows < 1 Then Error = "Please enter ingrediants for recipe"
If Val(TServeAmount.Text) <= 0 Then Error = "Please Enter a serving amount"
If TUnit.Text = "" Then Error = "Please enter a serving unit"
'If Val(TGrams.Text) = 0 Then Error = "Please enter or estimate number of grams per serving"

If Error <> "" Then MsgBox Error, vbOKOnly, ""
ValidateRecipe = (Error = "") 'Or ValidateFood
End Function

Private Function ShowSummary(Optional PrintIT As Boolean = True) As String
On Error GoTo errhandl
Dim NServe As Long, i As Long, ff As Long

Dim OutString As String

'check to make sure that everything is lined up
  If Not ValidateRecipe Then
     
     Exit Function
  End If

  Dim Nutrition As String
 
   
  OutString = "<Html ><body>"
  OutString = OutString & "<DIV style=""FLOAT: right; MARGIN: 20px 0px 10px 10px; WIDTH: 200px"">"
  OutString = OutString & Nutrition
  OutString = OutString & "</div>"
  
  OutString = OutString & "<h2>" & Foodname.Text & "</h2>"
  OutString = OutString & "<b>" & NumberOfServings.Text & " Servings</b>"
  
  OutString = OutString & "<ul>"
  Dim FoodJ As String
  Dim JunkB As Boolean
  Dim ID As Long, Serving As Single, Unit As String
  
  For i = 1 To Ingred.Rows - 1
    JunkB = True
    Call Ingred.GetRow(i, ID, Serving, Unit, "", , FoodJ)
    
    'check if all the ingredients are valid
    If ID = 0 Or ID = -1111 Then JunkB = False
    If Serving = 0 Then JunkB = False
    If Unit = "" Then JunkB = False
    
    If JunkB Then
       OutString = OutString & "<li>" & ConvertDecimalToFraction(Serving)
       OutString = OutString & " " & Unit & " " & FoodJ & "</li>"
    End If
  Next i
  OutString = OutString & "</ul><br><br>"
  OutString = OutString & "<h3>Instructions</h3>"
  OutString = OutString & Replace(Instructions.Text, vbCrLf, "<br>")
  
  If PrintIT Then
      ff = FreeFile
      Open App.path & "\Resources\temp\RecipeView.htm" For Output As #ff
        Print #ff, OutString
      Close #ff
     
    
  End If
  ShowSummary = OutString
  Exit Function
errhandl:
 MsgBox "Cannot build summary.  Please check recipe." & vbCrLf & Err.Description, vbOKOnly, ""
End Function



Public Sub CopyRow()
On Error Resume Next
   Ingred.Copy
   
End Sub
Public Sub CutRow()
On Error Resume Next
  Ingred.Copy
  Ingred.DeleteRows
End Sub

Public Sub InsertRow()
On Error Resume Next
  Ingred.InsertRows
End Sub

Public Sub DeleteRow()
On Error Resume Next
  Ingred.DeleteRows
End Sub

Public Sub Paste()
On Error Resume Next
  Ingred.Paste
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Foodname.RecipeOnly = False
    Call ClearGrid
   
End Sub

Private Sub Ingred_RowUpdated()
On Error Resume Next
  Changed = True
  Dim NServe As Single
  Dim document, i As Long
  Dim Elements, Element, g
  Dim SpanName As String
  Dim ideals As Recordset
  Dim NewText As String
  
  NServe = Val(NumberOfServings.Text)
  If NServe = 0 Then NServe = 1
  Set AllTotCollection = Ingred.GetAllTotals
  Set ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)

  ChangeServingAmount

  Set Elements = document.getElementsByTagName("span")
  
For Each Element In Elements
  SpanName = ""
  On Error Resume Next
  SpanName = Element.Name
  If SpanName <> "" Then
     NewText = ""
     SpanName = Replace(SpanName, "carbohydrates", "carbs")
     If Right$(SpanName, 3) = "__P" Then
       SpanName = Replace(Left$(SpanName, Len(SpanName) - 3), "_", " ")
       NewText = Round(AllTotCollection(SpanName) / ideals(SpanName) / NServe * 100)
     Else
       SpanName = Replace(SpanName, "_", " ")
       NewText = Round(AllTotCollection(SpanName) / NServe, 1)
     End If
     If NewText = "" Then NewText = "0"
     Element.innerText = " " & NewText
  End If
Next
'TGrams.Text = AllTotCollection("Total grams") / NServe
ideals.Close
Set ideals = Nothing
Set Elements = Nothing
Set document = Nothing
End Sub

Private Sub mnuDeleteItem_Click()
On Error Resume Next
    Call DeleteFood
    ClearGrid
    Changed = False
    Unload Me
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    Call ClearGrid
    Unload Me
End Sub


Private Sub mnuRename_Click()
On Error Resume Next
    Call Rename
End Sub

Private Sub munHelp_Click()
''dim helpwindowhandle As Long
On Error GoTo errhandl
HelpWindowHandle = htmlHelpTopic(frmMain.hWnd, HelpPath, _
         0, "newHTML/NewRecipeMeal.htm#NewFood")
 Exit Sub
errhandl:
 MsgBox "Cannot find help file." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub NumberOfServings_Change()
On Error Resume Next
Call Ingred_RowUpdated
End Sub
Private Sub ChangeServingAmount()
On Error Resume Next
  Changed = True
  Dim NServe As Single, Grams As Single
  Dim Elements, Element, SpanName As String
 
  NServe = Val(NumberOfServings.Text)
  If NServe <> 0 Then
     Grams = Round(Val(TGrams.Text) / NServe, 1)
  Else
     Grams = 100
  End If
  For Each Element In Elements
    On Error Resume Next
    SpanName = Element.Name
    If LCase$(SpanName) = "servingsize" Then
        Element.innerHTML = Foodname.Text & "<br>Serves " _
                 & NumberOfServings.Text & "<br>Serving Size<br>" & TServeAmount.Text & " " & TUnit.Text & "<br>" & Grams & " grams"
    End If
  Next
  Set Elements = Nothing
  Set Element = Nothing
End Sub

Private Sub RecipeNuts_DocumentComplete(ByVal pDisp As Object, URL As Variant)


    On Error GoTo Err_Proc
If FormLoadRedo Then
   Call Ingred_RowUpdated
  
End If
 FormLoadRedo = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "FNewFood", "RecipeNuts_DocumentComplete", Err.Description
    Resume Exit_Proc


End Sub


Private Sub TServeAmount_Change()
On Error Resume Next
  ChangeServingAmount
End Sub

Private Sub TUnit_Change()
On Error Resume Next
  ChangeServingAmount
End Sub
Private Function GetSelected(L As ListBox) As Long
On Error Resume Next
Dim i As Long
  For i = 0 To L.ListCount - 1
    If L.Selected(i) Then
       GetSelected = L.ItemData(i)
       Exit Function
    End If
  Next
  GetSelected = -1
End Function
Private Function FormatLine(ID As Long, Unit As String, Serving As Single) As String
On Error Resume Next
              Dim temp2 As Recordset
              Dim junk As String, Food As String
              Set temp2 = DB.OpenRecordset("Select * from abbrev " & _
              " where index = " & ID & ";", dbOpenDynaset)
              
              Food = temp2.Fields("foodname")
              If ID >= 0 Then
                   Dim GS As Single
                  ' Set temp3 = DB.OpenRecordset("SELECT *" _
                  '            & " From weight " _
                  '            & " WHERE ((index=" & ID & ") and " _
                  '            & "(msre_desc = '" & Unit & "'));", dbOpenDynaset)
                  '            On Error Resume Next
                  GS = Module1.TranslateUnitToGrams(ID, Unit) 'temp3.Fields("gm_wgt").Value / temp3.Fields("amount").Value / 100 * serving
             
                  junk = ConvertDecimalToFraction(Serving) & " <!-- -->" & Unit & "<!-- --> " & Food
                  junk = junk & "<!-- " & Round(temp2("Calories") * GS)
                  junk = junk & " " & Round(temp2("Sugar") * GS, 1)
                  junk = junk & " " & Round(temp2("Fiber") * GS, 1)
                  junk = junk & " " & Round(temp2("Carbs") * GS, 1)
                  junk = junk & " " & Round(temp2("Fat") * GS, 1)
                  junk = junk & " " & Round(temp2("Protein") * GS, 1)
                  junk = junk & " " & Round(GS * 100, 2) & "-->"
              End If
              Set temp2 = Nothing
              FormatLine = junk
End Function
Public Sub LoadUploader(Browser)
On Error Resume Next
    
    Dim document
    Dim e, i As Long
    Set document = Browser.document.Forms("form1")
    If document Is Nothing Then Exit Sub
    Set e = document.Elements

  Dim Nutrition As String, OutString As String
  
  e("Nutrients").innerText = Nutrition
  e("name").Value = Foodname.Text
  e("NumberOfServings").Value = NumberOfServings.Text
  e("author").Value = CurrentUser.Username
  
  
  OutString = "<ul>"
  Dim FoodJ As String
  Dim JunkB As Boolean
  Dim ID As Long, Serving As Single, Unit As String
  
  For i = 1 To Ingred.Rows - 1
    JunkB = True
    Call Ingred.GetRow(i, ID, Serving, Unit, "", , FoodJ)
    
    'check if all the ingredients are valid
    If ID = 0 Or ID = -1111 Then JunkB = False
    If Serving = 0 Then JunkB = False
    If Unit = "" Then JunkB = False
    If JunkB Then
       OutString = OutString & "<li>" & FormatLine(ID, Unit, Serving) & "</li>" & vbCrLf
    End If
  Next i
  OutString = OutString & "</ul><br>"
  
  e("Ingredients").Value = OutString
  e("Instructions").innerText = Instructions.Text
  
  Set e = Nothing
 ' Document.submit
 ' DoEvents
 ' MsgBox "Recipe has been submitted.", vbOKOnly, ""
'  Uploader.Move 0, 0, ScaleWidth, ScaleHeight
'  Fuploader.Visible = False
'  Uploader.Visible = False
'  Changed = False
'  Unload Me
  Exit Sub
errhandl:
  MsgBox "Cannot upload recipe." & vbCrLf & Err.Description, vbOKOnly, ""
   
  'Resume
End Sub









Private Sub Form_Load()
Label9.Caption = "To enter a recipe you must give it a name and then select serving information.  For example, if a bread recipe makes 10 slices of bread, then the serving size would be 1 slice and the number of servings would be 10. " & vbCrLf _
   & "If you do not know how many grams are in a serving enter 100." & vbCrLf _
   & "Choose the largest ingredient for the food group.  For example, bread would belong to grains." & vbCrLf _
   & "Then enter the ingredients in the ingredient grid and finish with instructions for the recipe below."

On Error Resume Next
Set Ingred.Font = frmMain.FlexDiet.Font
Ingred.Font = frmMain.FlexDiet.Font

Dim RS As Recordset
Set RS = DB.OpenRecordset("select * from daysinfo where date=#1975-01-01#;", dbOpenDynaset)
While Not RS.EOF
  RS.Delete
  RS.MoveNext
Wend
RS.Close
Set RS = Nothing

  
    
    Ingred.ZOrder
    Instructions.ZOrder
    


Foodname.Reset
Foodname.AddDataBase DB
Foodname.ZOrder
Foodname.Visible = True
Foodname.SuggestedHeight = Foodname.Height
Ingred.AddDataBase DB, CurrentUser.Username, DisplayDate, Nutmaxes
Call Ingred.SetHeads(WatchHeads)
Call Ingred.SetPopUpMenu(MnuPopUp)


FormLoadRedo = True

FoodGroup.Clear
FoodGroup.AddItem "Egg Products"
FoodGroup.AddItem "Dairy Products"
FoodGroup.AddItem "Spices and Herbs"
FoodGroup.AddItem "Baby Foods"
FoodGroup.AddItem "Fats and Oils"
FoodGroup.AddItem "Poultry Products"
FoodGroup.AddItem "Soups, Sauces, and Gravies"
FoodGroup.AddItem "Sausages and Luncheon Meats"
FoodGroup.AddItem "Breakfast Cereals"
FoodGroup.AddItem "Fruits and Fruit Juices"
FoodGroup.AddItem "Pork Products"
FoodGroup.AddItem "Vegetables and Vegetable Products"
FoodGroup.AddItem "Nut and Seed Products"
FoodGroup.AddItem "Beef Products"
FoodGroup.AddItem "Beverages"
FoodGroup.AddItem "Finfish and Shellfish Products"
FoodGroup.AddItem "Legumes and Legume Products"
FoodGroup.AddItem "Lamb, Veal, and Game Products"
FoodGroup.AddItem "Baked Products"
FoodGroup.AddItem "Sweets"
FoodGroup.AddItem "Cereal Grains And Pasta"
FoodGroup.AddItem "Fast Foods"
FoodGroup.AddItem "Meals, Entrees, and Sidedishes"
FoodGroup.AddItem "Snacks"
FoodGroup.AddItem "Ethnic Foods"

FoodGroup.ItemData(0) = "0100"
FoodGroup.ItemData(1) = "0200"
FoodGroup.ItemData(2) = "0300"
FoodGroup.ItemData(3) = "0400"
FoodGroup.ItemData(4) = "0500"
FoodGroup.ItemData(5) = "0600"
FoodGroup.ItemData(6) = "0700"
FoodGroup.ItemData(7) = "0800"
FoodGroup.ItemData(8) = "0900"
FoodGroup.ItemData(9) = "1000"
FoodGroup.ItemData(10) = "1100"
FoodGroup.ItemData(11) = "1200"
FoodGroup.ItemData(12) = "1300"
FoodGroup.ItemData(13) = "1400"
FoodGroup.ItemData(14) = "1500"
FoodGroup.ItemData(15) = "1600"
FoodGroup.ItemData(16) = "1700"
FoodGroup.ItemData(17) = "1800"
FoodGroup.ItemData(18) = "1900"
FoodGroup.ItemData(19) = "2000"
FoodGroup.ItemData(20) = "2100"
FoodGroup.ItemData(21) = "2200"
FoodGroup.ItemData(22) = "2500"
FoodGroup.ItemData(23) = "3500"
 
End Sub


Private Sub mnuCopy_Click()
On Error Resume Next
  CopyRow
End Sub

Private Sub mnuCut_Click()
On Error Resume Next
  CutRow
End Sub

Private Sub mnuDelete_Click()
On Error Resume Next
 Ingred.DeleteRows
End Sub

Private Sub mnuInsert_Click()
On Error Resume Next
  InsertRow
End Sub


Private Sub mnuPaste_Click()
On Error Resume Next
  Paste
End Sub

Private Sub FoodName_ExitFocus()
On Error Resume Next
  Call Foodname.CloseBox
End Sub


Private Sub FoodName_ItemSelected(SelectedID As Long)
'On Error GoTo errhandl
Dim temp2 As Recordset
Dim temp As Collection
Dim i As Long, BC As Single
Dim junkI As Long
Dim ideals As Recordset

Set temp = Foodname.GetNutrients
If temp("ndb_no") = -100 Then
  RecipeBound True
  If Not ShowRecipe(SelectedID) Then
  
    GoTo mistake
  End If
  Call Foodname.CloseBox
  Exit Sub
End If
mistake:
Call Foodname.CloseBox
On Error Resume Next
     
     Set Nutrients = document.Forms("Nutrients")
     Set General = document.Forms("theGenerals")


Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & SelectedID & "));", dbOpenDynaset)
Set ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)

If Not (temp2.EOF And temp2.BOF) Then
    General.Unit.Value = temp2.Fields("msre_desc").Value
    General.Amount.Value = ConvertDecimalToFraction(temp2.Fields("Amount").Value)
    General.Grams.Value = temp2("gm_wgt")
    BC = temp2.Fields("gm_wgt").Value / 100
    Set temp2 = Nothing
Else
   BC = 1
End If

  Dim junk As String, Switch As Boolean, junk2
  Dim Element, Elements
  Set Elements = Nutrients.Elements
  Switch = False
  On Error Resume Next
  For Each Element In Elements
        junk2 = 0
        junk = CleanList(Element.Name)
        If LCase$(junk) = "vitamin a" Then Switch = True
        junk2 = temp(junk)
        
        If Switch Then
           Element.Value = Round(junk2 * BC / ideals(junk) * 100, 1)
        Else
           Element.Value = Round(junk2 * BC, 2)
        End If
        Set Element = Nothing
  Next
  Set Elements = Nothing
  
  Dim junks As String
junks = temp("Foodgroup")

General.FoodGroup.Value = junks
'General.FoodGroup.Options(junkI).Selected = True
ideals.Close
errhandl:
Set ideals = Nothing
Set temp = Nothing

End Sub

Private Sub FoodName_NoneSelected()
On Error Resume Next
Dim B As Long
'B = RecipeCHK.Value
'Call ClearGrid(True)
'RecipeCHK.Value = B
Foodname.CloseBox
End Sub


Private Sub mnuSave_Click()
    On Error GoTo errhandl
    Dim j
    
        
        j = 2 ' GetSetting(App.Title, "settings", "SaveIntenetRecipe", 1)
        If j <> 2 Then
           frmSaveInternet.AbbrevID = AbbrevID
           frmSaveInternet.SaveFood = False
           Set frmSaveInternet.HostForm = Me
           Me.Hide
           frmSaveInternet.Show vbModal, Me
        End If
        If Not SaveRecipe Then
           Me.Show
           
           Exit Sub
        End If
 
    Changed = False
    Unload Me
    Exit Sub
errhandl:
   MsgBox "Cannot save this recipe." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub


Private Sub ShowUpLoad()
'On Error Resume Next
  If Not ValidateRecipe Then
     Exit Sub
  End If


 
   
End Sub

Private Function SaveRecipe() As Boolean
On Error GoTo errhandl

  Dim NServe As Long, i As Long
  SaveRecipe = True
  If Not ValidateRecipe Then
     SaveRecipe = False
     Exit Function
  End If
  NServe = NumberOfServings.Text
  Dim sTemp As Recordset
  Dim RecipesIndex As Recordset
  Dim AbbrevID As Long, RecipeID As Long
  'check if the recipes is already in the database
  Set RecipesIndex = DB.OpenRecordset("Select * from recipesindex where recipename = '" _
                             & Replace(Foodname.Text, "'", "''") & "';", dbOpenDynaset)
  If RecipesIndex.EOF Then
    'get the max
    Set sTemp = DB.OpenRecordset("Select max(index) as MAXIT from abbrev;", dbOpenDynaset)
    AbbrevID = sTemp("Maxit") + 1
    sTemp.Close
    Set sTemp = Nothing
    RecipesIndex.AddNew
    RecipesIndex("abbrevid") = AbbrevID
    Set sTemp = DB.OpenRecordset("Select max(recipeid) as MAXIT from recipesindex;", dbOpenDynaset)
    On Error Resume Next
    RecipeID = sTemp("Maxit") + 1
    If Err.Number = 94 Then
      RecipeID = 1
    End If
    On Error GoTo errhandl
    sTemp.Close
    Set sTemp = Nothing
  
    RecipesIndex("Recipeid") = RecipeID
  Else
    AbbrevID = RecipesIndex("abbrevid")
    RecipeID = RecipesIndex("Recipeid")
    RecipesIndex.Edit
  End If
  'set up this information
  RecipesIndex.Fields("RecipeName") = Foodname.Text
  If Instructions.Text = "" Then
    RecipesIndex.Fields("RecipeInstructions") = " "
  Else
    RecipesIndex.Fields("RecipeInstructions") = Replace(Instructions.TextRTF, vbCrLf, "")
  End If
  RecipesIndex.Fields("NumberOfServings") = NServe
  RecipesIndex.Update
  RecipesIndex.Close
  Set RecipesIndex = Nothing
  
  'now add the information to the main list
  Call SaveToFoodList(AbbrevID, NServe, RecipeID)
  
  'now save in the ingredients
  Set sTemp = DB.OpenRecordset("SELECT * FROM Recipes where recipeid=" & RecipeID & ";", dbOpenDynaset)
  While Not sTemp.EOF
     sTemp.Delete
     sTemp.MoveNext
  Wend
  Dim j As Long, JunkB As Boolean
  Dim ID As Long, Serving As Single, Unit As String
  For i = 1 To Ingred.Rows - 1
    JunkB = True
    Call Ingred.GetRow(i, ID, Serving, Unit, "")
    If ID = 0 Or ID = -1111 Then JunkB = False
    If Serving = 0 Then JunkB = False
    If Unit = "" Then JunkB = False
    If JunkB Then
      sTemp.AddNew
      sTemp.Fields("RecipeID") = RecipeID
      sTemp.Fields("itemID") = ID
      sTemp.Fields("unit") = Unit
      sTemp.Fields("servings") = Serving
      sTemp.Update
    End If
  Next i
  Set sTemp = Nothing
  
  Call ClearGrid
  
  MsgBox "Food Saved Succesfully", vbOKOnly, ""
  Exit Function
errhandl:
  MsgBox "Unable to save food.  Please check all the boxes.", vbOKOnly, ""
 ' Resume
End Function
Private Sub SaveToFoodList(AbbrevID As Long, NServings As Long, RecipeID As Long)
 'On Error Resume Next
  
    
    Dim cc As Single
    Dim Grams As Single

    Grams = Val(TGrams.Text)
    If Grams = 0 Then Grams = 100
    cc = 100 / Grams
  
  Dim i As Long
  
  
  Dim sTemp As Recordset, temp As Recordset
  'On Error Resume Next
  Set sTemp = DB.OpenRecordset("SELECT * FROM ABBREV where index = " & AbbrevID & ";", dbOpenDynaset)
  Set temp = DB.OpenRecordset("SELECT * from Weight where index = " & AbbrevID & ";", dbOpenDynaset)
    
  If sTemp.EOF <> True Then sTemp.Delete
  While temp.EOF = False
     temp.Delete
     temp.MoveNext
  Wend
    
    sTemp.AddNew
    temp.AddNew
    temp("Gm_Wgt") = Grams
    temp("index") = AbbrevID
    temp("Msre_Desc") = TUnit.Text
    temp("Amount") = Module1.ConvertFractionToDecimal(TServeAmount.Text)
    temp.Update
    
    sTemp.Fields("index") = AbbrevID
    sTemp.Fields("NDB_No") = -100
    
    Dim eValue, eName As String
    For i = 0 To sTemp.Fields.Count - 1
       eName = LCase$(sTemp.Fields(i).Name)
       If eName <> "ndb_no" And eName <> "index" And eName <> "foodname" And eName <> "foodgroup" Then
          sTemp.Fields(eName) = AllTotCollection(eName) / NServings * cc
       End If
    Next
    
    Dim FoodgroupS As String
    For i = 0 To FoodGroup.ListCount - 1
      If FoodGroup.Selected(i) Then
         FoodgroupS = FoodGroup.ItemData(i)
         Exit For
      End If
    Next i
   
    If Len(FoodgroupS) < 4 Then
      While Len(FoodgroupS) < 4
         FoodgroupS = "0" & FoodgroupS
      Wend
    End If
    sTemp.Fields("foodgroup") = FoodgroupS
    sTemp.Fields("FoodName") = Foodname.Text
    
    sTemp.Fields("Usage") = 10
    sTemp.Update
    Set sTemp = Nothing
    Set temp = Nothing
  
End Sub



Public Function ShowRecipe(AbbrevID As Long) As Boolean
On Error GoTo errhandl
  Dim i As Long
  Dim RecipeIndex  As Recordset
  Dim Recipes As Recordset
  Dim Weight As Recordset
  Dim Abbrev As Recordset
  Set Abbrev = DB.OpenRecordset("Select * from abbrev where (index = " & AbbrevID & ");", dbOpenDynaset)
  Set RecipeIndex = DB.OpenRecordset("Select * from Recipesindex where (abbrevid = " & AbbrevID & ");", dbOpenDynaset)
  
   If RecipeIndex.EOF Then
    Set Abbrev = Nothing
    Set RecipeIndex = Nothing
    ShowRecipe = False
    Exit Function
  Else
    ShowRecipe = True
  End If
  
  Set Recipes = DB.OpenRecordset("SELECT * from recipes where (recipeid = " & RecipeIndex.Fields("RecipeID") & ");", dbOpenDynaset)
  Set Weight = DB.OpenRecordset("Select * from weight where (index = " & AbbrevID & ");", dbOpenDynaset)
  
  Ingred.Clear
     
  Dim Row As Long, ID As Long, Unit As String, Amount As Single
  Row = 1
  While Not Recipes.EOF
     ID = Recipes.Fields("ItemId")
     Unit = Recipes.Fields("unit")
     Amount = Recipes.Fields("Servings")
     If Ingred.SetRow(Row, ID, Amount, Unit, "") Then Row = Row + 1
     Recipes.MoveNext
  Wend
  Call Ingred.SetRow(Row + 1, 0, 0, "", "")

  Instructions.TextRTF = RecipeIndex.Fields("RecipeInstructions")
  
 ' On Error
  TUnit.Text = Weight.Fields("Msre_desc")
  TServeAmount.Text = Module1.ConvertDecimalToFraction(Weight.Fields("amount"))
  On Error Resume Next
  Dim FoodgroupS As Long
  FoodgroupS = Val(Abbrev.Fields("Foodgroup"))
  For i = 0 To FoodGroup.ListCount - 1
     If FoodGroup.ItemData(i) = FoodgroupS Then
        FoodGroup.Selected(i) = True
     End If
  Next i
  Foodname.Text = Abbrev.Fields("FoodName")
  NumberOfServings.Text = RecipeIndex.Fields("NumberOfServings")
  TGrams.Text = Weight("gm_wgt")
  If Val(TGrams.Text) = 100 Then TGrams.Text = ""
  On Error Resume Next
  Weight.Close
  Abbrev.Close
  Recipes.Close
  RecipeIndex.Close
  Set Weight = Nothing
  Set Abbrev = Nothing
  Set Recipes = Nothing
  Set RecipeIndex = Nothing
  Call Ingred_RowUpdated
'  SetGrid Nutrients
  Exit Function
errhandl:
  MsgBox "Unable to show recipe" & vbCrLf & Err.Description, vbOKOnly, ""
  
End Function

Public Sub ClearGrid(Optional SkipFoodname As Boolean = False)
   On Error Resume Next
   Dim RS As Recordset
   Set RS = DB.OpenRecordset("select * from daysinfo where date=#1975-01-01#;", dbOpenDynaset)
   While Not RS.EOF
     RS.Delete
     RS.MoveNext
   Wend
   RS.Close
   Set RS = Nothing
       
     Set Nutrients = document.Forms("Nutrients")
     Set General = document.Forms("theGenerals")


   
   If Not SkipFoodname Then Foodname.Text = ""
   General.FoodGroup.Options(0).Selected = True
   FoodGroup.Selected(0) = True
   TServeAmount.Text = ""
   TUnit.Text = ""
   TGrams.Text = 100
   NumberOfServings.Text = 1
  
  
    Ingred.Clear
    Instructions.Text = ""
   
    Changed = False
End Sub
Private Function CleanList(ListItem As String) As String
  On Error Resume Next
  Dim junk As String
  junk = Trim$(ListItem)
  junk = Replace(junk, "__", "-")
  junk = Replace(junk, "_", " ")
  junk = Replace(junk, "Total", "", , , vbTextCompare)
  If junk = "Dietary fiber" Then junk = "Fiber"
  If LCase$(junk) = "carbohydrate" Then junk = "Carbs"
  CleanList = junk
End Function

Public Sub Rename()
On Error GoTo errhandl
Dim temp2 As Recordset
Dim junkString As String
Dim IsRecipe As Boolean, AbbrevID As Long
   Set temp2 = DB.OpenRecordset("SELECT * FROM abbrev WHERE (index = " & _
                                  Foodname.SelectedID & ");", dbOpenDynaset)
   If temp2.RecordCount = 0 Then
      MsgBox "Could not find food to rename.  Please enter a valid name", vbCritical, "Database Error"
      Exit Sub
   End If
                                  
   junkString = InputBox("Please enter a new name", "Rename Food", Foodname.Text)
   If temp2("ndb_no") = -100 Then IsRecipe = True Else IsRecipe = False
   AbbrevID = temp2("Index")
   temp2.Edit
   temp2.Fields("Foodname") = junkString
   temp2.Update
   Set temp2 = Nothing
   If IsRecipe Then
      Set temp2 = DB.OpenRecordset("Select * from recipesindex where abbrevid=" & AbbrevID & ";", dbOpenDynaset)
      temp2.Edit
      temp2.Fields("Recipename") = junkString
      temp2.Update
      temp2.Close
      Set temp2 = Nothing
   End If
   Foodname.Text = junkString
   Exit Sub
errhandl:
   MsgBox "Unable to rename food.", vbOKOnly, ""
End Sub

Public Sub DeleteFood()
On Error GoTo errhandl
Dim temp2 As Recordset
Dim temp As Recordset
Dim AbbrevID As Long
   Dim ret As VbMsgBoxResult
   ret = MsgBox("Are you sure that you wish to delete this food?", vbYesNoCancel, "Confirm Delete")
   If ret <> vbYes Then
      Exit Sub
   End If
  
   Set temp = DB.OpenRecordset("SELECT * FROM abbrev WHERE (index = " & _
                                  Foodname.SelectedID & ");", dbOpenDynaset)
   If temp.EOF Then
      MsgBox "Could not find food to delete.  Please enter a valid name", vbCritical, "Database Error"
      Exit Sub
   End If
   AbbrevID = Foodname.SelectedID
   Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & AbbrevID & "));", dbOpenDynaset)
                              
   Do While Not temp2.EOF
     temp2.Delete
     temp2.MoveNext
   Loop
  ' temp2.Update
   Set temp2 = Nothing
   
   temp.MoveFirst
   If temp.Fields("ndb_no") = -100 Then
    Dim RTemp As Recordset, rID As Long
    Set RTemp = DB.OpenRecordset("Select * From recipesindex where (abbrevid = " & AbbrevID & ");", dbOpenDynaset)
    If Not RTemp.EOF Then
      rID = RTemp.Fields("RecipeID")
      RTemp.Delete
      Set RTemp = Nothing
      Set RTemp = DB.OpenRecordset("Select * From Recipes where (recipeid = " & rID & ");", dbOpenDynaset)
      If RTemp.RecordCount <> 0 Then
        While Not RTemp.EOF
          RTemp.Delete
          RTemp.MoveNext
        Wend
      End If
    End If
   End If
   temp.Delete
   temp.Close
   Set temp = Nothing
  
   ClearGrid
   MsgBox "Food has been deleted.", vbOKOnly, ""
   Exit Sub
errhandl:
   MsgBox "Unable to rename food.", vbOKOnly, ""

End Sub
Public Sub ShowFood(DisplayName As String)
  On Error Resume Next
  If DisplayName = "" Then
     Foodname.Text = DisplayName
     Call FoodName_ItemSelected(Foodname.SelectedID)
     Foodname.CloseBox
  Else
     
  End If
End Sub
Private Sub RecipeBound(ByVal New_RecipeBound As Boolean)
  On Error Resume Next
    m_RecipeBound = New_RecipeBound
    
    Call Ingred.OpenDay("1975-01-01") 'this is the set date for recipes
    
End Sub



Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)

End Function

Private Sub VScroll1_Change()
Picture1.Top = (Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = (Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
