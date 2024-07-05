VERSION 5.00
Begin VB.Form FNewFood 
   Caption         =   "New Food"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7710
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   43
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete Selected Food"
      Height          =   495
      Left            =   6480
      TabIndex        =   42
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6480
      TabIndex        =   41
      Top             =   120
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7695
      LargeChange     =   100
      Left            =   5160
      Max             =   1000
      SmallChange     =   50
      TabIndex        =   40
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   13620
      Left            =   120
      ScaleHeight     =   13620
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin CalorieBalance.AutoComplete TFoodname 
         Height          =   375
         Left            =   1320
         TabIndex        =   44
         Top             =   480
         Width           =   3015
         _extentx        =   5318
         _extenty        =   661
         font            =   "FNewFoodN.frx":0000
      End
      Begin VB.TextBox TVitamins 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   17
         Top             =   8760
         Width           =   1335
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   16
         Top             =   7920
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   15
         Top             =   7560
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   14
         Top             =   7200
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   4
         Left            =   2040
         TabIndex        =   13
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   6120
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   11
         Left            =   2040
         TabIndex        =   10
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   10
         Left            =   2040
         TabIndex        =   9
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   9
         Left            =   2040
         TabIndex        =   8
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   7
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox TMacros 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox TCalories 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   3840
         Width           =   2295
      End
      Begin VB.ListBox FoodGroup 
         Height          =   1035
         ItemData        =   "FNewFoodN.frx":002C
         Left            =   0
         List            =   "FNewFoodN.frx":0045
         TabIndex        =   4
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox TGrams 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TUnit 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TAmount 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "% Daily Value"
         Height          =   255
         Left            =   2160
         TabIndex        =   39
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Line Line5 
         BorderWidth     =   5
         Index           =   0
         X1              =   0
         X2              =   4440
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fat"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Grams in Serving"
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nutrition Facts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label LServingSize 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Serving Size: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   960
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   0
         X1              =   0
         X2              =   4440
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label LFoodname 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Food Name: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Per Serving"
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label LCalories 
         BackStyle       =   0  'Transparent
         Caption         =   "Calories  "
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   3840
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   3240
         X2              =   0
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   1
         X1              =   0
         X2              =   4440
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   0
         X2              =   3240
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Saturated Fat"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Line Line4 
         Index           =   8
         X1              =   0
         X2              =   3240
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cholesterol"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   27
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   0
         X2              =   3240
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sodium"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   26
         Top             =   6600
         Width           =   855
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   0
         X2              =   3240
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Carbohydrates"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Top             =   6960
         Width           =   1095
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   0
         X2              =   3240
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiber"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   7320
         Width           =   615
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   0
         X2              =   3240
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sugars"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   7680
         Width           =   855
      End
      Begin VB.Line Line4 
         Index           =   6
         X1              =   0
         X2              =   3240
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Protein"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   22
         Top             =   8040
         Width           =   1575
      End
      Begin VB.Label LabelV 
         BackStyle       =   0  'Transparent
         Caption         =   "Vitamin A"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   8760
         Width           =   1575
      End
      Begin VB.Line LineV 
         Index           =   0
         X1              =   0
         X2              =   3240
         Y1              =   9000
         Y2              =   9000
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "MonoUnsaturated Fat"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   20
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Line Line4 
         Index           =   9
         X1              =   0
         X2              =   3240
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PolyUnsaturated Fat"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   19
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Line Line4 
         Index           =   10
         X1              =   0
         X2              =   3240
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Fat"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   18
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Line Line4 
         Index           =   11
         X1              =   0
         X2              =   3240
         Y1              =   6120
         Y2              =   6120
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   5535
      Left            =   5760
      TabIndex        =   45
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "FNewFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private MacroNutrientMap As Collection
Private VitaminMap As Collection



Private Sub DisplayFood(FoodInfo As Recordset, ideals As Recordset, Weight As Recordset)

On Error Resume Next
    TFoodname.Text = FoodInfo("Foodname")
    TAmount.Text = Weight.Fields("amount").Value
    TUnit.Text = Weight.Fields("msre_desc").Value
    TGrams.Text = Weight.Fields("gm_wgt").Value
    
    Dim FoodGroupI As String
    FoodGroupI = Left$(FoodInfo("foodgroup") & "0000", 4)
    Dim BC As Double
    Dim Amount As Double
    Amount = Weight.Fields("amount").Value
    BC = Weight.Fields("gm_wgt").Value / 100 * Amount
   
    TCalories = Round(FoodInfo("calories") * BC)
    Dim Macro As New Collection
    Macro.Add "Fat"
    Macro.Add "Saturated Fat"
    Macro.Add "Monounsaturated Fat"
    Macro.Add "PolyUnsaturated Fat"
    Macro.Add "trans fat"
    Macro.Add "Cholesterol"
    Macro.Add "Sodium"
    Macro.Add "Carbs"
    Macro.Add "Fiber"
    Macro.Add "Sugar"
    Macro.Add "Protein"
    Dim i As Long
    For i = 1 To Macro.Count
     If Not IsNull(FoodInfo(Macro(i))) Then
       TMacros(MacroNutrientMap(Macro(i))).Text = Round(FoodInfo(Macro(i)) * BC, 1)
     End If
    
    Next
    
    
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
    B.Add "Vitamin D"
    B.Add "Vitamin E"
    B.Add "Vitamin K"
    B.Add "Alpha-carotene"
    
    For i = 1 To B.Count
       If Not IsNull(FoodInfo(B(i))) Then
         TVitamins(VitaminMap(B(i))).Text = Round(FoodInfo(B(i)) * BC / ideals(B(i)) * 100, 1)
       End If
    Next
    
    For i = 0 To FoodGroup.ListCount - 1
       If Val(FoodGroup.ItemData(i)) = Val(FoodGroupI) Then
          FoodGroup.Selected(i) = True
       End If
    Next i
End Sub
Private Sub ClearGrid()
   TFoodname.Text = ""
   
   
    TAmount.Text = ""
    TUnit.Text = ""
    TGrams.Text = ""
    
   
    On Error Resume Next
    Dim i As Long
    For i = TMacros.LBound To TMacros.UBound
       TMacros(i).Text = ""
    
    Next
    
      
    For i = TVitamins.LBound To TVitamins.UBound
       TVitamins(i).Text = ""
    Next
End Sub
Public Sub ShowFood(DisplayName As String)
  On Error Resume Next
  If DisplayName = "" Then
     ClearGrid
  Else
     Dim RS As Recordset, Weight As Recordset, ideals As Recordset
     Set RS = DB.OpenRecordset("select * from abbrev where foodname='" & DisplayName & "';", dbOpenDynaset)
     Set Weight = DB.OpenRecordset("select * from weight where index=" & RS("index") & ";", dbOpenDynaset)
     Set ideals = DB.OpenRecordset("select * from ideals where user='AnyUser';", dbOpenDynaset)
     Call DisplayFood(RS, ideals, Weight)
     RS.Close
     Weight.Close
     ideals.Close
     
  End If
End Sub
Private Function ValidateFood() As Boolean
    ValidateFood = True
    If TFoodname.Text = "" Then
       Call MsgBox("Please give the food a new name.", vbOKOnly, "")
       ValidateFood = False
       Exit Function
    End If
    Dim Grams As Single
    Grams = Val(TGrams.Text)
    If Grams = 0 Then
        MsgBox "Please enter number of grams in serving" & vbCrLf & "(estimate if needed, 100 grams works if you just do not know)", vbOKOnly, ""
        ValidateFood = False
        Exit Function
    End If
    
End Function
Public Function SaveFood() As Boolean
'On Error Resume Next
Dim cc As Single
Dim Grams As Single
Dim AbbrevID As Long
Dim Amount As Single
    SaveFood = True
       
    If Not ValidateFood Then
      SaveFood = False
      Exit Function
    End If
    
   ' AbbrevID = TFoodname.SelectedID
    Grams = Val(TGrams.Text)
    Amount = Val(ConvertFractionToDecimal(TAmount.Text))
    If Grams = 0 Then Grams = 100
    cc = 100 / Grams

    Dim temp2 As Recordset
    Dim sTemp As Recordset
    
    Dim ideals As Recordset
    'since they are entering from a label, need to use normal persons information
    Set ideals = DB.OpenRecordset("Select * from ideals where user='AnyUser';", dbOpenDynaset)

    Set sTemp = DB.OpenRecordset("SELECT * FROM ABBREV where foodname ='" _
                      & Replace(TFoodname.Text, "'", "''") & "';", dbOpenDynaset)
                      
    If sTemp.EOF Then
        sTemp.AddNew
        Set temp2 = DB.OpenRecordset("Select max(index) as MaxIt from abbrev", dbOpenDynaset)
        AbbrevID = temp2("maxit") + 1
        temp2.Close
        Set temp2 = Nothing
        sTemp("Index") = AbbrevID
    Else
        AbbrevID = sTemp("index")
        sTemp.Edit
    End If
  
  
    Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & AbbrevID & "));", dbOpenDynaset)
'todo: this is rather scorched earth.  need to just delete the overlapping weight inforation
    Do While Not temp2.EOF
       temp2.Delete
       temp2.MoveNext
    Loop
    temp2.AddNew

    sTemp.Fields("Usage") = 10
    sTemp.Fields("Foodname") = TFoodname.Text
  
    Dim junk As String, junk2
    
  
    Dim Macro As New Collection
    Macro.Add "Fat"
    Macro.Add "Saturated Fat"
    Macro.Add "Monounsaturated Fat"
    Macro.Add "PolyUnsaturated Fat"
    Macro.Add "trans fat"
    Macro.Add "Cholesterol"
    Macro.Add "Sodium"
    Macro.Add "Carbs"
    Macro.Add "Fiber"
    Macro.Add "Sugar"
    Macro.Add "Protein"
    
    sTemp("calories") = Val(TCalories.Text) * cc
    
    Dim i As Long
    For i = 1 To Macro.Count
        If Not TMacros(MacroNutrientMap(Macro(i))).Text = "" Then
          junk2 = Val(TMacros(MacroNutrientMap(Macro(i))).Text)
          sTemp(Macro(i)) = junk2 * cc
        End If
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
    B.Add "Vitamin D"
    B.Add "Vitamin E"
    B.Add "Vitamin K"
    B.Add "Alpha-carotene"
    
    For i = 1 To B.Count
        If Not TVitamins(VitaminMap(B(i))).Text = "" Then
          junk2 = Val(TVitamins(VitaminMap(B(i))).Text)
          sTemp(B(i)) = junk2 / 100 * ideals(B(i)) * cc
        End If
    Next i
    
    For i = 0 To FoodGroup.ListCount - 1
       If FoodGroup.Selected(i) = True Then
          sTemp("Foodgroup") = Right$("0000" & FoodGroup.ItemData(i), 4)
       
       End If
       
    Next i
    
  

  sTemp.Update

  temp2.Fields("Gm_Wgt") = Grams
  temp2.Fields("Amount") = Amount
  temp2.Fields("Msre_desc") = TUnit.Text
  temp2.Fields("Index") = AbbrevID
  temp2.Update
  Set temp2 = Nothing
  Set sTemp = Nothing
  MsgBox "Food was saved successfully", vbOKOnly, ""
  
   
End Function
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
                                  TFoodname.SelectedID & ");", dbOpenDynaset)
   If temp.EOF Then
      MsgBox "Could not find food to delete.  Please enter a valid name", vbCritical, "Database Error"
      Exit Sub
   End If
   AbbrevID = TFoodname.SelectedID
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
  If DoDebug Then
     Stop
     Resume
  Else
   MsgBox "Unable to rename food.", vbOKOnly, ""
  End If
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

Private Sub CCancel_Click()
   Unload Me
End Sub

Private Sub CDelete_Click()
   DeleteFood
   Unload Me
End Sub

Private Sub CSave_Click()
On Error GoTo errhandl
   SaveFood
   Unload Me
   Exit Sub
errhandl:
   Call MsgBox("Something has gone wrong with save.  Please check everything over and try again.", vbOKOnly, "")
End Sub

Private Sub Form_Load()
TFoodname.AddDataBase Module1.DB
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

   ' TFoodname.AddDataBase Module1.DB
    Set MacroNutrientMap = New Collection
    
    MacroNutrientMap.Add 0, "Fat"
    MacroNutrientMap.Add 8, "Saturated Fat"
    MacroNutrientMap.Add 9, "Monounsaturated Fat"
    MacroNutrientMap.Add 10, "PolyUnsaturated Fat"
    MacroNutrientMap.Add 11, "trans fat"
    MacroNutrientMap.Add 2, "Cholesterol"
    MacroNutrientMap.Add 3, "Sodium"
    MacroNutrientMap.Add 4, "Carbs"
    MacroNutrientMap.Add 5, "Fiber"
    MacroNutrientMap.Add 6, "Sugar"
    MacroNutrientMap.Add 7, "Protein"
    
    
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
    B.Add "Vitamin D"
    B.Add "Vitamin E"
    B.Add "Vitamin K"
    B.Add "Alpha-carotene"
    Set VitaminMap = New Collection
    VitaminMap.Add 0, B(1)
    Dim Index As Integer, DY As Integer, i As Long
    For i = 2 To B.Count
       Index = i - 1
       VitaminMap.Add Index, B(i)
       
       Load LabelV(Index)
       Load TVitamins(Index)
       Load LineV(Index)
       
       LabelV(Index).Visible = True
       TVitamins(Index).Visible = True
       LineV(Index).Visible = True
       
       DY = (i - 1) * TVitamins(0).Height * 1.2
       LabelV(Index).Top = LabelV(0).Top + DY
       TVitamins(Index).Top = TVitamins(0).Top + DY
       LineV(Index).Y1 = LineV(0).Y1 + DY
       LineV(Index).Y2 = LineV(0).Y1 + DY
       
       LabelV(Index).Caption = B(i)
    Next i
    Picture1.Height = LineV(0).Y1 + DY + 100
    

    Label8.Caption = "To enter a new food:" & vbCrLf & _
       "Enter a name.  The drop down will show all the existing foods.  Choose a name that is original and then press enter to move to the next step. " & vbCrLf & _
       "Enter a useful serving amount.  Examples can be 10 grapes, 1 apple, 1 cheese burger. " & vbCrLf & _
       "You must enter a grams per serving amount.  This is on most nutrition labels.  If you do not know this number then enter 100" & vbCrLf & _
       "You must enter the number of calories per serving." & vbCrLf & vbCrLf & _
       "You can now enter as much nutrition information as possible.  The information in the top part is all entered in absolute amounts like shown in a nutrition label.  The vitamin information should be entered in percents as is shown on a standard nutrition label.  If you need help just turn over any food package and the information should be readily available."


End Sub



Private Sub Form_Resize()
VScroll1.Height = ScaleHeight

End Sub

Private Sub TAmount_Change()
   TFoodname.CloseBox
End Sub

Private Sub TAmount_GotFocus()
   TFoodname.CloseBox
End Sub

Private Sub TFoodname_ExitFocus()
    TFoodname.CloseBox
End Sub

Private Sub TFoodname_ItemSelected(SelectedID As Long)
 
     Dim RS As Recordset, Weight As Recordset, ideals As Recordset
     Set RS = DB.OpenRecordset("select * from abbrev where index=" & SelectedID & ";", dbOpenDynaset)
     Set Weight = DB.OpenRecordset("select * from weight where index=" & SelectedID & ";", dbOpenDynaset)
     Set ideals = DB.OpenRecordset("select * from ideals where user='AnyUser';", dbOpenDynaset)
     Call DisplayFood(RS, ideals, Weight)
     RS.Close
     Weight.Close
     ideals.Close
     TFoodname.CloseBox
     
End Sub

Private Sub TFoodname_LostFocus()
   TFoodname.CloseBox
End Sub

Private Sub TFoodname_NoneSelected()
  TFoodname.CloseBox
End Sub

Private Sub TFoodname_TabPress(Shift As Boolean)
   TFoodname.CloseBox
End Sub

Private Sub TGrams_GotFocus()
   TFoodname.CloseBox
End Sub

Private Sub TUnit_Change()
   TFoodname.CloseBox
End Sub

Private Sub TUnit_GotFocus()
   TFoodname.CloseBox
End Sub

Private Sub VScroll1_Change()
Picture1.Top = (Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = (Height - Picture1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
