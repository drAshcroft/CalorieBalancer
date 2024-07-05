VERSION 5.00
Begin VB.UserControl RDADisplay 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   ScaleHeight     =   7200
   ScaleWidth      =   3330
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   40
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   39
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   38
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   37
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   36
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   35
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   960
      TabIndex        =   34
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   33
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   32
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   30
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   0
      X2              =   3240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line4 
      Index           =   11
      X1              =   0
      X2              =   3240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   29
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Fat"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Line Line4 
      Index           =   10
      X1              =   0
      X2              =   3240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   10
      Left            =   2280
      TabIndex        =   27
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PolyUnsaturated Fat"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Line Line4 
      Index           =   9
      X1              =   0
      X2              =   3240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   25
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MonoUnsaturated Fat"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   24
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Line Line4 
      Index           =   12
      X1              =   0
      X2              =   3240
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   12
      Left            =   2280
      TabIndex        =   23
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vitamin A"
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   22
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   7
      X1              =   0
      X2              =   3240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   21
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Protein"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   20
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   6
      X1              =   0
      X2              =   3240
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   19
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sugars"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   5
      X1              =   0
      X2              =   3240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   17
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiber"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   4
      X1              =   0
      X2              =   3240
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   15
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Carbohydrates"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   14
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   0
      X2              =   3240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sodium"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   0
      X2              =   3240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cholesterol"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line4 
      Index           =   8
      X1              =   0
      X2              =   3240
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Saturated Fat"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   0
      X2              =   3240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fat"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "% of Daily Value"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   3240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   0
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label LCalories 
      BackStyle       =   0  'Transparent
      Caption         =   "Calories    0    Calories From Fat  0"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Per Serving"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label LFoodname 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Food Name: "
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label LServingSize 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Serving Size: "
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3255
      WordWrap        =   -1  'True
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
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "RDADisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private NutrientMap As Collection

Public Sub DisplayFoodCollection(FoodInfo As exCollection, ideals As Recordset, ServingSize As String)
    LFoodname.Caption = ""
    LServingSize = "Serving Size: " & ServingSize
   
    LCalories = "Calories   " & Round(FoodInfo("calories")) & " Calories From Fat " & Round(FoodInfo("fat") * 9)
    Dim Macro As New Collection
    Macro.Add 0, "Fat"
    Macro.Add 8, "Saturated Fat"
    Macro.Add 9, "Monounsaturated Fat"
    Macro.Add 10, "PolyUnsaturated Fat"
    Macro.Add 11, "trans fat"
    Macro.Add 2, "Cholesterol"
    Macro.Add 3, "Sodium"
    Macro.Add 4, "Carbohydrates"
    Macro.Add 5, "Fiber"
    Macro.Add 6, "Sugars"
    Macro.Add 7, "Protein"
    Dim i As Long
    For i = 1 To Macro.Count
       Label8(Macro(i)).Caption = Round(FoodInfo(Macro(i)))
    
    Next
    For i = 1 To Macro.Count
       Label9(Macro(i)).Caption = Round(FoodInfo(Macro(i)) / ideals(Macro(i)) * 100) & "%"
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
       Label9(NutrientMap(B(i))).Caption = Round(100 * FoodInfo(B(i)) / ideals(B(i))) & "%"
    Next
End Sub
Public Sub DisplayFood(FoodInfo As Recordset, ideals As Recordset, Weight As Recordset, Amount As Double, Unitname As String, Grams As Double)


    LFoodname.Caption = FoodInfo("Foodname")
    LServingSize = "Serving Size: " & Amount & Unitname & " (" & Grams & ")"
    Dim BC As Double
    BC = Weight.Fields("gm_wgt").Value / 100 * Amount
    
    LCalories = "Calories   " & Round(FoodInfo("calories") * BC) & " Calories From Fat " & Round(FoodInfo("fat") * BC * 9)
    Dim Macro As New Collection
    Macro.Add 0, "Fat"
    Macro.Add 1, "Saturated Fat"
    Macro.Add 9, "Monounsaturated Fat"
    Macro.Add 10, "PolyUnsaturated Fat"
    Macro.Add 11, "trans fat"
    Macro.Add 2, "Cholesterol"
    Macro.Add 3, "Sodium"
    Macro.Add 4, "Carbohydrates"
    Macro.Add 5, "Fiber"
    Macro.Add 6, "Sugars"
    Macro.Add 7, "Protein"
    Dim i As Long
    For i = 0 To Macro.Count
       Label8(NutrientMap(Macro(i))).Caption = Round(FoodInfo(Macro(i)) * BC)
    
    Next
    For i = 0 To Macro.Count
       Label9(NutrientMap(Macro(i))).Caption = Round(FoodInfo(Macro(i)) * BC / ideals(Macro(i)))
    Next
    
    Dim B As New Collection
    B.Add 12, "Vitamin A"
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
    
    For i = 0 To B.Count
       Label9(NutrientMap(B(i))).Caption = Round(FoodInfo(B(i)) * BC / ideals(B(i)))
    Next
End Sub

Private Sub UserControl_Initialize()
    Set NutrientMap = New Collection
    
    NutrientMap.Add 0, "Fat"
    NutrientMap.Add 1, "Saturated Fat"
    NutrientMap.Add 9, "Monounsaturated Fat"
    NutrientMap.Add 10, "PolyUnsaturated Fat"
    NutrientMap.Add 11, "trans fat"
    NutrientMap.Add 2, "Cholesterol"
    NutrientMap.Add 3, "Sodium"
    NutrientMap.Add 4, "Carbohydrates"
    NutrientMap.Add 5, "Fiber"
    NutrientMap.Add 6, "Sugars"
    NutrientMap.Add 7, "Protein"
    NutrientMap.Add 12, "Vitamin A"
    
    Dim B As New Collection
    
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
    Dim Index As Integer, DY As Integer, i As Long
    For i = 1 To B.Count
       Index = i + 12
       NutrientMap.Add Index, B(i)
       
       Load Label7(Index)
       Load Label9(Index)
       Load Line4(Index)
       
       Label7(Index).Visible = True
       Label9(Index).Visible = True
       Line4(Index).Visible = True
       
       DY = i * Label7(12).Height * 1.2
       Label7(Index).Top = Label7(12).Top + DY
       Label9(Index).Top = Label9(12).Top + DY
       Line4(Index).Y1 = Line4(12).Y1 + DY
       Line4(Index).Y2 = Line4(12).Y1 + DY
       
       Label7(Index).Caption = B(i)
    Next i
    UserControl.Height = Line4(12).Y1 + DY + 100
End Sub

