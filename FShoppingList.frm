VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FShoppingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shopping List"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   Icon            =   "FShoppingList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtbShopList 
      Height          =   5535
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FShoppingList.frx":57E2
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin CalorieBalance.MonthDayPicker mFrom 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
   End
   Begin CalorieBalance.MonthDayPicker mTo 
      Height          =   2775
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
   End
   Begin VB.Label Label3 
      Caption         =   "List Preview"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FShoppingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Option Compare Text ' compare
Private Type FoodType
   FName As String
   FUnit As String
   FGrams As Single
   FGramsInUnit As Single
End Type


Dim Items As Collection, IDs As Collection, ItemNames As Collection
Dim Units As Collection

Dim Foods() As FoodType
Dim FoodIndex As Collection
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub DoShopping()


    On Error GoTo Err_Proc

Dim junkL As Long
Dim RName As String
Dim temp As Recordset
Dim temp2 As Recordset
Dim SQL As String
Dim TMP

Set FoodIndex = New Collection
ReDim Foods(0)

'put up the waiting page

DoEvents
'change the mousepointer
TMP = Me.MousePointer
Me.MousePointer = 11

'get the records
SQL = "SELECT DaysInfo.*, Abbrev.*, DaysInfo.Date, DaysInfo.Date, DaysInfo.User, DaysInfo.Date, DaysInfo.meal " _
      & "FROM DaysInfo INNER JOIN Abbrev ON DaysInfo.ItemID=Abbrev.Index " _
      & "WHERE (((DaysInfo.Date)>=#" & FixDate(mFrom.GetDate) & "# And " _
      & "(DaysInfo.Date)<=#" & FixDate(mTo.GetDate) & "#) AND " _
      & "((DaysInfo.User)='" & CurrentUser.Username & "' Or (DaysInfo.User)=' temp')) " _
      & " and abbrev.index>-200 " _
      & "ORDER BY abbrev.foodname;"
Set temp = DB.OpenRecordset(SQL, dbOpenDynaset)


'now loop through the records
While Not temp.EOF
    junkL = 0
    If IsNull(temp("ndb_no")) Then
       junkL = 0
    Else
       junkL = temp("ndb_no")
    End If
    'if it is a recipe, get the sub items
    If junkL = -100 Then
'       RName = temp("foodname")
           SQL = " SELECT RecipesIndex.RecipeID, abbrev.*, Abbrev.Foodname, Recipes.Unit, Recipes.Servings " _
           & "FROM (RecipesIndex INNER JOIN Recipes ON RecipesIndex.RecipeID = Recipes.RecipeID) INNER JOIN Abbrev ON Recipes.ItemID = Abbrev.Index " _
           & "WHERE (((RecipesIndex.abbrevid)=" & temp("index") & "));"

           Set temp2 = DB.OpenRecordset(SQL, dbOpenDynaset)
           While Not temp2.EOF
                  Call PutInFood(temp2, "abbrev.foodname")
                  temp2.MoveNext
           Wend
           temp2.Close
           Set temp2 = Nothing
    Else
       Call PutInFood(temp, "foodname")
    End If
    temp.MoveNext
Wend
Dim junk2 As String
'clean up the database
temp.Close
Set temp = Nothing




'*******************Make Html**********************************
Dim junk As String
Dim TName As String
'junk = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f1\fswiss\fcharset0 Arial;}}" & vbCrLf
'junk = junk & "{\*\generator Msftedit 5.41.15.1515;}\viewkind4\uc1\pard\keepn\sb100\sa100\kerning36\b\f0\fs48 Shopping list for " & CurrentUser.Username & "\par" & vbCrLf
'junk = junk & "\trowd\trgaph10\trleft-10\trpaddl10\trpaddr10\trpaddfl3\trpaddfr3" & vbCrLf
'junk = junk & "\clvertalc\cellx1738\clvertalc\cellx2582\clvertalc\cellx8638\clvertalc\cellx9348\pard\intbl\sb100\sa100\kerning0\b0\fs24\cell\b Amount\b0\cell\b Foodname\b0\cell\b Grams\b0\cell\row" & vbCrLf
junk = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f1\fswiss\fcharset0 Arial;}{\f2\fnil\fcharset2 Symbol;}}"
junk = junk & "{\*\generator Msftedit 5.41.15.1515;}\viewkind4\uc1\pard\keepn\sb100\sa100\kerning36\b\f0\fs48 Shopping list for " & CurrentUser.Username & "\par\pard"

   Dim i As Long, j As Long, minI As Long, minString As String
    
For i = 1 To UBound(Foods)
  minString = "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
  For j = 1 To UBound(Foods)
      If Foods(j).FName <> "" Then
         If LCase$(Foods(j).FName) <= minString Then
            minString = LCase$(Foods(j).FName)
            minI = j
         End If
      End If
  Next j
   
   With Foods(minI)
        'junk = junk & "\trowd\trgaph10\trleft-10\trpaddl10\trpaddr10\trpaddfl3\trpaddfr3"
        'junk = junk & "\clvertalc\cellx1738\clvertalc\cellx2582\clvertalc\cellx8638\clvertalc\cellx9348\pard\intbl\sb100\sa100\cell "
        junk = junk & "\pard{\pntext\f2\'B7\tab}{\*\pn\pnlvlblt\pnf2\pnindent0{\pntxtb\'B7}}\fi-720\li720\b0\fs32 "
       
        If .FGramsInUnit <> 0 Then
           junk2 = Module1.ConvertDecimalToFraction(Round(.FGrams / .FGramsInUnit / 100 + 0.5))
        Else
           junk2 = ""
        End If
        junk2 = junk2 & " " & .FUnit
        
        
       'junk = junk & junk2 & "\cell "
       ' junk = junk & .FName & "\cell "
       ' junk = junk & Round(.FGrams, 1) & " \cell\row\ "
      

       junk = junk & junk2 & " "
       junk = junk & .FName & " \tab\tab"
       junk = junk & STR(Round(.FGrams / 100, 1)) & " grams \tab\kerning0\f1\fs20\par\par"
        junk2 = ""
        .FName = ""
   End With
  
Next i

'junk = junk & "\pard\sb100\sa100\par\pard\f1\fs20\par}"
junk = junk & "}"
'***********************Save list**************************
rtbShopList.TextRTF = junk

errhandl:
Set FoodIndex = Nothing

If Err.Number <> 0 Then
   MsgBox Err.Description
 
End If
Me.MousePointer = TMP



Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "FShoppingList", "DoShopping", Err.Description
    Resume Exit_Proc


End Sub


Private Sub PutInFood(temp As Recordset, NameString As String)
Dim TName As String, TID As Long
Dim TUnit As String
Dim junk As String
Dim sum
Dim Parts() As String, parts2() As String
Dim found As Boolean



     Dim GS As Single
     Dim i As Long
     On Error Resume Next
      GS = Module1.TranslateUnitToGrams(temp("index"), temp("unit"))
      i = -1
      i = FoodIndex(temp(NameString))
      'On Error
      If i = -1 Then
         FoodIndex.Add UBound(Foods) + 1, temp(NameString)
         ReDim Preserve Foods(UBound(Foods) + 1)
         i = UBound(Foods)
         With Foods(i)
            .FName = temp(NameString)
            .FUnit = temp("unit")
            .FGrams = GS
            .FGramsInUnit = GS
         End With
      Else
         With Foods(i)
           .FGrams = .FGrams + GS
         End With
      End If
End Sub



Private Sub Command2_Click()
'On Error GoTo errhandl
 PrintRTF rtbShopList, 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch
      Exit Sub
errhandl:
MsgBox "Unable to print." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim mSun As Date
Dim mSat As Date
Dim Today As Date
Today = Date 'DateAdd("w", -3, Date)
mSun = DateAdd("d", -1 * Weekday(Today) + 1, Today)
mSat = DateAdd("d", 6, mSun)
Call mFrom.SetDate(mSun)
Call mTo.SetDate(mSat)
Call DoShopping

End Sub


Private Sub mFrom_DateSelected(NewDate As Date)
On Error Resume Next
Call DoShopping
End Sub

Private Sub mTo_DateSelected(NewDate As Date)
On Error Resume Next
Call DoShopping
End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
