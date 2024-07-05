VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMainO 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Calorie Balance 2.8"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -2685
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox DragTab 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   9840
      MousePointer    =   9  'Size W E
      Picture         =   "frmMain.frx":57E2
      ScaleHeight     =   2445
      ScaleWidth      =   210
      TabIndex        =   25
      Top             =   2760
      Width           =   210
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Website"
      Height          =   1215
      Left            =   1440
      Picture         =   "frmMain.frx":64E0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6480
      Width           =   1095
   End
   Begin CalorieBalance.AdvancedFlex FlexDiet 
      Height          =   6195
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   8655
      _ExtentX        =   5741
      _ExtentY        =   6271
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
      ShowAsPercent   =   0   'False
   End
   Begin CalorieBalance.cMealPlanner MP 
      Height          =   5295
      Left            =   2040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   15055
      _ExtentY        =   11668
   End
   Begin VB.Frame FMeals 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2010
      Left            =   8040
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   3255
      Begin MSComctlLib.TreeView CatSearch 
         Height          =   1455
         Left            =   1080
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2566
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "Folderlist1"
         Appearance      =   1
      End
      Begin MSComctlLib.TabStrip TSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Foods"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Favorite Foods"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Meals"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CEditMeal 
         Caption         =   "Edit Meal"
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CNewMeal 
         Caption         =   "New Meal"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin MSComctlLib.TreeView Meals 
         Height          =   6285
         Left            =   360
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   11086
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "Folderlist1"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView Favorites 
         Height          =   1455
         Left            =   0
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2566
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "Folderlist1"
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag and Drop items to desired location"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   2640
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label Label1 
         Caption         =   "Downloading Plan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1440
         TabIndex        =   16
         Top             =   1200
         Width           =   3240
      End
   End
   Begin MSComctlLib.ImageList Folderlist1 
      Left            =   3000
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":99A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   105
      Left            =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6690
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   185
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":9E88
   End
   Begin CalorieBalance.MonthDayPicker Calendar 
      Height          =   2535
      Left            =   3960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4471
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   1920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2040
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CalorieBalance.Exercise Exercise 
      Height          =   3135
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      BackColor       =   8388608
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   900
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Skin_Up         =   "frmMain.frx":9F13
      Skin_Hover      =   "frmMain.frx":AB3F
      Skin_Down       =   "frmMain.frx":D50D
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   900
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Skin_Up         =   "frmMain.frx":FF8F
      Skin_Hover      =   "frmMain.frx":1085B
      Skin_Down       =   "frmMain.frx":13229
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   900
      Index           =   2
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Skin_Up         =   "frmMain.frx":15CAB
      Skin_Hover      =   "frmMain.frx":162BF
      Skin_Down       =   "frmMain.frx":18C8D
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   900
      Index           =   3
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Skin_Up         =   "frmMain.frx":1B70F
      Skin_Hover      =   "frmMain.frx":1BD11
      Skin_Down       =   "frmMain.frx":1E6DF
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   900
      Index           =   4
      Left            =   2880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Skin_Up         =   "frmMain.frx":21161
      Skin_Hover      =   "frmMain.frx":2194F
      Skin_Down       =   "frmMain.frx":2431D
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   885
      Index           =   5
      Left            =   5760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1561
      Skin_Up         =   "frmMain.frx":26D9F
      Skin_Hover      =   "frmMain.frx":2976D
      Skin_Down       =   "frmMain.frx":2C13B
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   885
      Index           =   7
      Left            =   4800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1561
      Skin_Up         =   "frmMain.frx":2EB09
      Skin_Hover      =   "frmMain.frx":314D7
      Skin_Down       =   "frmMain.frx":33EA5
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   990
      Index           =   8
      Left            =   7680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1746
      Skin_Up         =   "frmMain.frx":36873
      Skin_Hover      =   "frmMain.frx":39A45
      Skin_Down       =   "frmMain.frx":3C413
   End
   Begin CalorieBalance.EasyHover EasyHover1 
      Height          =   915
      Index           =   9
      Left            =   6720
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1614
      Skin_Up         =   "frmMain.frx":3F5E5
      Skin_Hover      =   "frmMain.frx":4211B
      Skin_Down       =   "frmMain.frx":44AE9
   End
   Begin VB.PictureBox PMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   2280
      ScaleHeight     =   5985
      ScaleWidth      =   6825
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton COKMessage 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Label DropLabel 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3420
      TabIndex        =   14
      Top             =   1755
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Make New User"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditUser 
         Caption         =   "&Edit Current User"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "&Delete Current User"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Change to Different User"
      End
      Begin VB.Menu mnusep45678 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLastEntry 
         Caption         =   "&Review Last Entry"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Personal File"
      End
      Begin VB.Menu mnuOpenPersonal 
         Caption         =   "&Open Personal File"
      End
      Begin VB.Menu mnusep091934 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Make Database &Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Backup"
      End
      Begin VB.Menu DownloadFoods 
         Caption         =   "&Download Roaming"
         Visible         =   0   'False
      End
      Begin VB.Menu Uploaddaysfood 
         Caption         =   "&Upload Roaming"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep423 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveMealplan 
         Caption         =   "Save Meal Plan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveWeek 
         Caption         =   "Save Week as Meal Plan"
      End
      Begin VB.Menu MnuSaveExercisePlan 
         Caption         =   "Save Exercise Plan"
      End
      Begin VB.Menu mnuOpenExercisePlan 
         Caption         =   "Open Plan"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuPrintRecipe 
         Caption         =   "Print Today's Recipes"
      End
      Begin VB.Menu mnuPrintInstructions 
         Caption         =   "Print Meal Instructions"
      End
      Begin VB.Menu mnuWeeksInstructions 
         Caption         =   "Print Week's Meal Instructions"
      End
      Begin VB.Menu mnuShoppingList 
         Caption         =   "Print Shopping &List"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuToolbars 
      Caption         =   "&Toolbars"
      Visible         =   0   'False
      Begin VB.Menu mnuMeals 
         Caption         =   "Meals"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Options"
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change &Font"
      End
      Begin VB.Menu mnuPercents 
         Caption         =   "Show Nutrients as Percent"
      End
      Begin VB.Menu mnuRemind 
         Caption         =   "Remind me to enter weight"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOrder 
         Caption         =   "Order Meals "
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "      Replace Meals when Dropped In Planner"
      End
   End
   Begin VB.Menu mnuGetPlans 
      Caption         =   "Get Meal Plans"
   End
   Begin VB.Menu mnuGetRecipes 
      Caption         =   "Get Recipes"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuBugReport 
      Caption         =   "Bug Report"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContactFirm 
         Caption         =   "Contact Us"
      End
      Begin VB.Menu mnuHelpMovies 
         Caption         =   "Help Movies"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "Get Support"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteAllRowsInMeal 
         Caption         =   "Remove Meal"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuSep1243 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutRow 
         Caption         =   "Cut Row"
      End
      Begin VB.Menu mnuCopyRow 
         Caption         =   "Copy Row"
      End
      Begin VB.Menu mnuPasteRow 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnusep412 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGetNuts 
         Caption         =   "Get Nutrient Information"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMakeMeal 
         Caption         =   "Make Into Meal"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "PopUpMenu2"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteRow2 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuInsertRow2 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnusep432 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutRow2 
         Caption         =   "Cut Row"
      End
      Begin VB.Menu mnuCopyRow2 
         Caption         =   "Copy Row"
      End
      Begin VB.Menu mnuPasteRow2 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSep0987 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowCalories 
         Caption         =   "Show Calories"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnupop3 
      Caption         =   "Popupmenu3"
      Visible         =   0   'False
      Begin VB.Menu mnu_insert_meal 
         Caption         =   "Insert Meal In Planner"
      End
      Begin VB.Menu mnuNewMeal 
         Caption         =   "New Meal"
      End
      Begin VB.Menu mnuEditMeal 
         Caption         =   "Edit Meal"
      End
      Begin VB.Menu mnuDeleteMeal 
         Caption         =   "Delete Meal"
      End
      Begin VB.Menu mnuViewMeal 
         Caption         =   "View Meal"
      End
      Begin VB.Menu mnuasdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewPlan 
         Caption         =   "New Plan"
      End
      Begin VB.Menu mnuDeletePlan 
         Caption         =   "Delete Plan"
      End
   End
   Begin VB.Menu mnuPopMealPlanner 
      Caption         =   "popupMealPlanner"
      Visible         =   0   'False
      Begin VB.Menu mnuMealCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuMealCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuMealPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnusep09835 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMealPlanner 
         Caption         =   "View Meal"
      End
      Begin VB.Menu mnuEditMealPlanner 
         Caption         =   "Edit Meal"
      End
   End
End
Attribute VB_Name = "frmMainO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
'Private Declare Function OSWinHelp% Lib "USER32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public LastUser As String
Dim UpdateS As String, UUrl As String, HUrl As String
Dim PrintMode As String
Dim PrintURL As String
Dim NavigateError As Boolean
Dim NewName As String
Dim ExerciseOnTop As Boolean
Dim MealTime As String, MealRowNumber As Long, MealReplace As Boolean

Dim LastMessage As Long

Dim CatSearchN As String, CSM As Boolean, csX As Single, csY As Single
Dim fSM As Boolean, fsX As Single, fsY As Single


Public MealSaved As Boolean
Private Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer

Dim MessageFoods As Boolean
Dim MessageMeals As Boolean
Dim MessageWeb As Boolean
Dim MessageExercise As Boolean

Dim DSM As Boolean  ' = True
Dim dsX As Single  ' = x
Dim dsY As Single   ' = y
Dim RFM As Boolean ' for resizing the meals control
Dim CPlanID As Long
Dim MD As Boolean
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub LoadDeepSearch()
On Error GoTo errhandl
    Dim i As Long, j As String, J2 As String
    
    Dim node1 As Node
    Dim expands As New Collection
    Dim ENames As New Collection
    For i = 1 To CatSearch.Nodes.Count
      Set node1 = CatSearch.Nodes(i)
      expands.Add node1.Expanded, node1.Key
      ENames.Add node1.Key
    Next i
    
    Dim FoodgroupS  As Recordset
    Set FoodgroupS = DB.OpenRecordset("select * from foodgroups where parentnumber=-10;", dbOpenDynaset)
    
    CatSearch.Nodes.Clear
    Favorites.Nodes.Clear
    Dim junk As String
    While Not FoodgroupS.EOF
       i = FoodgroupS("catnumber")
       If i < 0 Then
         j = "M" & i
         Call CatSearch.Nodes.Add(, , j, FoodgroupS("category"), 1)
         Call Favorites.Nodes.Add(, , j, FoodgroupS("category"), 1)
         
       Else
         j = "M" & Format(i, "0000")
         Call CatSearch.Nodes.Add(, , j, FoodgroupS("category"), 1)
         Call Favorites.Nodes.Add(, , j, FoodgroupS("category"), 1)
         J2 = "I" & Format(i, "0000")
         Call CatSearch.Nodes.Add(j, 4, J2, "General", 2)
         Call Favorites.Nodes.Add(j, 4, J2, "General", 2)
       End If
       FoodgroupS.MoveNext
    Wend
    FoodgroupS.Close
    Set FoodgroupS = DB.OpenRecordset("select * from foodgroups where parentnumber<>-10 order by category;", dbOpenDynaset)
    While Not FoodgroupS.EOF
       i = FoodgroupS("parentnumber")
       If i < 0 Then
          j = "M" & i
       Else
          j = "M" & i
       End If
       J2 = Format(FoodgroupS("catnumber"), "0000")
       Call CatSearch.Nodes.Add(j, 4, "I" & J2, FoodgroupS("category"), 2)
       Call Favorites.Nodes.Add(j, 4, "I" & J2, FoodgroupS("category"), 2)
       FoodgroupS.MoveNext
    Wend

    Dim RS As Recordset
    Set RS = DB.OpenRecordset("Select * from abbrev order by foodgroup,foodname;", dbOpenDynaset)
    On Error Resume Next
    While Not RS.EOF
      j = RS("foodgroup") & ""
      j = Format(j, "0000")
      If j = "" Then j = "0000"
      CatSearch.Nodes.Add "I" & j, 4, "a" & RS("index"), RS("foodname"), 3
      j = ""
      RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    i = 0
    Set RS = DB.OpenRecordset("select foodname,index,foodgroup from abbrev where index>0 order by usage desc;", dbOpenDynaset)
    
    While (Not RS.EOF) And (i < 100)
      j = RS("foodgroup") & ""
      j = Format(j, "0000")
      Favorites.Nodes.Add "I" & j, 4, "a" & RS("index"), RS("foodname"), 3
      j = ""
      RS.MoveNext
      i = i '+ 1
    Wend
    RS.Close
    Set RS = Nothing
    
    
    On Error Resume Next
    For i = 1 To ENames.Count
       CatSearch.Nodes(ENames(i)).Expanded = expands(ENames(i))
    Next i

    
    Exit Sub
errhandl:
  If DoDebug Then
    Debug.Print Err.Description
    Stop
    Resume
  Else
    Resume Next
  End If
End Sub

Private Function MakeUserHash(nName As String) As String


    On Error GoTo Err_Proc
 MakeUserHash = CurrentUser.HashNumber
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "frmMain", "MakeUserHash", Err.Description
    Resume Exit_Proc


End Function
Public Sub UpdateUserNuts()
On Error GoTo errhandl
   Dim ff As Long
   Dim RS As Recordset, ps As Recordset, Abbrev As Recordset, lastdate As Date
   Dim Foods As String, Exer As String, Journ As String, i As Long
   Dim Sep As String, Username As String
   Username = MakeUserHash(CurrentUser.Username)
   Sep = "^^"
   ff = FreeFile
   Open App.path & "\resources\temp\temp_updateuser.htm" For Output As ff
   Set ps = DB.OpenRecordset("select * from profiles where user='" & CurrentUser.Username & "';", dbOpenDynaset)
   If IsNull(ps("journalupdate")) Then
     lastdate = "2000-01-01"
   Else
     lastdate = ps("journalupdate")
   End If
   Set RS = DB.OpenRecordset("SELECT * FROM DaysInfo " _
         & " WHERE ( ((DaysInfo.date)>=#" & FixDate(lastdate) _
         & "# and (DaysInfo.date)<#" & FixDate(Today) _
         & "#) and user='" & CurrentUser.Username & "') order by daysinfo.date,Daysinfo.order;", dbOpenDynaset)
   ps.Edit
   ps("JournalUpdate") = FixDate(DateAdd("d", -1, Today))
   ps.Update
   ps.Close
   Set ps = Nothing
   Dim gm As Single
   While Not RS.EOF
     Set Abbrev = DB.OpenRecordset("select * from abbrev where index=" & RS("itemid") & ";", dbOpenDynaset)
     If RS("itemid") > 0 Then
        gm = Module1.TranslateUnitToGrams(RS("itemid"), RS("unit"))
     Else
        gm = 0
     End If
     Foods = Foods & RS("date") & Sep
     Foods = Foods & Username & Sep
     Foods = Foods & RS("itemid") & Sep
     Foods = Foods & Abbrev("foodname") & Sep
     Foods = Foods & RS("unit") & Sep
     Foods = Foods & RS("servings") & Sep
     Foods = Foods & RS("meal") & Sep
     Foods = Foods & gm & Sep
     Foods = Foods & Abbrev("calories") & Sep
     Foods = Foods & Abbrev("protein") & Sep
     Foods = Foods & Abbrev("fat") & Sep
     Foods = Foods & Abbrev("carbs") & Sep
     Foods = Foods & Abbrev("sugar") & Sep
     Foods = Foods & Abbrev("fiber") & Sep
     Foods = Foods & Abbrev("calcium") & Sep
     Foods = Foods & Abbrev("sodium") & vbCrLf
     If Not ps Is Nothing Then ps.Close
     Set ps = Nothing
     Abbrev.Close
     Set Abbrev = Nothing
     RS.MoveNext
   Wend
   RS.Close
   Set RS = Nothing
   Set RS = DB.OpenRecordset("select * from dailylog where (user='" & CurrentUser.Username & "' and " _
            & " dailylog.Date>=#" & FixDate(lastdate) _
            & "# and dailylog.Date<#" & FixDate(Today) _
            & "#)  order by dailylog.date;", dbOpenDynaset)
   While Not RS.EOF
     For i = 0 To RS.Fields.Count - 2
       If LCase$(RS.Fields(i).Name) <> "user" Then
         Journ = Journ & RS(i) & Sep
       Else
         Journ = Journ & Username & Sep
       End If
     Next i
     Journ = Journ & RS(i) & vbCrLf
     RS.MoveNext
   Wend
   RS.Close
   Set RS = Nothing
   Dim Fdate As Date
   Fdate = Module1.FindFirstDay(Today)
   Set RS = DB.OpenRecordset("select * from exerciselog where user='" & CurrentUser.Username & "' and week=#" & FixDate(Fdate) & "#;", dbOpenDynaset)
   While Not RS.EOF
     Set ps = DB.OpenRecordset("select * from AbbrevExercise where index=" & RS("exerciseid") & ";", dbOpenDynaset)
     Exer = Exer & FixDate(RS("week")) & Sep
     Exer = Exer & Username & Sep
     Exer = Exer & RS("weekinfo") & Sep
     Exer = Exer & ps("exercisename") & Sep & vbCrLf
     ps.Close
     Set ps = Nothing
     RS.MoveNext
   Wend
   RS.Close
   Set RS = Nothing
   Print #ff, "<html><body onload=""document.NutUpdate.submit()"">"
   Print #ff, "<form name=""NutUpdate"" method=""post"" action=""http://www.caloriebalancediet.com/dietBattle_update.asp"">"
   Print #ff, "<textarea name=foods>" & Foods & "</textarea>"
   Print #ff, "<textarea name=journ>" & Journ & "</textarea>"
   Print #ff, "<textarea name=exer>" & Exer & "</textarea>"
   'Print #ff, "<input type=submit >"
   Print #ff, "</form></body></html>"
   Close #ff
  
   OpenURL App.path & "\resources\temp\temp_updateuser.htm", vbMaximizedFocus
errhandl:
End Sub

Public Sub UploadInfo(Today As Date)
 On Error GoTo errhandl
  Dim junk As String
  Dim temp As Recordset, i As Long
  Dim ID As Long, Unit As String, Serving As Single, Food As String
  Dim Coled As String
  Dim ff As Long
  ff = FreeFile
  Open App.path & "\resources\daily\Roaming.htm" For Output As #ff
  
  Print #ff, "<Html><body>"
  Print #ff, "<form method=""POST"" action=""http://www.caloriebalancediet.com/submittoday.asp"">"
  Print #ff, "<input type=""text"" name=""User"" value =""" & CurrentUser.Username & """><br>"
  Print #ff, "<textarea  name=""Info"" >"
  Coled = "<td style=""border-left-style: solid; border-left-width: 6; border-top-width: 1; border-bottom-width: 1"" align=""right"">"
  Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.date)=#" & FixDate(Today) & "#) AND (DaysInfo.user='" & CurrentUser.Username & "')) ORDER BY daysinfo.order;", dbOpenDynaset)
  'get the date
   i = 0
  'if there are records, then put them into the data record
  If Not temp.EOF Then
     If Not (temp.EOF = True And temp.BOF = True) Then
        junk = "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
        junk = junk & "<tr><td><b>Servings </b></td><td><b>Units</b></td><td align=""left""><b>Foodname</b></td>"
        junk = junk & Coled & "<b>Calories</b></td>"
        junk = junk & Coled & "<b>Sugar</b>"
        junk = junk & Coled & "<b>Fiber</b>"
        junk = junk & Coled & "<b>Carbs</b>"
        junk = junk & Coled & "<b>Fat</b>"
        junk = junk & Coled & "<b>Protein</b>"
        junk = junk & Coled & "<b>Grams</b>"
        junk = junk & "</tr>"

        temp.MoveFirst
        i = 1
        
        Do While Not temp.EOF
          ID = temp.Fields("itemID")

          Unit = ""
          Serving = 0
          If ID <> 0 And ID <> -1111 Then
             Unit = temp.Fields("unit")
             Serving = temp.Fields("Servings")
          End If
          If (Unit <> "" And Serving <> 0) Or ID <= -200 Then
              Dim temp2 As Recordset
              Set temp2 = DB.OpenRecordset("Select * from abbrev " & _
              " where index = " & ID & ";", dbOpenDynaset)
              Food = temp2.Fields("foodname")
             If ID >= 0 Then
                   Dim GS As Single
                   
 
                  GS = Serving * Module1.TranslateUnitToGrams(ID, Unit) / 100 'temp3.Fields("gm_wgt").Value / temp3.Fields("amount").Value / 100 * serving
             
                  junk = junk & "  <tr><td>" & Module1.ConvertDecimalToFraction(Serving) & "</td>" & vbCrLf
                  junk = junk & "<td>" & Unit & "</td>" & vbCrLf
                  junk = junk & "<td align=""left"">" & Food & "</td>" & vbCrLf
              
                  junk = junk & Coled & Round(temp2("Calories") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(temp2("Sugar") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(temp2("Fiber") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(temp2("Carbs") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(temp2("Fat") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(temp2("Protein") * GS, 1) & "</td>" & vbCrLf
                  junk = junk & Coled & Round(GS * 100, 2) & "</td>" & vbCrLf
                  junk = junk & "</tr>" & vbCrLf
               Else
                  junk = junk & "<tr><td colspan=10 bgcolor=""#99CCFF""><h3><b>" & Food & "</b></h3> </td></tr>" & vbCrLf
               End If
          End If
          temp.MoveNext
        Loop
        junk = junk & "</table>"
     End If
  End If
  Print #ff, junk
  Print #ff, "</textarea>"
  Print #ff, "<input type=""text"" name=""date"" value =""" & DisplayDate & """><br>"
  Print #ff, "<input type=""submit"" value=""Submit"" name=""B1"">"
  Print #ff, "</form></body</html>"
  Close #ff
  
  UpdateS = junk
  Set temp = Nothing
  Set temp2 = Nothing
  'Set temp3 = Nothing
  Dim ITime As Date
  Set temp = DB.OpenRecordset("Select * from profiles where user = '" & CurrentUser.Username & "';", dbOpenDynaset)
  ITime = Now
  temp.Edit
  temp("InternetUpdate") = FixDate(ITime)
  temp.Update
  temp.Close
  Set temp = Nothing
  OpenURL App.path & "\resources\daily\Roaming.htm", vbMaximizedFocus
  Do
    DoEvents
  Loop Until UpdateS = ""
 
  Exit Sub
errhandl:
  MsgBox "Unable to upload roaming information." & vbCrLf & Err.Description, vbOKOnly, ""
End Sub


Private Sub Calendar_DateSelected(Dateclicked As Date)
On Error Resume Next
If Dateclicked <> DisplayDate Then
    Call SaveDay(DisplayDate)
    
    DisplayDate = Dateclicked

    'check if this has already been entered and they are returning
    'if it is a new day then check the database and load that information
    Call DisplayDay(DisplayDate)
End If
End Sub


Private Sub CatSearch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    Dim node1 As MSComctlLib.Node
    Set node1 = CatSearch.SelectedItem
    If Not node1 Is Nothing Then
       If LCase$(Left$(node1.Key, 1)) = "m" Then
    
       ElseIf LCase$(Left$(node1.Key, 1)) = "i" Then
    
       Else
         Call FlexDiet.DropRow(node1.Text)
       End If
    End If
End If
End Sub

Private Sub COKMessage_Click()
If LastMessage = 1 Then
   Call SaveSetting(App.Title, "Settings", "MessageFoods", False)
   MessageFoods = False
End If
If LastMessage = 2 Then
   Call SaveSetting(App.Title, "Settings", "MessageExercise", False)
   MessageExercise = False
End If
If LastMessage = 3 Then
   Call SaveSetting(App.Title, "Settings", "MessageMeals", False)
   MessageMeals = False
End If
If LastMessage = 4 Then
    Call SaveSetting(App.Title, "Settings", "MessageWeb", False)
    MessageWeb = False
End If
PMessage.Visible = False

End Sub

Private Sub DragTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RFM = True
End Sub

Private Sub DragTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RFM = True Then
   X = X + DragTab.Left
   Dim W As Single
   W = ScaleWidth - X - 100
   FMeals.Left = X + 100
   FMeals.Width = W
   Meals.Width = W
   TSearch.Width = W
   CatSearch.Width = W
   Favorites.Width = W
 
   W = Me.ScaleWidth - 150 - W
   FlexDiet.Width = W
   MP.Width = W
   DragTab.Left = X - 50
   DoEvents
Else
  ' DragTab.BorderStyle = 1
  ' Timer1.Enabled = True
End If

End Sub

Private Sub DragTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RFM = True Then
  Form_Resize
End If
RFM = False
End Sub

Private Sub favorites_DblClick()
On Error Resume Next
Dim node1 As MSComctlLib.Node
Set node1 = Favorites.SelectedItem
If Not node1 Is Nothing Then
   If LCase$(Left$(node1.Key, 1)) = "m" Then

   ElseIf LCase$(Left$(node1.Key, 1)) = "i" Then

   Else
     Call FlexDiet.DropRow(node1.Text)
   End If
End If
End Sub

Private Sub CatSearch_DblClick()
On Error Resume Next
Dim node1 As MSComctlLib.Node
Set node1 = CatSearch.SelectedItem
If Not node1 Is Nothing Then
   If LCase$(Left$(node1.Key, 1)) = "m" Then

   ElseIf LCase$(Left$(node1.Key, 1)) = "i" Then

   Else
     Call FlexDiet.DropRow(node1.Text)
   End If
End If
End Sub





Private Sub CEditMeal_Click()
mnuEditMeal_Click
End Sub

Private Sub CNewMeal_Click()
mnuNewMeal_Click
End Sub

Public Sub EasyHover1_Click(Index As Integer)
On Error Resume Next
Dim i As Long
Dim MV As Boolean
For i = 0 To EasyHover1.UBound
   If i <> Index Then Call EasyHover1(i).Release
Next i

On Error GoTo errhandl
DoEvents

'mnuMeals.Visible = False
If MP.Visible And Index <> 0 Then
'  Call SaveDay(DisplayDate)
  Call DisplayDay(DisplayDate)
End If
If Index = 0 And (Not MP.Visible) Then
  'MP.SaveWeek DisplayDate
  Call SaveDay(DisplayDate)
  MP.Clear
  MP.OpenWeek
End If
Dim FMTOP As Single, FMWidth As Single
FMWidth = FMeals.Width
 FMTOP = TSearch.Height + Label2.Height + 10
 Meals.Move 0, FMTOP, FMWidth, FMeals.Height - CNewMeal.Height * 1
 CatSearch.Move 0, FMTOP, FMWidth, FMeals.Height
 Favorites.Move 0, FMTOP, FMWidth, FMeals.Height
 TSearch.ZOrder
'FDeepSearch.Visible = False
Select Case Index
  Case 8 'new meal
     Call frmMeals.Show(vbModal, Me)
  Case 9 'new recipe
     Dim FNewRecipe As New FNewRecipe
     FNewRecipe.DisplayRecipe -1
     Call FNewRecipe.Show(vbModal, Me)
  
  Case 10 'profile
     On Error Resume Next
     Unload FUserSummary
     
     FUserSummary.Show vbModal, Me
     Unload FUserSummary

         
  Case 2 'flexdiet
     
     FlexDiet.Visible = True
     FlexDiet.ZOrder
     Exercise.Visible = False
     Calendar.ZOrder
     ExerciseOnTop = False
     MP.Visible = False
     'mnuMeals.Visible = True
     mnuToolbars.Visible = True
     FMeals.Visible = True
     FMeals.ZOrder
     EasyHover1(7).Visible = True
     EasyHover1(5).Visible = True
     EasyHover1(8).Visible = True
     EasyHover1(9).Visible = True
     If MessageFoods Then
         PMessage.ZOrder
         Label3.Caption = "  You can enter foods by clicking on a cell under the ""Foodname"" column and then entering a search keyword. " & vbCrLf & vbCrLf _
            & " You can also use the category search to the right to find the food you have eaten.  Drag the food over to the calorie counter grid." & vbCrLf & vbCrLf & _
            "The calories that you eat are shown below." & vbCrLf & vbCrLf _
            & "Also notice the calendar at the bottom, this allows you to enter foods for different days."
         PMessage.Visible = True
         LastMessage = 1
     End If
  Case 5 'new food
     Dim fNewFood1 As New FNewFood
     fNewFood1.ShowFood ""
     fNewFood1.Show vbModal, Me
  Case 3 'exercise
     mnuToolbars.Visible = False
     FlexDiet.Visible = True
     Exercise.ZOrder
     Exercise.Visible = True
     ExerciseOnTop = True
     EasyHover1(5).Visible = False
     EasyHover1(8).Visible = False
     EasyHover1(9).Visible = False
     EasyHover1(7).Visible = True
     'mnuMeals.Visible = False
     MP.Visible = False
     FMeals.Visible = False
     If MessageExercise Then
        PMessage.ZOrder
        Label3.Caption = "  Click the cell under ""Exercise"" to enter your exercise keyword.  Once you have selected an exercise from the drop down box, you can enter the amount that you have done by clicking on the appropriate day and " _
        & " and entering the number of minutes that you exercised." & vbCrLf & vbCrLf & "The calories burned per day are shown at the bottom of each day."
        PMessage.Visible = True
        LastMessage = 2
     End If
  Case 4 'journal
     Call FlexDiet.SaveDay(DisplayDate)
     Call Exercise.SaveWeek(DisplayDate)
     
     Unload frmJournal
     DoEvents
     'Call frmJournal.ChangeMode(1)
     frmJournal.Show 'vbModal, Me
     'Unload frmJournal
 Case 0 'menu planner
     FlexDiet.ZOrder
     Exercise.Visible = False
     Calendar.ZOrder
     ExerciseOnTop = False
     'mnuMeals.Visible = True
     FlexDiet.Visible = False
     mnuToolbars.Visible = True
     FMeals.Visible = True
     FMeals.ZOrder
     MP.ZOrder
     MP.Visible = True
     EasyHover1(7).Visible = True
     EasyHover1(5).Visible = True
     EasyHover1(8).Visible = True
     EasyHover1(9).Visible = True
     'TSearch.SelectedItem = 2
     CatSearch.Visible = False
     Meals.Visible = True
     Favorites.Visible = False
      Meals.Move 0, 0, FMeals.Width, FMeals.Height - CNewMeal.Height * 1
      Meals.ZOrder
     Form_Resize
     Calendar.ZOrder
     Label2.Caption = "Click and drag meals to the meal planner."
     If MessageMeals Then
        PMessage.ZOrder
        Label3.Caption = "Meals are entered into the planner by selecting the meal toolbar to the right and dragging the desired meal to the desired day and time." & vbCrLf & vbCrLf _
        & "You can download an example meal by clicking on the globe button to the left, navigating to the ""Diet & Fitness Plans"" link and then clicking the download link. " & vbCrLf & "You then can return to this page and the meals " & _
        "will be displayed in the toolbar to the right."
        LastMessage = 3
        PMessage.Visible = True
     End If
 Case 1 'webpage
     
     Dim junk As String
     On Error Resume Next
     junk = ""
     junk = LCase(Trim(Branding("OpenOnWebsite")))
     'If InStr(1, HUrl, "checklogin", vbTextCompare) <> 0 Or junk = "true" Or junk = "yes" Then
       If FreeVersion Then
           OpenURL "http://www.caloriebalancediet.com/checklogin.asp?member=1&username=" & CurrentUser.Username & "&password=" & CurrentUser.Password & "&source=small", vbMaximizedFocus
       Else
           OpenURL (Branding("BaseWebsite")), vbMaximizedFocus
       End If

     'End If
     
 Case 7 'new exercise
     FNewExercise.Show
     EasyHover1(7).Visible = True
End Select
'FDeepSearch.ZOrder
errhandl:
If DoDebug Then
  Debug.Print Err.Description
  Resume Next
End If
End Sub

Private Sub Exercise_TodaysCalories(NewCals As Single)
On Error Resume Next
  Call FlexDiet.SetExerciseCals(NewCals)
End Sub



Private Sub Favorites_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Dim node1 As MSComctlLib.Node
    Set node1 = Favorites.SelectedItem
    If Not node1 Is Nothing Then
        If LCase$(Left$(node1.Key, 1)) = "m" Then

        ElseIf LCase$(Left$(node1.Key, 1)) = "i" Then

        Else
            Call FlexDiet.DropRow(node1.Text)
        End If
    End If
End If
End Sub

Private Sub FlexDiet_DragDrop(Source As Control, X As Single, Y As Single)
On Error GoTo errhandl
If Left$(Source.Caption, 3) = "~~~" Then
   Dim Parts() As String
   
   Parts = Split(Source.Caption, "~~~")
   
   If X = 0 And Y = 0 Then ' if x and y are 0 then this is because someone clicked the
                           ' make into meal button.  need to replace this meal?
     Call FlexDiet.DropMeal(Source.Caption, True, DisplayDate, , Val(Parts(UBound(Parts))))
   Else
     If (GetKeyState(vbKeyShift) And &H1000) Then
          Call FlexDiet.DropMeal(Source.Caption, False, DisplayDate)
     Else
          Call FlexDiet.DropMeal(Source.Caption, mnuReplace.Checked And mnuReplace.Enabled, DisplayDate)
     End If
   End If
ElseIf Left$(Source.Caption, 3) = "!!!" Then
  ' FlexDiet.SelectRow y
  ' FlexDiet.InsertRows
   Call FlexDiet.DropRow(Replace(Source.Caption, "!!!", ""), Y)
Else

   Call FlexDiet.DropRow(Source.Caption)
  'FDeepSearch.Visible = False
End If
errhandl:
End Sub


Private Sub FlexDiet_LostFocus()
On Error Resume Next
FlexDiet.ForceLoseFocus
End Sub

Private Sub FlexDiet_MakeOfficalMeal(mMealRowNumber As Long, mMealTime As Long)


    On Error GoTo Err_Proc
  MealRowNumber = mMealRowNumber + 1
  MealTime = mMealTime
  MealReplace = True
  mnuMakeMeal_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "FlexDiet_MakeOfficalMeal", Err.Description
    Resume Exit_Proc


End Sub

Private Sub FlexDiet_PresentPopup(PopUp As Object, Un As Variant, X As Single, Y As Single, MealRow As Boolean, MealName As String)

On Error Resume Next
MealTime = MealName
If MealRow Then mnuMakeMeal.Visible = True Else mnuMakeMeal.Visible = False
MealRowNumber = Un
PopUpMenu PopUp, , X, Y

End Sub

Private Sub FlexDiet_RowUpdated()
On Error Resume Next
If DisplayDate < FirstChanged Then FirstChanged = DisplayDate
End Sub

Private Sub FMealClose_Click()


    On Error GoTo Err_Proc
Dim W As Single
FMeals.Visible = False
   W = Me.ScaleWidth - EasyHover1(0).Width - 150
   FlexDiet.Width = W

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "FMealClose_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_Load()
On Error Resume Next
'MsgBox " 1"
 
 MessageFoods = GetSetting(App.Title, "Settings", "MessageFoods", True)
MessageMeals = GetSetting(App.Title, "Settings", "MessageMeals", True)
MessageWeb = GetSetting(App.Title, "Settings", "MessageWeb", True)
MessageExercise = GetSetting(App.Title, "Settings", "MessageExercise", True)
    'frmBrowser.SearchText = "big mac"
    'frmBrowser.Show vbModal, Me

    CPlanID = 1
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    FMeals.Width = GetSetting(App.Title, "Settings", "MealsWidth", FMeals.Width)
    
   Me.Left = 0
   Me.Top = 0
   Me.Width = ScaleX(800, vbPixels, vbTwips)
   Me.Height = ScaleY(600, vbPixels, vbTwips)
     
 '   MsgBox " 2"
    mnuRemind.Checked = GetSetting(App.Title, "Settings", "Remind", False)
    mnuOrder.Checked = GetSetting(App.Title, "settings", "ordermeals", True)
    mnuReplace.Enabled = mnuOrder.Checked
    mnuReplace.Checked = GetSetting(App.Title, "settings", "replace", True)
    
    FlexDiet.OrderMeals = mnuOrder.Checked
    Dim i As Long
    i = GetSetting(App.Title, "Settings", "SearchTab", 1)
    If i = 1 Then
        CatSearch.Visible = True
        Meals.Visible = False
        Favorites.Visible = False
    ElseIf i = 2 Then
        Favorites.Visible = True
        CatSearch.Visible = False
        Meals.Visible = False
    Else
        CatSearch.Visible = False
        Meals.Visible = True
        Favorites.Visible = False
    End If
    TSearch.Tabs(i).Selected = True
  '  MsgBox " 3"
    mnuPercents.Checked = Not (GetSetting(App.Title, "settings", "Percents", False) = "True")
    'FlexDiet.ShowAsPercent = (mnuPercents.Checked)
    Frame1.Move (Me.ScaleWidth - Frame1.Width) / 2, (Me.ScaleHeight - Frame1.Height) / 2
    'MsgBox "FlexDiet"
    
    Call mnuPercents_Click
    Call FlexDiet.SetPopUpMenu(MnuPopUp)
    Call Exercise.SetPopUp(mnuPopup2)
    Call FlexDiet.SetBackGround(BackColor)
    Call Exercise.SetBackGround(BackColor)
    
    FlexDiet.ZOrder
    Calendar.ZOrder
    Meals.Visible = True
    Exercise.Visible = False
  '  MsgBox " 4"
    'Form_Resize
    FMeals.Visible = False
    'FDeepSearch.Visible = False
    
    FMealClose_Click
    
    If GetSetting(App.Title, "settings", "MealsV", True) = True Then
      mnuMeals_Click
    End If
    
    MakeMealList
    'Call Form_Resize
  '  MsgBox " 5"
    Caption = Branding("caption") & " for " & CurrentUser.Username
    Dim f As Font
    Set f = FlexDiet.Font
     f.Name = GetSetting(App.Title, "Settings", "FontName", f.Name)
     f.Size = GetSetting(App.Title, "Settings", "Fontsize", f.Size)
     
    If Not (f.Size = FlexDiet.Font.Size And f.Name = FlexDiet.Font.Name) Then
     Set FlexDiet.Font = f
     Set Exercise.Font = f
    End If
    
    Dim FMTOP As Single, FMWidth As Single
FMWidth = FMeals.Width
 FMTOP = TSearch.Height + Label2.Height + 10
 Meals.Move 0, FMTOP, FMWidth, FMeals.Height - CNewMeal.Height * 1
 CatSearch.Move 0, FMTOP, FMWidth, FMeals.Height
 Favorites.Move 0, FMTOP, FMWidth, FMeals.Height
errhandl:
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
If X > FlexDiet.Left And Y < Exercise.Height Then
'   RFM = True
End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Form_MouseDown", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
If RFM = True Then
   Dim W As Single
   W = Me.ScaleWidth - X - 100
   FMeals.Left = X + 100
   FMeals.Width = W
   Meals.Width = W
   TSearch.Width = W
   CatSearch.Width = W
   Favorites.Width = W
  
   W = Me.ScaleWidth - EasyHover1(0).Width - 150 - W
   FlexDiet.Width = W
   MP.Width = W
End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Form_MouseMove", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
RFM = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Form_MouseUp", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    Dim i As Integer
   
    Call EndIt(Cancel)
    
    
    If Cancel <> 0 Then Exit Sub
   ' If Not DoDebug Then HtmlHelp Me.hWnd, "", HH_CLOSE_ALL, 0&
    NoQuestions = True
    'close all sub forms
    
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", frmMain.Left
        SaveSetting App.Title, "Settings", "MainTop", frmMain.Top
        SaveSetting App.Title, "Settings", "MainWidth", frmMain.Width
        SaveSetting App.Title, "Settings", "MainHeight", frmMain.Height
        
    End If
   
    Call FlexDiet.EndIt
    For i = 1 To TSearch.Tabs.Count
      If TSearch.Tabs(i).Selected Then
         SaveSetting App.Title, "Settings", "SearchTab", i
      End If
    Next i
    SaveSetting App.Title, "Settings", "Remind", mnuRemind.Checked
    SaveSetting App.Title, "settings", "Percents", mnuPercents.Checked
    SaveSetting App.Title, "settings", "replace", mnuReplace.Checked
    SaveSetting App.Title, "settings", "ordermeals", mnuOrder.Checked
    SaveSetting App.Title, "settings", "MealsV", (FMeals.Visible = True)
    SaveSetting App.Title, "Settings", "MealsWidth", FMeals.Width
   
   
    DB.Close
    Set DB = Nothing
    
    
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
  
    End
End Sub

Private Sub Form_Resize()
Dim W As Single, L As Single
On Error Resume Next
Dim T As Single, H As Single
Dim IT As Single

T = EasyHover1(0).Height + 50
W = Me.ScaleWidth - FMeals.Width - DragTab.Width
H = Me.ScaleHeight - T - 150
TSearch.Left = 0
TSearch.Width = FMeals.Width
FlexDiet.Move 0, T, W, H

PMessage.Move (ScaleWidth - PMessage.Width) / 2, (ScaleHeight - PMessage.Height) / 2
MP.Move 0, T, W, H
IT = FlexDiet.ReadOut

Calendar.Move W + 150, Me.ScaleHeight - IT, FMeals.Width, IT
Exercise.Move 0, T, Me.ScaleWidth, H - IT


Dim FMWidth As Single, FMHeight As Single
FMWidth = FMeals.Width
FMHeight = Me.ScaleHeight - IT

FMeals.Move Me.ScaleWidth - FMeals.Width, 0, FMWidth, FMHeight
Meals.Move 0, Meals.Top, FMWidth, FMHeight - CNewMeal.Height * 1
  CNewMeal.Move 0, FMHeight - CNewMeal.Height, FMWidth / 2
  CEditMeal.Move FMWidth / 2, FMHeight - CNewMeal.Height, FMeals.Width / 2

CatSearch.Move 0, Meals.Top, FMWidth, FMHeight
Favorites.Move 0, Meals.Top, FMWidth, FMHeight

 
 
  DragTab.Move FlexDiet.Left + FlexDiet.Width, (ScaleHeight - DragTab.Height) / 2
  DragTab.ZOrder
  DragTab.Visible = True
 Err.Clear
 
 
End Sub









Private Sub FWEB_Click()
'Form2.Show
End Sub


Private Sub Meals_DblClick()


    On Error GoTo Err_Proc
Dim node1 As MSComctlLib.Node
Set node1 = Meals.SelectedItem ' (X, Y)
If Not node1 Is Nothing Then
  If Left(node1.Key, 1) = "K" Then
     CPlanID = Val(Replace(node1.Parent.Key, "I", ""))
  Else
     CPlanID = Val(Replace(node1.Key, "I", ""))
  End If
  Dim tion As String
  tion = "~~~" & node1.Text & "~~~" & Replace(node1.Key, "K", "")
  Dim Parts() As String
  Parts = Split(tion, "~~~")
 
     If (GetKeyState(vbKeyShift) And &H1000) Then
          Call FlexDiet.DropMeal(tion, False, DisplayDate)
     Else
          Call FlexDiet.DropMeal(tion, mnuReplace.Checked And mnuReplace.Enabled, DisplayDate)
     End If
  
End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Meals_DblClick", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Meals_DragDrop(Source As Control, X As Single, Y As Single)


    On Error GoTo Err_Proc

Dim node1 As Node, nodeP As Node, NewPlan As Long
Dim NewMeal As String
NewMeal = ""
Set node1 = Meals.HitTest(X, Y)
If node1 Is Nothing Then
  Set node1 = Meals.SelectedItem
  If node1 Is Nothing Then Exit Sub
End If
If Left(node1.Key, 1) <> "I" Then
   Set node1 = node1.Parent
End If
NewPlan = Val(Right(node1.Key, Len(node1.Key) - 1))
If Len(node1.Key) > Len("I" & NewPlan) Then
   NewMeal = Right$(node1.Key, Len(node1.Key) - Len("i" & NewPlan))
End If
Dim RS As Recordset, TT As String, Parts() As String
Parts = Split(Source.Caption, "~~~")

TT = Parts(2)
'TT = Replace(TT, "'", "''")
Set RS = DB.OpenRecordset("Select * from mealplanner where mealid =" & TT & ";", dbOpenDynaset)
RS.Edit
RS("planid") = NewPlan
If NewMeal <> "" Then RS("meal") = NewMeal
RS.Update
Call MakeMealList
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Meals_DragDrop", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Meals_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next

Dim node1 As Node, nodeP As Node, NewPlan As Long
Dim NewMeal As String
NewMeal = ""
Set node1 = Meals.HitTest(X, Y)
If node1 Is Nothing Then
 
  Exit Sub
End If
node1.Selected = True


End Sub

Private Sub Meals_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Call Meals_DblClick
   KeyCode = 0
End If
End Sub

Private Sub Meals_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
Dim node1 As MSComctlLib.Node
Set node1 = Meals.HitTest(X, Y)
If Not node1 Is Nothing Then
  If Left(node1.Key, 1) = "K" Then
     CPlanID = Val(Replace(node1.Parent.Key, "I", ""))
  Else
     CPlanID = Val(Replace(node1.Key, "I", ""))
  End If
  DropLabel.Caption = "~~~" & node1.Text & "~~~" & Replace(node1.Key, "K", "")
End If
If Button = 1 Then
  
  If node1 Is Nothing Then Exit Sub
  DSM = True
  dsX = X
  dsY = Y
  If Left(node1.Key, 1) <> "K" Then
     DSM = False
  End If
  
Else
  PopUpMenu mnupop3, , FMeals.Left + X, FMeals.Top + Y
  
End If
errhandl:
End Sub

Private Sub Meals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If DSM Then
  If (dsX - X) ^ 2 + (dsY - Y) ^ 2 > 10000 Then
     DropLabel.Move X + FMeals.Left, Y + FMeals.Top
     DropLabel.Visible = True
     DropLabel.Drag vbBeginDrag
     DSM = False
  End If
Else
  dsX = X
  dsY = Y
End If
errhandl:
End Sub

Private Sub Meals_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
  DSM = False
  DropLabel.Visible = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Meals_MouseUp", Err.Description
    Resume Exit_Proc


End Sub



Private Sub catsearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
Dim node1 As MSComctlLib.Node
Set node1 = CatSearch.HitTest(X, Y)
If Not node1 Is Nothing Then
  CatSearchN = node1.Text
  DropLabel.Caption = "!!!" & CatSearchN
End If
If Button = 1 Then
  If node1 Is Nothing Then Exit Sub
  CSM = True
  csX = X
  csY = Y
  If Left(node1.Key, 1) <> "a" Then
     CSM = False
  End If
End If
errhandl:
End Sub
Private Sub favorites_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
Dim node1 As MSComctlLib.Node
Set node1 = Favorites.HitTest(X, Y)
If Not node1 Is Nothing Then
  CatSearchN = node1.Text
  DropLabel.Caption = "!!!" & CatSearchN
End If
If Button = 1 Then
  If node1 Is Nothing Then Exit Sub
  fSM = True
  fsX = X
  fsY = Y
  If Left(node1.Key, 1) <> "a" Then
     fSM = False
  End If
End If
errhandl:
End Sub

Private Sub catsearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If CSM Then
  If (csX - X) ^ 2 + (csY - Y) ^ 2 > 10000 Then
     DropLabel.Move X + FMeals.Left, Y + FMeals.Top
     DropLabel.Visible = True
     DropLabel.Drag vbBeginDrag
     CSM = False
  End If
Else
  csX = X
  csY = Y
End If
errhandl:
End Sub
Private Sub favorites_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandl
If fSM Then
  If (fsX - X) ^ 2 + (fsY - Y) ^ 2 > 10000 Then
     DropLabel.Move X + FMeals.Left, Y + FMeals.Top
     DropLabel.Visible = True
     DropLabel.Drag vbBeginDrag
     fSM = False
  End If
Else
  fsX = X
  fsY = Y
End If
errhandl:
End Sub

Private Sub catsearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
  CSM = False
  DropLabel.Visible = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Meals_MouseUp", Err.Description
    Resume Exit_Proc


End Sub
Private Sub favorites_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    On Error GoTo Err_Proc
  fSM = False
  DropLabel.Visible = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "Meals_MouseUp", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnu_insert_meal_Click()
On Error GoTo errhandl
     Dim node1 As MSComctlLib.Node
   
     Set node1 = Meals.HitTest(dsX, dsY)
     If node1 Is Nothing Then Exit Sub
     
     If Left(node1.Key, 1) = "K" Then
        FlexDiet.DropMeal Replace(node1.Key, "K", ""), mnuReplace.Checked And mnuReplace.Enabled, DisplayDate
     End If
     
errhandl:
End Sub

Private Sub mnuBackup_Click()
On Error GoTo errhandl
Call SaveDay(DisplayDate)
CD.CancelError = True
CD.Filter = "Calorie Balance File (*.cbm)|*.cbm"
CD.ShowSave

Call CopyFile(App.path & "\resources\sr16-2.mdb", CD.Filename, False)

FlexDiet.Changed = False
errhandl:

End Sub


Private Sub mnuBigSearch_Click()

End Sub

Private Sub mnuBugReport_Click()
OpenURL "http://www.caloriebalancediet.com/bug report.asp", vbNormalFocus
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo errhandl
With dlgCommonDialog
   dlgCommonDialog.Flags = cdlCFBoth Or cdlCFEffects
   .CancelError = True
   Dim f As Font
   Set f = FlexDiet.Font
   .FontName = f.Name
   .FontSize = f.Size
   
   dlgCommonDialog.ShowFont
   '.FontBold

   f.Bold = .FontBold
   f.Italic = .FontItalic
   f.Name = .FontName
   f.Size = .FontSize
'   F.Strikethru = .FontStrikethru
 '  F.Underline = .FontUnderline
   SaveSetting App.Title, "Settings", "FontName", f.Name
   SaveSetting App.Title, "Settings", "Fontsize", f.Size
   Set FlexDiet.Font = f
   Set Exercise.Font = f
End With
errhandl:
End Sub

Private Sub mnuContact_Click()


    On Error GoTo Err_Proc
OpenURL "Mailto:support@CalorieBalanceDiet.com"
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuContact_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuContactFirm_Click()


    On Error GoTo Err_Proc
OpenURL Branding("SupportAddress")
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuContactFirm_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuCopyRow_Click()
On Error Resume Next
If FlexDiet.Visible Then Call FlexDiet.Copy

End Sub

Private Sub mnuCopyRow2_Click()
On Error Resume Next
Exercise.Copy
End Sub

Private Sub mnuCutRow_Click()
On Error Resume Next
If FlexDiet.Visible Then Call FlexDiet.Copy
If FlexDiet.Visible Then Call FlexDiet.DeleteRows
End Sub

Private Sub mnuCutRow2_Click()
On Error Resume Next
Exercise.CutRow
End Sub




Private Sub mnuDelete_Click()
On Error Resume Next
If FlexDiet.Visible Then Call FlexDiet.DeleteRows
End Sub

Private Sub mnuDeleteAllRowsInMeal_Click()
On Error Resume Next
If FlexDiet.Visible Then Call FlexDiet.DeleteMeal

End Sub

Private Sub mnuDeleteMeal_Click()
On Error Resume Next
Dim RS As Recordset, TT As String, mealID As Long
Dim ret As VbMsgBoxResult
Dim junk() As String
If DropLabel.Caption = "" Then Exit Sub
junk = Split(DropLabel.Caption, "~~~")

Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & junk(2) & ";", dbOpenDynaset)
mealID = RS("mealid")
ret = MsgBox("Are you sure that you wish to delete " & junk(1), vbYesNo, "")
If ret = vbYes Then
   RS.Delete
   RS.Close
   Call MakeMealList
End If

End Sub

Private Sub mnuDeletePlan_Click()
On Error GoTo errhandl
Dim RS As Recordset, ret As VbMsgBoxResult
Set RS = DB.OpenRecordset("select * from mealplanner where planid=" & CPlanID & " and mealid=-1;", dbOpenDynaset)
ret = MsgBox("Are you sure that you wish to delete " & RS("Mealname"), vbYesNo, "")
If ret = vbYes Then
   RS.Close
   Set RS = DB.OpenRecordset("select * from mealplanner where planid=" & CPlanID & ";", dbOpenDynaset)
   While Not RS.EOF
      RS.Delete
      RS.MoveNext
   Wend
End If
RS.Close
errhandl:
On Error Resume Next
Call MakeMealList
End Sub

Private Sub mnuDeleteRow2_Click()
On Error Resume Next
Exercise.DeleteRow
End Sub

Private Sub mnuDeleteUser_Click()
Dim ret As VbMsgBoxResult
ret = MsgBox("Are you sure you wish to delete " & CurrentUser.Username, vbYesNo)
If ret = vbYes Then

    Dim T As TableDefs, j As TableDef, junk As String
    Dim temp As Recordset, i As Long

    Set T = DB.TableDefs
    On Error GoTo NextTable
    For Each j In T
      junk = j.Name
      For i = 0 To j.Fields.Count - 1
          If LCase$(j.Fields(i).Name) = "user" Then
            GoTo ClearTable
          End If
     Next i
     GoTo NextTable
ClearTable:
   Set temp = DB.OpenRecordset("select * from " & junk & " where user='" & CurrentUser.Username & "';", dbOpenDynaset)
   While Not temp.EOF
     temp.Delete
     temp.MoveNext
   Wend
   temp.Close
NextTable:
   Err.Clear
Next
 Call Module1.OpenUser
 FlexDiet.Changed = False

End If
End Sub


Private Sub mnuEditCopy_Click()
On Error Resume Next
If MP.Visible Then
   MP.Copy
ElseIf (Not ExerciseOnTop) And FlexDiet.Visible Then
   FlexDiet.Copy
  
ElseIf Exercise.Visible Then
   Exercise.Copy
  
End If

End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
If MP.Visible Then
   MP.Cut
ElseIf ExerciseOnTop = False And FlexDiet.Visible Then
   FlexDiet.Copy
   FlexDiet.DeleteRows
ElseIf Exercise.Visible Then
   Exercise.Copy
   Exercise.DeleteRow
End If
End Sub

Private Sub mnuEditMeal_Click()
On Error Resume Next
Dim RS As Recordset, mealID As Long
Dim junk() As String
If DropLabel.Caption = "" Or DropLabel.Caption = "Label1" Then
ErrMsg:
   MsgBox "You must choose a meal to edit" & vbTab & "Please click on a meal above to select it", vbOKOnly, ""
   Exit Sub
End If
junk = Split(DropLabel.Caption, "~~~")

Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & junk(2) & ";", dbOpenDynaset)
If RS Is Nothing Then GoTo ErrMsg
If RS.EOF Then GoTo ErrMsg
mealID = RS("mealid")

If Not RS.EOF Then
   RS.Close
   Call frmMeals.ViewMeal(mealID, CPlanID)
   Call frmMeals.Show
End If
Set RS = Nothing
End Sub



Private Sub mnuEditMealPlanner_Click()


    On Error GoTo Err_Proc
  Call mnuEditMeal_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuEditMealPlanner_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
If MP.Visible Then
   MP.Paste
ElseIf FlexDiet.Visible And (Not ExerciseOnTop) Then
   FlexDiet.Paste
ElseIf Exercise.Visible Then
   Exercise.PasteRow
End If

End Sub

Private Sub mnuEditUser_Click()


    On Error GoTo Err_Proc
  SaveDay DisplayDate
  On Error Resume Next
  Unload FUserSummary
  DoEvents
  
  FUserSummary.Show vbModal, frmMain
  
   Call MsgBox("Program must be restarted for edits to completed.", vbOKOnly, "")
   DB.Close
   DoEvents
   Call Shell(App.path & "\Calorie Balance Tracker.exe", vbNormalFocus)
   End
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuEditUser_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuGetNuts_Click()
  On Error Resume Next
  FNewFood.Show
  DoEvents
  Call FNewFood.ShowFood(FlexDiet.Text)
  
End Sub

Private Sub mnuGetPlans_Click()


    On Error GoTo Err_Proc
DoEvents

OpenURL Branding("PlanWebsite"), vbMaximizedFocus
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuGetPlans_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuGetRecipes_Click()


    On Error GoTo Err_Proc
DoEvents
OpenURL Branding("RecipeWebsite"), vbMaximizedFocus
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuGetRecipes_Click", Err.Description
    Resume Exit_Proc


End Sub


Private Sub mnuHelpAbout_Click()
On Error Resume Next
 OpenURL Branding("basewebsite")
    'frmAbout.Show vbModal, Me
End Sub



Private Sub mnuHelpContents_Click()
  On Error GoTo errhandl
  Call HtmlHelp(frmMain.hWnd, App.HelpFile, 0, ByVal 0)
Exit Sub
errhandl:
MsgBox "Unable to find help file." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub


Private Sub mnuHelpMovies_Click()
On Error Resume Next
OpenURL "http://www.caloriebalancediet.com/HelpMovies/HelpFiles.asp"
End Sub

Private Sub mnuInsert_Click()
On Error Resume Next
If FlexDiet.Visible Then Call FlexDiet.InsertRows
End Sub

Private Sub mnuInsertRow2_Click()
On Error Resume Next
Exercise.InsertRow
End Sub

Public Sub RefreshDay()


    On Error GoTo Err_Proc
    Call DisplayDay(DisplayDate)
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "RefreshDay", Err.Description
    Resume Exit_Proc


End Sub

Public Sub MakeMealList()
On Error Resume Next

Call ClearBlank
Dim RS As Recordset, junk As String, i As Long
Dim node1 As Node
Dim expands As New Collection
Dim ENames As New Collection
For i = 1 To Meals.Nodes.Count
  Set node1 = Meals.Nodes(i)
  expands.Add node1.Expanded, node1.Key
  ENames.Add node1.Key
Next i
Meals.Nodes.Clear
Set RS = DB.OpenRecordset("select * from mealplanner where mealid=-1 and calories=0 and planid>-2 and user='" & CurrentUser.Username & "';", dbOpenDynaset)
Set node1 = Meals.Nodes.Add(, , "I1", "My Meals", 1)
node1.Expanded = True
Meals.Nodes.Add "I1", 4, "I1breakfast", "Breakfast", 2
Meals.Nodes.Add "I1", 4, "I1brunch", "Brunch", 2
Meals.Nodes.Add "I1", 4, "I1lunch", "Lunch", 2
Meals.Nodes.Add "I1", 4, "I1snack", "Snack", 2
Meals.Nodes.Add "I1", 4, "I1dinner", "Dinner", 2
Meals.Nodes.Add "I1", 4, "I1treat", "Treat", 2
While Not RS.EOF
  junk = "I" & RS("planid")
  Meals.Nodes.Add , , junk, RS("Mealname"), 1
  Meals.Nodes.Add junk, 4, junk & "breakfast", "Breakfast", 2
  Meals.Nodes.Add junk, 4, junk & "brunch", "Brunch", 2
  Meals.Nodes.Add junk, 4, junk & "lunch", "Lunch", 2
  Meals.Nodes.Add junk, 4, junk & "snack", "Snack", 2
  Meals.Nodes.Add junk, 4, junk & "dinner", "Dinner", 2
  Meals.Nodes.Add junk, 4, junk & "treat", "Treat", 2

  RS.MoveNext
Wend
RS.Close
Set node1 = Meals.Nodes.Add(, , "I-1", "Family Meals", 1)
node1.Expanded = True
Meals.Nodes.Add "I-1", 4, "I-1breakfast", "Breakfast", 2
Meals.Nodes.Add "I-1", 4, "I-1brunch", "Brunch", 2
Meals.Nodes.Add "I-1", 4, "I-1lunch", "Lunch", 2
Meals.Nodes.Add "I-1", 4, "I-1snack", "Snack", 2
Meals.Nodes.Add "I-1", 4, "I-1dinner", "Dinner", 2
Meals.Nodes.Add "I-1", 4, "I-1treat", "Treat", 2

Set RS = DB.OpenRecordset("select * from mealplanner where mealid<>-1 and (planid =-1 or user='" & CurrentUser.Username & "');", dbOpenDynaset)
While Not RS.EOF
   i = RS("planid")
   If i = 0 Then i = 1
   Meals.Nodes.Add "I" & i & LCase$(RS("Meal")), 4, "K" & RS("Mealid"), RS("Mealname"), 3
   RS.MoveNext
Wend
RS.Close
Set RS = Nothing
On Error Resume Next
For i = 1 To ENames.Count
   Meals.Nodes(ENames(i)).Expanded = expands(ENames(i))
Next i
End Sub

Private Sub mnuLastEntry_Click()


    On Error GoTo Err_Proc
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("select max(date) as maxit " _
  & "from daysinfo where date<#" & FixDate(Today) & "# and user='" & CurrentUser.Username & "';", dbOpenDynaset)
  
  Dim lastentry As Date
  If IsNull(RS("maxit")) = False Then
    lastentry = RS("maxit")
    Calendar.SetDate lastentry
    Calendar_DateSelected lastentry
  Else
    MsgBox "Cannot find any past entries. ", vbOKOnly, ""
  End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuLastEntry_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuMakeMeal_Click()
On Error GoTo errhandl
Dim RS As Recordset, ID As Long
Dim rs2 As Recordset
Dim ms(5) As String, nS(5) As Long, i As Long
Dim tMealTime As String
tMealTime = MealTime 'save the global value for the procedure
ms(0) = "breakfast"
ms(1) = "brunch"
ms(2) = "lunch"
ms(3) = "snack"
ms(4) = "dinner"
ms(5) = "treat"

Call FlexDiet.SaveDay(DisplayDate)
On Error Resume Next
Call ClearBlank
On Error GoTo errhandl

Set RS = DB.OpenRecordset("select max(mealid) as maxit from mealplanner;", dbOpenDynaset)
ID = RS("maxit") + 1
If ID <= 1 Then ID = 2
RS.Close
Set RS = DB.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
RS.AddNew
RS("planid") = 1
RS("mealid") = ID
RS("mealname") = "blank"
RS("user") = CurrentUser.Username
If MealTime <> "" Then
   RS("meal") = ms(Val(MealTime))
End If
MealTime = ""
RS.Update
RS.Close
nS(0) = -200
nS(1) = -201
nS(2) = -202
nS(3) = -203
nS(4) = -204
nS(5) = -205
Set RS = DB.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
Dim aID As Long, Serving As Single, Unit As String, extra As String
Dim tMealRowNumber As Long
tMealRowNumber = MealRowNumber
Call FlexDiet.GetRow(MealRowNumber, aID, Serving, Unit, extra)
While aID > 0
   RS.AddNew
   RS("mealid") = ID
   RS("abbrevid") = aID
   RS("serving") = Serving
   RS("unit") = Unit
   RS.Update
   MealRowNumber = MealRowNumber + 1
   Call FlexDiet.GetRow(MealRowNumber, aID, Serving, Unit, extra)
Wend
RS.Close

Set RS = Nothing


frmMeals.PlanID = 1
MealSaved = False
If frmMeals.ViewMeal(ID, 1) = 0 Then

   frmMeals.Show vbModal, Me
   
   If MealSaved Then
   Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & ID & ";", dbOpenDynaset)
   If (Not RS.EOF) Then
   
   
   Set RS = DB.OpenRecordset("select * from daysinfo where user='" & CurrentUser.Username _
       & "' and date=#" & FixDate(DisplayDate) & "# and " _
       & "daysinfo.order>=" & tMealRowNumber - 1 _
       & " and daysinfo.order<=" & MealRowNumber - 1 & ";", dbOpenDynaset)
Dim MealNumber As Long
While Not RS.EOF
  MealNumber = RS("meal")
  RS.Delete
  RS.MoveNext
Wend
RS.Close
Set RS = Nothing
   
   Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & ID & ";", dbOpenDynaset)
   
     If LCase$(RS("mealname")) <> "blank" Then
        If FlexDiet.OrderMeals Then
           If Trim$(MealTime) = "" Then
             MealTime = tMealTime
           End If
           If Trim$(MealTime) = "" Then
             MealTime = RS("meal")
             For i = 0 To UBound(ms)
               If LCase(MealTime) = LCase(ms(i)) Then
                 MealTime = i
               End If
             Next i
           End If
           
            DropLabel.Caption = "~~~" & RS("mealname") & "~~~" & ID & "~~~" & Val(MealTime)
        Else
            DropLabel.Caption = "~~~" & RS("mealname") & "~~~" & ID & "~~~" & MealNumber + 1
        End If
        Call FlexDiet_DragDrop(DropLabel, 0, 0)
      End If
   End If
   End If
End If
Set RS = Nothing
MealReplace = False
errhandl:
On Error Resume Next
Call ClearBlank

End Sub
Private Sub ClearBlank()
Dim RS As Recordset, ID As Long
Set RS = DB.OpenRecordset("select * from mealplanner where mealname='blank'", dbOpenDynaset)
If Not (RS Is Nothing) Then
    If Not RS.EOF Then
       ID = RS("mealid")
       RS.Delete
       Set RS = DB.OpenRecordset("select * from mealdefinition where mealid=" & ID & ";", dbOpenDynaset)
       If Not (RS Is Nothing) Then
            While Not RS.EOF
                 RS.Delete
                 RS.MoveNext
            Wend
       End If
    End If
End If
End Sub
Private Sub mnuMealCopy_Click()


    On Error GoTo Err_Proc
mnuEditCopy_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuMealCopy_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuMealCut_Click()


    On Error GoTo Err_Proc
mnuEditCut_Click

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuMealCut_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuMealPaste_Click()


    On Error GoTo Err_Proc
mnuEditPaste_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuMealPaste_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuMeals_Click()


    On Error GoTo Err_Proc


FMeals.Visible = True
FMeals.ZOrder
Dim W As Single
   W = Me.ScaleWidth - FMeals.Width - EasyHover1(0).Width - 100
   
   FlexDiet.Width = W
   MP.Width = W
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuMeals_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuNewMeal_Click()
On Error Resume Next
frmMeals.PlanID = CPlanID
Call frmMeals.Show
End Sub

Private Sub mnuNewPlan_Click()
On Error GoTo errhandl
Dim ret As String
ret = InputBox("Please enter new plan name.", "", "")
If ret <> "" Then
    Dim RS As Recordset, PlanID As Long
    Set RS = DB.OpenRecordset("select max(planid) as maxit from mealplanner;", dbOpenDynaset)
    PlanID = RS("maxit") + 1
    Set RS = DB.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
    RS.AddNew
    RS("mealid") = -1
    RS("calories") = 0
    RS("mealname") = ret
    RS("planid") = PlanID
    RS("user") = CurrentUser.Username
    RS.Update
    RS.Close
errhandl:
    On Error Resume Next
    Call MakeMealList
End If
End Sub

Private Sub mnuOpenExercisePlan_Click()
On Error GoTo errhandl
dlgCommonDialog.CancelError = True
dlgCommonDialog.Filter = "Plan (*.cbm; *.mdb) | *.cbm; *.mdb"
dlgCommonDialog.InitDir = App.path & "\resources\plans"
dlgCommonDialog.ShowOpen

Dim vars As Collection
  Set vars = New Collection
   Call REadScriptMod.ReadScript(dlgCommonDialog.Filename, CurrentUser.Username, True)
   MsgBox "Plan has been loaded.", vbOKOnly, ""
Exit Sub
errhandl:

End Sub

Private Sub mnuOpenPersonal_Click()
On Error GoTo errhandl
CD.CancelError = True
CD.Filter = "Calorie Balance File (*.cbm)|*.cbm"
CD.ShowOpen
Call REadScriptMod.UpdateScript(CD.Filename)

errhandl:
End Sub

Private Sub mnuOrder_Click()


    On Error GoTo Err_Proc
mnuOrder.Checked = Not mnuOrder.Checked
mnuReplace.Enabled = mnuOrder.Checked
FlexDiet.OrderMeals = mnuOrder.Checked

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuOrder_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuPasteRow_Click()
On Error Resume Next
FlexDiet.Paste
End Sub

Private Sub mnuPasteRow2_Click()
On Error Resume Next
Exercise.PasteRow
End Sub

Private Sub mnuPercents_Click()


    On Error GoTo Err_Proc

Call SaveDay(DisplayDate)
FlexDiet.Changed = False

If mnuPercents.Checked Then
   mnuPercents.Checked = False
   FlexDiet.ShowAsPercent = False
Else
   mnuPercents.Checked = True
   FlexDiet.ShowAsPercent = True
End If
Call DisplayDay(DisplayDate)
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuPercents_Click", Err.Description
    Resume Exit_Proc


End Sub



Private Sub mnuPrintInstructions_Click()
On Error GoTo errhandl
Dim temp As Recordset, RS As Recordset
  
  Call SaveDay(DisplayDate)
  FlexDiet.Changed = False
  PrintMode = "Preview"
Set temp = DB.OpenRecordset("SELECT [meals].[mealid], [Meals].[User], [Meals].[EntryDate], [MealPlanner].[MealName], [MealPlanner].[Description], [MealPlanner].[Instructions] " _
 & "FROM Meals INNER JOIN MealPlanner ON [Meals].[MealId]=[MealPlanner].[MealID] " _
 & "WHERE ((([Meals].[User])='" & CurrentUser.Username & "') And" _
 & "(([Meals].[EntryDate])=#" & FixDate(DisplayDate) & "#));", dbOpenDynaset)
 
If temp.EOF And temp.BOF Then
   MsgBox "There are no meal instructions", vbOKOnly, ""
   Set temp = Nothing
   Exit Sub
End If
Dim ff As Long
openfile:
On Error Resume Next
ff = FreeFile
Open App.path & "\resources\temp\temp_MealInstruc.html" For Output As #ff
If Err.Number = 55 Then
  Dim OnTime As Boolean
  Close #ff
  Err.Clear
  If Not OnTime Then
    OnTime = True
    GoTo openfile
  End If
End If
On Error GoTo errhandl
Print #ff, "<html><body>"
While Not temp.EOF
            Print #1, "<bigger><b>" & temp("mealname") & "</b></bigger><br><hr>"
            Dim rs2 As Recordset
            Set RS = DB.OpenRecordset("SELECT MealDefinition.*, Abbrev.* " _
              & "FROM MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID=Abbrev.Index " _
              & "WHERE (((MealDefinition.MealID)=" & temp("mealid") & "));", dbOpenDynaset)
            Dim cc As Single
            
            Print #1, "<ul>"
            RTB.TextRTF = temp("description") & " "
            Print #1, "<li>" & Replace(RTB.Text & "<br> ", vbCrLf, "<br>") & "</li>"
            Print #1, "</ul>"
            
            Print #1, "<table><tr><td>Servings</td><td>Units</td><td>Foodname</td>"
            Print #1, "<td>Calories</td><td>Fat (gm)</td><td>Carbs (gm)</td><td>Protein (gm)</td></tr>"
            While Not RS.EOF
            
              Print #1, "<tr><td>" & RS("serving") & "</td><td>" & RS("unit") & "</td><td>" & RS("foodname") & "</td>"
              'Set rs2 = DB.OpenRecordset("select * from weight where index=" & rs("abbrev.index") & " and msre_desc='" & rs("unit") & "';", dbOpenDynaset)
              'If Not rs2.EOF Then
                  cc = TranslateUnitToGrams(RS("abbrev.index"), RS("unit")) / 100 * RS("serving")
               'cc = 1
               'Debug.Print TranslateUnitToGrams(rs("abbrev.index"), rs("unit")), rs2("gm_wgt")
               Dim Cals As Single, fff As Single, cbs As Single, pro As Single
                  Cals = Cals + RS("calories") * cc
                  fff = fff + RS("fat") * cc
                  cbs = cbs + RS("carbs") * cc
                  pro = pro + RS("protein") * cc
                  
                  Print #1, "<td>" & Round(RS("calories") * cc) & "</td>"
                  Print #1, "<td>" & Round(RS("fat") * cc) & "</td>"
                  Print #1, "<td>" & Round(RS("carbs") * cc) & "</td>"
                  Print #1, "<td>" & Round(RS("protein") * cc) & "</td>"
              'End If
              Print #1, "</tr>"
              RS.MoveNext
            Wend
            Print #1, "<tr><td colspan=7><hr></td></tr>"
            Print #1, "<tr><td></td><td></td><td>Totals</td>"
            Print #1, "<td>" & Round(Cals) & "</td><td>" & Round(fff) & "</td><td>" & Round(cbs) & "</td><td>" & Round(pro) & "</td></tr>"

            Print #1, "</table>"
            
            Print #1, "<ul>"
            RTB.TextRTF = temp("instructions") & " "
            Print #1, "<li>" & Replace(RTB.Text & "<br> ", vbCrLf, "<br>") & "</li>"
            Print #1, "</ul>"
            Print #1, "<br><br>"
            temp.MoveNext
Wend
  Print #ff, "</body></html>"
  Close #ff
  DoEvents
  PrintURL = App.path & "\resources\temp\temp_MealInstruc.html"
  OpenURL PrintURL, vbMaximizedFocus
 Exit Sub
errhandl:
On Error Resume Next
 MsgBox "Unable to make print preview." & vbCrLf & Err.Description, vbOKOnly, ""
 If DoDebug Then Resume
 Close #ff
End Sub

Private Sub mnuPrintRecipe_Click()
On Error GoTo errhandl
  Call SaveDay(DisplayDate)
  FlexDiet.Changed = False
  Dim ff As Long
  Dim temp As Recordset, temp2 As Recordset
   PrintMode = "Preview"
  Set temp = DB.OpenRecordset("SELECT [ABBREV].[NDB_No], [DaysInfo].[date], [DaysInfo].[User], [ABBREV].[Index]" _
  & " FROM DaysInfo INNER JOIN ABBREV ON [DaysInfo].[ItemID]=[ABBREV].[Index] " _
  & " WHERE ((([ABBREV].[NDB_No])='-100') And (([DaysInfo].[date])=#" & FixDate(DisplayDate) & "#) And (([DaysInfo].[User])='" & CurrentUser.Username & "'));", dbOpenDynaset)
  
  
  If temp.EOF And temp.BOF Then
     MsgBox "There are no recipes today", vbOKOnly, ""
     Exit Sub
  End If
  ff = FreeFile
  Open App.path & "\resources\temp\temp_recipes.html" For Output As #ff
     Print #ff, "<html><body>"
     While Not temp.EOF
        Set temp2 = DB.OpenRecordset("SELECT RecipesIndex.AbbrevID, RecipesIndex.*, Recipes.*, ABBREV.Foodname " _
        & "FROM (RecipesIndex INNER JOIN Recipes ON RecipesIndex.RecipeID = Recipes.RecipeID) INNER JOIN ABBREV ON Recipes.ItemID = ABBREV.Index " _
        & " WHERE (((RecipesIndex.AbbrevID)=" & temp("index") & "));", dbOpenDynaset)

        If Not (temp2.EOF And temp2.BOF) Then
            Print #1, "<bigger><b>" & temp2("recipename") & "</b></bigger><br>"
            Print #1, "<ul>"
            While Not temp2.EOF
                 Print #1, "<li>" & temp2("servings") & " " & temp2("unit") & " " & temp2("foodname")
                 temp2.MoveNext
            Wend
            temp2.MoveFirst
            Print #1, "</ul>"
            Print #1, "<bigger>Description</bigger><br>"
            RTB.TextRTF = temp2("recipedescription") & " "
            Print #1, Replace(RTB.Text & "<br> ", vbCrLf, "<br>")
            
            Print #1, "<bigger>Instructions</bigger><br>"
            Print #1, "Serves: " & temp2("numberofservings") & "<br>"
            RTB.TextRTF = temp2("recipeinstructions") & " "
            Print #1, Replace(RTB.Text & "<br>", vbCrLf, "<br>") & "<br>"
            
            
            
        End If
        temp2.Close
        Set temp2 = Nothing
        temp.MoveNext
     Wend
     Print #ff, "</body></html>"
  Close #ff
  PrintURL = App.path & "\resources\temp\temp_recipes.html"
  DoEvents
  OpenURL PrintURL, vbMaximizedFocus
Exit Sub
errhandl:
 MsgBox "Unable to make print preview." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub
Private Sub mnuRemind_Click()


    On Error GoTo Err_Proc
    mnuRemind.Checked = Not mnuRemind.Checked
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuRemind_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuReplace_Click()


    On Error GoTo Err_Proc
mnuReplace.Checked = Not mnuReplace.Checked
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuReplace_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuRestore_Click()
On Error GoTo errhandl
Call SaveDay(DisplayDate)

Dim ret As VbMsgBoxResult
ret = MsgBox("This will erase your current database and replace it with the backup." & _
   vbCrLf & "Do you wish to continue.", vbYesNoCancel, "")
If ret = vbYes Then
    CD.CancelError = True
    CD.Filter = "Calorie Balance File (*.cbm)|*.cbm"
    CD.ShowOpen

    Call CopyFile(CD.Filename, App.path & "\resources\sr16-2.mdb", False)
End If
FlexDiet.Changed = False
errhandl:
End Sub

Private Sub MnuSaveExercisePlan_Click()
On Error Resume Next
  Call SaveDay(DisplayDate)
  Form1.PlanMode = 1
  Form1.Show
End Sub

Private Sub mnuSaveMealplan_Click()
On Error Resume Next
Form1.PlanMode = 0
Form1.Show vbModal, Me
End Sub

Private Sub mnuSaveWeek_Click()
On Error Resume Next
Form1.PlanMode = 0
Form1.Show vbModal, Me

End Sub

Private Sub mnuShoppingList_Click()
On Error Resume Next
   FShoppingList.Show
End Sub

Private Sub mnuShowCalories_Click()
On Error Resume Next
If mnuShowCalories.Caption = "Show Calories" Then
  mnuShowCalories.Caption = "Show Units"
  Call Exercise.ShowCalories
Else
  mnuShowCalories.Caption = "Show Calories"
  Call Exercise.ShowUnits
End If
End Sub





Private Sub mnuFileExit_Click()
On Error Resume Next
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()


    On Error GoTo Err_Proc
Call mnuFilePrintPreview_Click

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmMain", "mnuFilePrint_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuFilePrintPreview_Click()
'On Error GoTo errhandl
  Dim i As Long
  PrintMode = "Preview"
 ' PrintPreview.Navigate2 App.path & "\resources\temp\please_wait.htm"
  
'  Exit Sub
If MP.Visible Then
        On Error Resume Next
     Kill App.path & "\resources\temp\temp_Meals.html"
     On Error GoTo errhandl
     Call MP.PrintMeal(App.path & "\resources\temp\temp_Meals.html")
     PrintURL = App.path & "\resources\temp\temp_Meals.html"
     DoEvents
     OpenURL PrintURL, vbMaximizedFocus
  ElseIf Not ExerciseOnTop Then
     On Error Resume Next
     Kill App.path & "\resources\temp\temp_day.html"
     On Error GoTo errhandl
     Call FlexDiet.PrintDay(App.path & "\resources\temp\temp_day.html")
     
     PrintURL = App.path & "\resources\temp\temp_day.html"
     DoEvents
     OpenURL PrintURL, vbMaximizedFocus
     
  Else 'If Exercise.Visible Then
     On Error Resume Next
     Kill App.path & "\resources\temp\temp_exercise.html"
     On Error GoTo errhandl
     
     Call Exercise.PrintExercise(App.path & "\resources\temp\temp_Exercise.html")
     PrintURL = App.path & "\resources\temp\temp_exercise.html"
     DoEvents
     OpenURL PrintURL, vbMaximizedFocus
     
  End If
  DoEvents
  
   Exit Sub
errhandl:
 MsgBox "Unable to make print preview." & vbCrLf & Err.Description, vbOKOnly, ""

End Sub



Private Sub mnuFileSave_Click()
On Error GoTo errhandl
Call SaveDay(DisplayDate)
CD.CancelError = True
CD.Filter = "Calorie Balance File (*.cbm)|*.cbm"
CD.ShowSave
Call REadScriptMod.SaveWeek(CD.Filename, DB, CurrentUser.Username)

FlexDiet.Changed = False
errhandl:

End Sub


Private Sub mnuFileOpen_Click()
On Error Resume Next
 Call SaveDay(DisplayDate)
 Call Module1.OpenUser
 FlexDiet.Changed = False
End Sub

Private Sub mnuFileNew_Click()
On Error Resume Next
 Call SaveDay(DisplayDate)
 CurrentUser.Username = ""
 Call Module1.OpenUser(True)
 FlexDiet.Changed = False
End Sub

Private Sub mnuViewMealPlanner_Click()
On Error Resume Next
  Call mnuViewMeal_Click
  
End Sub

Private Sub mnuWeeksInstructions_Click()


On Error GoTo errhandl
Dim temp As Recordset, RS As Recordset
  Call SaveDay(DisplayDate)
  FlexDiet.Changed = False
  PrintMode = "Preview"
  
Dim ff As Long
openfile:
'On Error Resume Next
ff = FreeFile
Open App.path & "\resources\temp\temp_MealInstruc.html" For Output As #ff
If Err.Number = 55 Then
  Dim OnTime As Boolean
  Close #ff
  Err.Clear
  If Not OnTime Then
    OnTime = True
    GoTo openfile
  End If
End If
Print #ff, "<html><body>"
Dim dd As Date, ddi As Long, ddd As Date
dd = firstSunday(DisplayDate)
For ddi = 0 To 6
  ddd = DateAdd("d", ddi, dd)
 
  Set temp = DB.OpenRecordset("SELECT Meals.MealId, Meals.User, Meals.EntryDate, MealPlanner.MealName, MealPlanner.Description, MealPlanner.Instructions, Meals.EntryDate, Meals.MealNumber " _
   & "FROM Meals INNER JOIN MealPlanner ON Meals.MealId=MealPlanner.MealID " _
   & "WHERE (((Meals.User)='" & CurrentUser.Username & "') AND ((Meals.EntryDate)=#" & FixDate(ddd) & "#)) " _
   & "ORDER BY Meals.EntryDate, Meals.MealNumber;", dbOpenDynaset)



  On Error GoTo errhandl
  Print #ff, "<h2>" & WeekdayName(ddi + 1) & "</h3>"
  While Not temp.EOF
            Print #1, "<bigger><b>" & temp("mealname") & "</b></bigger><br><hr>"
            Dim rs2 As Recordset
            Set RS = DB.OpenRecordset("SELECT MealDefinition.*, Abbrev.* " _
              & "FROM MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID=Abbrev.Index " _
              & "WHERE (((MealDefinition.MealID)=" & temp("mealid") & "));", dbOpenDynaset)
            Dim cc As Single
            
            Print #1, "<ul>"
            RTB.TextRTF = temp("description") & " <Br>"
            Print #1, "<li>" & Replace(RTB.Text & "<br> ", vbCrLf, "<br>") & "</li>"
            Print #1, "</ul>"
            
            Print #1, "<table><tr><td>Servings</td><td>Units</td><td>Foodname</td>"
            Print #1, "<td>Calories</td><td>Fat (gm)</td><td>Carbs (gm)</td><td>Protein (gm)</td></tr>"
            Dim Cals As Single, fff As Single, cbs As Single, pro As Single
            Cals = 0
            fff = 0
            cbs = 0
            pro = 0
            While Not RS.EOF
            
              Print #1, "<tr><td>" & RS("serving") & "</td><td>" & RS("unit") & "</td><td>" & RS("foodname") & "</td>"
              cc = TranslateUnitToGrams(RS("abbrev.index"), RS("unit")) / 100 * RS("serving")
              If Not IsNull(RS("calories")) Then Cals = Cals + RS("calories") * cc
              If Not IsNull(RS("fat")) Then fff = fff + RS("fat") * cc
              If Not IsNull(RS("carbs")) Then cbs = cbs + RS("carbs") * cc
              If Not IsNull(RS("protein")) Then pro = pro + RS("protein") * cc
            
              Print #1, "<td>" & Round(RS("calories") * cc) & "</td>"
              Print #1, "<td>" & Round(RS("fat") * cc, 1) & "</td>"
              Print #1, "<td>" & Round(RS("carbs") * cc, 1) & "</td>"
              Print #1, "<td>" & Round(RS("protein") * cc, 1) & "</td>"
              Print #1, "</tr>"
              RS.MoveNext
            Wend
            Print #1, "<tr><td colspan=7><hr></td></tr>"
            Print #1, "<tr><td></td><td></td><td>Totals</td>"
            Print #1, "<td>" & Round(Cals) & "</td><td>" & Round(fff) & "</td><td>" & Round(cbs) & "</td><td>" & Round(pro) & "</td></tr>"

            Print #1, "</table>"
            
            Print #1, "<ul>"
            RTB.TextRTF = temp("instructions") & " "
            Print #1, "<li>" & Replace(RTB.Text & "<br> ", vbCrLf, "<br>") & "</li>"
            Print #1, "</ul>"
            Print #1, "<br><br>"
            temp.MoveNext
Wend

 
 Next ddi
 Print #ff, "</body></html>"
 Close #ff
  DoEvents
  PrintURL = App.path & "\resources\temp\temp_MealInstruc.html"
  DoEvents
  OpenURL PrintURL, vbMaximizedFocus
Exit Sub
errhandl:
On Error Resume Next
 MsgBox "Unable to make print preview." & vbCrLf & Err.Description, vbOKOnly, ""
 If DoDebug Then Resume
 Close #ff
End Sub

Private Sub MP_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
  MP.DragDrop Source, X, Y
End Sub


Private Sub MP_DropMeal(Caption As String, DropDate As Date, MealNumber As Long)
On Error GoTo errhandl
If Left$(Caption, 3) = "~~~" Then
   Dim Parts() As String, T As Boolean
   Parts = Split(Caption, "~~~")
   T = FlexDiet.OrderMeals
   FlexDiet.OrderMeals = True
   Call FlexDiet.DropMeal(Caption, True, DropDate, False, MealNumber)
   FlexDiet.OrderMeals = T
End If
errhandl:
End Sub



Private Sub mnuViewMeal_Click()
On Error Resume Next
Dim RS As Recordset, TT As String, mealID As Long
Dim rs2 As Recordset
Dim ret As VbMsgBoxResult
Dim junk() As String
   If DropLabel.Caption = "" Then Exit Sub
   junk = Split(DropLabel.Caption, "~~~")

   Set RS = DB.OpenRecordset("select * from mealplanner where mealid=" & junk(2) & ";", dbOpenDynaset)
   mealID = RS("mealid")
   
   
   Set rs2 = DB.OpenRecordset("SELECT MealDefinition.*, Abbrev.Foodname " _
    & "FROM MealDefinition INNER JOIN Abbrev ON MealDefinition.AbbrevID = Abbrev.Index " _
    & "where mealid=" & mealID & ";", dbOpenDynaset)
   TT = ""
   While Not rs2.EOF
     TT = TT & rs2("serving") & " " & rs2("unit") & " " & rs2("Foodname") & vbCrLf
     rs2.MoveNext
   Wend
   TT = TT & vbCrLf & RS("instructions")
   TT = TT & vbCrLf & "Calories = " & Round(RS("Calories"), 1)
   TT = TT & vbCrLf & "Fat(g) = " & Round(RS("fat"), 1)
   TT = TT & vbCrLf & "Carbs(g) = " & Round(RS("Carbs"), 1)
   TT = TT & vbCrLf & "Protein(g) = " & Round(RS("protein"), 1)
   TT = TT & vbCrLf & "fiber(g) = " & Round(RS("fiber"), 1)
   RS.Close
   MsgBox TT, vbOKOnly, ""
End Sub

Private Sub MP_ShowPopUp(Meal As String, X As Single, Y As Single)
On Error Resume Next
   DropLabel.Caption = Meal
   PopUpMenu mnuPopMealPlanner, , X, Y
End Sub

Private Sub MP_StartDrag(MealName As String, X As Single, Y As Single)
On Error Resume Next
     DropLabel.Caption = MealName
     DropLabel.Move X + MP.Left, Y + MP.Top
     DropLabel.Visible = True
     DropLabel.Drag vbBeginDrag
End Sub






Private Sub TSearch_Click()
 Dim i As Long
 i = TSearch.SelectedItem.Index
 If i = 1 Then
    CatSearch.Visible = True
    CatSearch.ZOrder
    Meals.Visible = False
    Favorites.Visible = False
    Label2.Caption = "Click and drag foods to calorie counter."
 ElseIf i = 2 Then
   Favorites.Visible = True
   Favorites.ZOrder
    CatSearch.Visible = False
    Meals.Visible = False
   Label2.Caption = "Click and drag foods to calorie counter."
 Else
    CatSearch.Visible = False
    Meals.Visible = True
    Meals.ZOrder
    Favorites.Visible = False
    Label2.Caption = "Click and drag foods to meal Planner."
 End If
 
 
End Sub

Private Sub UpDater_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
UUrl = URL
End Sub



Private Sub Uploaddaysfood_Click()
On Error Resume Next
Call SaveDay(DisplayDate)
FlexDiet.Changed = False
Call UploadInfo(Today)
End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    '* Purpose: Module scope error handling function
    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)

End Function
