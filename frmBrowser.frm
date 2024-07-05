VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   8190
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WB3 
      Height          =   135
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
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
   Begin VB.TextBox label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   2640
      TabIndex        =   1
      Text            =   "Please Wait..."
      Top             =   600
      Width           =   3975
   End
   Begin VB.Frame Nutframe 
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   4935
      Begin SHDocVwCtl.WebBrowser WB 
         Height          =   3735
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   4335
         ExtentX         =   7646
         ExtentY         =   6588
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
      Begin VB.CommandButton Command3 
         Caption         =   "Save Food"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   7080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   5160
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rename Selected Food"
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save New Food"
         Height          =   495
         Left            =   5160
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin MSComctlLib.TreeView DeepSearch 
         Height          =   4215
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   7435
         _Version        =   393217
         Style           =   7
         ImageList       =   "Folderlist1"
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList Folderlist1 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmBrowser.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":0552
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":0A54
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"frmBrowser.frx":0F36
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   4575
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Community Database"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet Search"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Category Search"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   7695
      Begin SHDocVwCtl.WebBrowser brwWebBrowser 
         Height          =   2055
         Left            =   0
         TabIndex        =   21
         Top             =   1320
         Width           =   3375
         ExtentX         =   5953
         ExtentY         =   3625
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
      Begin VB.CommandButton Command10 
         Caption         =   "Transfer Nutrition Info"
         Height          =   975
         Left            =   6480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBrowser.frx":1073
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2520
         MaskColor       =   &H00000000&
         Picture         =   "frmBrowser.frx":2051
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBrowser.frx":3263
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBrowser.frx":46E5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBrowser.frx":5B67
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.PictureBox picAddress 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   13530
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   13530
         Begin VB.ComboBox cboAddress 
            Height          =   315
            Left            =   45
            TabIndex        =   3
            Text            =   "¯¯END!"
            Top             =   300
            Width           =   3795
         End
         Begin VB.Label lblAddress 
            Caption         =   "&Address:"
            Height          =   255
            Left            =   45
            TabIndex        =   4
            Tag             =   "&Address:"
            Top             =   60
            Width           =   3075
         End
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   2670
         Top             =   2325
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":6AD9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":6DBB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":709D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":737F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":7661
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":7943
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8880
      Top             =   240
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save Food"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StartingAddress As String
Public SearchText As String
Dim mbDontNavigateNow As Boolean
Public SelIndex As String
Public SelText As String
Private lURL As String

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub LoadDeepSearch()
On Error GoTo errhandl
    Dim i As Long, j As String, J2 As String
    
    Dim node1 As Node
    Dim expands As New Collection
    Dim ENames As New Collection
    For i = 1 To DeepSearch.Nodes.Count
      Set node1 = DeepSearch.Nodes(i)
      expands.Add node1.Expanded, node1.Key
      ENames.Add node1.Key
    Next i
    
    Dim FoodgroupS  As Recordset
    Set FoodgroupS = DB.OpenRecordset("select * from foodgroups where parentnumber=-10;", dbOpenDynaset)
    
    DeepSearch.Nodes.Clear
    While Not FoodgroupS.EOF
       i = FoodgroupS("catnumber")
       If i < 0 Then
         j = "M" & i
         Call DeepSearch.Nodes.Add(, , j, FoodgroupS("category"), 1)
       Else
         j = "M" & Format(i, "0000")
         Call DeepSearch.Nodes.Add(, , j, FoodgroupS("category"), 1)
         J2 = "I" & Format(i, "0000")
         Call DeepSearch.Nodes.Add(j, 4, J2, "General", 2)
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
       Call DeepSearch.Nodes.Add(j, 4, "I" & J2, FoodgroupS("category"), 2)
       FoodgroupS.MoveNext
    Wend
    
    Dim rs As Recordset
    Set rs = DB.OpenRecordset("Select * from abbrev order by foodgroup,foodname;", dbOpenDynaset)
    While Not rs.EOF
      j = rs("foodgroup") & ""
      j = Format(j, "0000")
      DeepSearch.Nodes.Add "I" & j, 4, "a" & rs("index"), rs("foodname"), 3
      j = ""
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    On Error Resume Next
    For i = 1 To ENames.Count
       DeepSearch.Nodes(ENames(i)).Expanded = expands(ENames(i))
    Next i
    
    Exit Sub
errhandl:
    Resume Next

End Sub


Private Function SaveFood() As Boolean 'return true to cancel save, false if saved
Dim cc As Single, i As Long
Dim Grams As Single

On Error GoTo errhandl
Dim General
Set General = WB.document.theGenerals
Grams = Val(General.Grams.Value)
If General.Foodname.Value = "" Then
  Call General.Foodname.insertAdjacentHTML("beforeBegin", "<big style=""color: rgb(255, 0, 0);""><big>*</big></big>")
  MsgBox "Please enter the foodname.", vbOKOnly, ""
  SaveFood = True
  Exit Function
End If
SelText = General.Foodname.Value
If Val(General.amount.Value) = 0 Or General.Unit.Value = "" Then
   Call General.amount.insertAdjacentHTML("beforeBegin", "<big style=""color: rgb(255, 0, 0);""><big>*</big></big>")
   MsgBox "Please enter serving information", vbOKOnly, ""
   SaveFood = True
   Exit Function
End If

If Grams = 0 Then
    Call General.Grams.insertAdjacentHTML("beforeBegin", "<big style=""color: rgb(255, 0, 0);""><big>*</big></big>")
  MsgBox "Please enter number of grams in serving" & vbCrLf & "(estimate if needed.  If you cannot guess enter 100)", vbOKOnly, ""
  SaveFood = True
  Exit Function
End If

If Val(WB.document.Nutrients.Calories.Value) = 0 Then
  Call WB.document.Nutrients.Calories.insertAdjacentHTML("beforeBegin", "<big style=""color: rgb(255, 0, 0);""><big>*</big></big>")
  MsgBox "Please enter the number of calories in serving", vbOKOnly, ""
  SaveFood = True
  Exit Function
End If
Dim OO, found As Boolean, O
Set OO = WB.document.getElementsByName("foodgroup")
For Each O In OO(0).Options
     If O.Selected Then
        found = True
        Exit For
     End If
Next
Set O = Nothing
Set OO = Nothing
If Not found Then
  Call WB.document.theGenerals.insertAdjacentHTML("beforeBegin", "<big style=""color: rgb(255, 0, 0);""><big>*</big></big>")
  MsgBox "Please enter the food group of the new food.", vbOKOnly, ""
  SaveFood = True
  Exit Function
End If

SaveFood = False

cc = 100 / Grams

Dim temp2 As Recordset
Dim sTemp As Recordset
Dim AbbrevID As Long

  Dim Ideals As Recordset
  Set Ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)

  If Ideals.EOF And Ideals.BOF Then
    Set Ideals = DB.OpenRecordset("Select * from ideals;", dbOpenDynaset)
    
  
  End If
  
  Set sTemp = DB.OpenRecordset("SELECT * FROM ABBREV where foodname ='" _
                      & Replace(SearchText, "'", "''") & "';", dbOpenDynaset)
                      
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
  
  SelIndex = AbbrevID
  
  Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & AbbrevID & "));", dbOpenDynaset)
  Do While Not temp2.EOF
     temp2.Delete
     temp2.MoveNext
  Loop
  temp2.AddNew

  sTemp.Fields("Usage") = 10
  sTemp.Fields("Foodname") = General.Foodname.Value
  
  Dim junk As String, junk2
  Dim Element, Elements
  Dim Nutrients
  For i = 0 To 1
    If i = 0 Then
       Set Nutrients = WB.document.Nutrients
    Else
       Set Nutrients = WB.document.vitamins
    End If
    Set Elements = Nutrients.Elements
    On Error Resume Next
    Dim jjj As Long, Switch As Boolean
    For Each Element In Elements
        
        junk2 = 0
        junk = CleanList(Element.name)
        junk2 = Element.Value
        If LCase$(junk) = "vitamin a" Then Switch = True
        If i = 0 Then
           sTemp(junk) = Val(junk2) * cc
        Else
           sTemp(junk) = Val(junk2) / 100 * Ideals(junk) * cc
        End If
        Set Element = Nothing
        jjj = jjj + 1
    Next
    Set Elements = Nothing
  Next i
  

  sTemp("Foodgroup") = Trim$(General.FoodGroup.Value & " ")
  
  sTemp.Update

  temp2.Fields("Gm_Wgt") = Grams
  temp2.Fields("Amount") = Val(ConvertFractionToDecimal(General.amount.Value))
  temp2.Fields("Msre_desc") = General.Unit.Value
  temp2.Fields("Index") = AbbrevID
  temp2.Update
  Set temp2 = Nothing
  Set sTemp = Nothing
  Call frmMain.LoadDeepSearch
  MsgBox "Food was saved successfully", vbOKOnly, ""
errhandl:
  
End Function

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

Private Sub brwWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)


    On Error GoTo Err_Proc
   'WB3.Navigate2 URL
   'Call GetFoodData
   lURL = URL
   If InStr(1, URL, "yahoo", vbTextCompare) <> 0 Then
      Dim ele, e
      Set ele = brwWebBrowser.document.getElementsByTagName("input")
      For Each e In ele
        If LCase$(e.Type) <> "hidden" And LCase$(e.Type) <> "submit" Then
           e.Value = "calories in " & SearchText
        End If
      Next
   
   End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "brwWebBrowser_DocumentComplete", Err.Description
    Resume Exit_Proc


End Sub

Private Sub LoadFood(Index As String)


    On Error GoTo Err_Proc
    SelIndex = Index
    Dim temp As Recordset, temp2 As Recordset, Ideals As Recordset
    
    Set temp = DB.OpenRecordset("SELECT * FROM ABBREV where index=" & Index & ";", dbOpenDynaset)
    
    Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & Index & "));", dbOpenDynaset)
    Set Ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)
    
    
    Dim nutinfo As New exCollection, cc As Single, i As Long
    
    cc = temp2("gm_wgt") / 100
    
    For i = 0 To temp.Fields.Count - 1
       nutinfo.Add Round(Val(temp.Fields(i).Value & " ") * cc, 2), temp.Fields(i).name
    Next i
    
    nutinfo.Add temp("foodname"), "foodname"
    nutinfo.Add temp2("amount"), "amount"
    nutinfo.Add temp2("Msre_Desc"), "unit"
    nutinfo.Add temp2("gm_wgt"), "grams"
    
   Dim html2
   
   Set html2 = WB.document
   Dim ele, junk As String, junk2 As String
   Dim e
   Set ele = html2.getElementsByTagName("input")
   
   For i = 1 To nutinfo.Count
     junk = nutinfo.ItemName(i)
     junk2 = LCase$(Replace(junk, " ", "_"))
     For Each e In ele
        If junk2 = LCase(e.name) Then
          e.Value = nutinfo(i)
        End If
        
     Next
   Next i
   Dim OO, O
   
   Set OO = html2.getElementsByTagName("select")
   For Each O In OO(0).Options
     If O.Value = temp("foodgroup") Then
        O.Selected = True
     End If
   Next
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "LoadFood", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command1_Click()


    On Error GoTo Err_Proc
Dim Cancel As Boolean
Cancel = SaveFood
If Not Cancel Then Unload Me

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command1_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command10_Click()


    On Error GoTo Err_Proc
  Label2.Visible = True
  Label2.ZOrder 0
  DoEvents
  WB3.Navigate2 lURL
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command10_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command2_Click()


    On Error GoTo Err_Proc
Dim temp As Recordset
Set temp = DB.OpenRecordset("select * from abbrev where index=" & SelIndex & ";", dbOpenDynaset)
If SelIndex = "" Then
   MsgBox "Please select a food to be renamed", vbOKOnly, ""
   Exit Sub
End If
If WB.document.theGenerals.Foodname.Value = "" Then
   MsgBox "Please enter a name into the foodname box.", vbOKOnly, ""
   Exit Sub
End If
If temp.EOF = False Then
  temp.Edit
  temp("foodname") = WB.document.theGenerals.Foodname.Value
  temp.Update
End If
SelText = WB.document.theGenerals.Foodname.Value
temp.Close
Set temp = Nothing
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command2_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command3_Click()


    On Error GoTo Err_Proc
Dim Cancel As Boolean
Cancel = SaveFood
If Not Cancel Then Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command3_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command4_Click()


    On Error GoTo Err_Proc
SelIndex = -1
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command4_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command5_Click()


    On Error GoTo Err_Proc
SelIndex = -1
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command5_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command6_Click()


    On Error GoTo Err_Proc

            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command6_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command7_Click()


    On Error GoTo Err_Proc
   
            brwWebBrowser.GoForward
        
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command7_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command8_Click()


    On Error GoTo Err_Proc
brwWebBrowser.GoBack
     
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command8_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command9_Click()


    On Error GoTo Err_Proc

            brwWebBrowser.Navigate2 "http://www.yahoo.com"
        
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "Command9_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub DeepSearch_DblClick()


    On Error GoTo Err_Proc
  Dim node1 As MSComctlLib.Node
  Set node1 = DeepSearch.SelectedItem
  If Left(node1.Key, 1) = "a" Then
    Call LoadFood(Right$(node1.Key, Len(node1.Key) - 1))
  End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "DeepSearch_DblClick", Err.Description
    Resume Exit_Proc


End Sub

Private Sub DeepSearch_KeyUp(KeyCode As Integer, Shift As Integer)


    On Error GoTo Err_Proc
 If KeyCode = 13 Then
  Dim node1 As MSComctlLib.Node
  Set node1 = DeepSearch.SelectedItem
  If Left(node1.Key, 1) = "a" Then
    Call LoadFood(Right$(node1.Key, Len(node1.Key) - 1))
  End If
 End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "DeepSearch_KeyUp", Err.Description
    Resume Exit_Proc


End Sub

Private Sub DeepSearch_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


    On Error GoTo Err_Proc
  
  Dim node1 As MSComctlLib.Node
   
  Set node1 = DeepSearch.HitTest(x, y)
  If node1 Is Nothing Then Exit Sub
  
  'DropLabel.Caption = node1.Text
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "DeepSearch_MouseUp", Err.Description
    Resume Exit_Proc


End Sub
Public Sub ClearAll()
Label2.Visible = False
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Label2.Visible = False
    Dim junk As String
    junk = GetSetting(App.Title, "Settings", "IntroInternetUsers", "")
       
    If InStr(1, junk, CurrentUser.Username, vbTextCompare) = 0 Then
    
       junk = junk & "," & CurrentUser.Username
       SaveSetting App.Title, "Settings", "IntroInternetUsers", junk
       StartingAddress = App.path & "\resources\help\QuickInternetStart.htm"
    Else
       
       StartingAddress = Branding("CalorieCounter")
    End If
    
    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
    WB.Navigate2 App.path & "\resources\temp\internetfood.htm"
    Call Form_Resize
End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)


    On Error GoTo Err_Proc
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "brwWebBrowser_NavigateComplete", Err.Description
    Resume Exit_Proc


End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    
On Error Resume Next
    Dim W As Single, H As Single
    W = Me.ScaleWidth - WB.Width - 100
    H = Me.ScaleHeight - 100
    Frame1.Move 0, TabStrip1.Height, W, H
    Frame2.Move 0, TabStrip1.Height, W, H
    'WB2.Move 0, TabStrip1.Height, W, H
    Label1.Left = W - Label1.Width
    Command1.Left = W - Label1.Width
    Command2.Left = W - Label1.Width
    Command5.Left = W - Label1.Width
    
    DeepSearch.Move 0, 0, W - Label1.Width, H - 100
    W = W - 100
    H = H - 100 - TabStrip1.Height
    cboAddress.Width = W - Command10.Width
    
    brwWebBrowser.Width = W 'Me.ScaleWidth - 100 - WB.width
    brwWebBrowser.Height = H - (brwWebBrowser.Top) - 100
    
    TabStrip1.Move 0, 0, W
    
    Nutframe.Move Me.ScaleWidth - WB.Width, 0, WB.Width, Me.ScaleHeight - 100
    WB.Move 0, 0, WB.Width, Me.ScaleHeight - 100 - 2 * Command3.Height
    Command3.Top = Me.ScaleHeight - 2 * Command3.Height
    Command4.Top = Me.ScaleHeight - 2 * Command3.Height
    Command3.Left = WB.Width - Command4.Width - 200 - Command3.Width
    Command4.Left = WB.Width - Command4.Width - 100
    
    
    Command10.Move Me.ScaleWidth - WB.Width - Command10.Width, 0
    
    Label2.Move Me.ScaleWidth / 2 - Label2.Width / 2, (Me.ScaleHeight - Label2.Width) / 2
    
End Sub

Private Sub mnuExit_Click()


    On Error GoTo Err_Proc
Command4_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "mnuExit_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub mnuHelp_Click()
On Error GoTo errhandl
HelpWindowHandle = htmlHelpTopic(frmMain.hWnd, HelpPath, _
         0, "newHTML/InternetNewFood.htm")
 Exit Sub
errhandl:
 MsgBox "Cannot find help file." & vbCrLf & Err.Description, vbOKOnly, ""
End Sub

Private Sub mnuSave_Click()


    On Error GoTo Err_Proc
 Command3_Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "mnuSave_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub TabStrip1_Click()
Dim i As Long
On Error Resume Next
i = TabStrip1.SelectedItem.Index
Select Case i
  Case 2
     Frame1.Visible = True
     Frame2.Visible = False
     brwWebBrowser.Navigate2 Branding("DefaultSearch")
  Case 1
     Frame1.Visible = True
     Frame2.Visible = False
     brwWebBrowser.Navigate Branding("caloriecounter")
  Case 3
     Frame1.Visible = False
     Frame2.Visible = True
     Call LoadDeepSearch
     'Call MakeCategories
     
  
End Select
End Sub

Private Sub timTimer_Timer()


    On Error GoTo Err_Proc
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "timTimer_Timer", Err.Description
    Resume Exit_Proc


End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    'On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            
    End Select

End Sub
Private Sub GetUSDAInfo(html)
   Dim ele
   Dim Table
   Dim r, c, i As Long, j As Long
   Dim nutGrid()
   Dim nutinfo As New exCollection
   Set ele = html.getElementsByTagName("table")
   
   Set Table = ele(0)
   For i = 0 To Table.Rows.length - 1
      ele = Table.Rows(i).cells.length
      If ele > j Then j = ele
   Next i
   ReDim Preserve nutGrid(i, j)
   For i = 0 To Table.Rows.length - 1
      Set r = Table.Rows(i)
      For j = 0 To r.cells.length - 1
          On Error Resume Next
          'nutinfo.Add r.cells(r.cells.length - 1).innerText, r.cells(0).innerText
          nutGrid(i, j) = r.cells(j).innerText
      Next j
      
   Next i
   For i = 0 To UBound(nutGrid, 1)
      If Trim$(LCase$(nutGrid(i, 0) & nutGrid(i, 1))) = "energy kcal" Then
         nutGrid(i, 0) = "calories"
      End If
   Next i
   
   Dim trans As New Collection
   trans.Add "fat", "Total lipid (fat)"
   trans.Add "Carbohydrate", "Carbohydrate, by difference"
   trans.Add "fiber", "Fiber, total dietary"
   trans.Add "sugar", "Sugars, total"
   trans.Add "vitamin e", "Vitamin E (alpha-tocopherol)"
   trans.Add "vitamin k", "Vitamin K (phylloquinone)"
   trans.Add "saturated fat", "Fatty acids, total saturated"
   trans.Add "monounsaturated fat", "Fatty acids, total monounsaturated"

   trans.Add "polyunsaturated fat", "Fatty acids, total polyunsaturated"
   trans.Add "beta__carotene", "Carotene, beta"
   trans.Add "alpha__carotene", "Carotene, alpha"
   trans.Add "vitamin b12", "Vitamin B-12"
   trans.Add "vitamin b6", "Vitamin B-6"
   
   Dim html2, junk As String, junk2 As String, e, junk3() As String
   On Error Resume Next
   For i = 0 To UBound(nutGrid, 1)
        junk = Trim$(nutGrid(i, 0))
        For j = 1 To trans.Count
           junk = trans(junk)
        Next j
        junk3 = Split(junk, ",")
        nutinfo.Add nutGrid(i, UBound(nutGrid, 2) - 1), Trim$(junk3(0))
       
   Next i
   
   
   Set html2 = WB.document
   Set ele = html2.getElementsByTagName("input")
   
   For i = 1 To nutinfo.Count
     junk = nutinfo.ItemName(i)
     junk2 = LCase$(Replace(junk, " ", "_"))
     For Each e In ele
        If junk2 = LCase$(e.name) Then
          e.Value = nutinfo(i)
        End If
        
     Next
   Next i
   Dim ele2
   Dim inpt
   Dim Ideals As Recordset
   Set Ideals = DB.OpenRecordset("Select * from ideals where user='" & CurrentUser.Username & "';", dbOpenDynaset)

   If Ideals.EOF And Ideals.BOF Then
     Set Ideals = DB.OpenRecordset("Select * from ideals;", dbOpenDynaset)
   End If
   Set ele2 = html2.vitamins
   For Each inpt In ele2.Elements
    
     inpt.Value = Round(100 * inpt.Value / Ideals(CleanList(inpt.name)), 2)
     
   Next
End Sub
Private Sub GetFoodData()
   Dim html, longjunk As String, junk As String
   Dim Parts() As String
   Dim e
   Dim e2
   Dim e3
   Dim td
   Dim ele As Object
   
   On Error Resume Next
   Set html = WB3.document
   If html Is Nothing Then Exit Sub
  ' On Error GoTo 0
   If InStr(1, html.URL, "www.nal.usda.gov", vbTextCompare) <> 0 Then
      Call GetUSDAInfo(brwWebBrowser.document)
      Exit Sub
   End If
   Set ele = html.getElementsByTagName("input")
   For Each e In ele
       
       If LCase(e.Type) <> "hidden" Then
            Call e.insertAdjacentText("BeforeBegin", e.Value)
       Else
       End If
       Call e.parentNode.removeChild(e)
   Next
   Set ele = html.getElementsByTagName("select")
   For Each e2 In ele
      
      Set e3 = e2.Options(e2.selectedIndex)
      Call e2.insertAdjacentText("BeforeBegin", e3.Text)
      Call e2.parentNode.removeChild(e2)
   Next
   
   Set ele = html.getElementsByTagName("td")
   longjunk = ""
   For Each td In ele
      
      junk = td.innerHTML
      If InStr(1, junk, "td", vbTextCompare) = 0 Then
        longjunk = longjunk & td.innerText & vbCrLf
      End If
   Next
   Dim SP As New Collection 'Seach Phrases
   SP.Add "Calories from fat"
   SP.Add "Calories from Carbohydrate"
   SP.Add "Calories from Protein"
   SP.Add "Calories from Alcohol"
   SP.Add "calories"
   SP.Add "calorie count"
   SP.Add "calorie"
  ' SP.Add "Serving size"
  ' SP.Add "serving:"
  ' SP.Add "Serving"
   SP.Add "Monounsaturated fat"
   SP.Add "Polyunsaturated fat"
   SP.Add "Polyunsat. Fat"
   SP.Add "Monounsat. Fat"
   SP.Add "saturated fat"
   SP.Add "Total Fat"
 
   
   SP.Add "fat"
   SP.Add "Total Carbohydrates"
   SP.Add "total carbohydrate"
   SP.Add "Carbohydrates"
   SP.Add "Sodium"
   SP.Add "Dietary Fiber"
   SP.Add "fiber"
   SP.Add "Sugars"
   SP.Add "sugar"
   SP.Add "total protein"
   SP.Add "Protein"
   SP.Add "Vitamin A"
   SP.Add "vitamin c"
   SP.Add "calcium"
   SP.Add "iron"
   SP.Add "Cholesterol"
   SP.Add "Vitamin E"
   SP.Add "vitamin k"
   SP.Add "vitamin b-12"
   SP.Add "vitamin b12"
   SP.Add "vitamin b-6"
   SP.Add "vitamin b6"
   SP.Add "Pantothenic acid"
   SP.Add "Niacin"
   SP.Add "Riboflavin"
   SP.Add "Thiamin"
   SP.Add "Selenium"
   SP.Add "Manganese"
   SP.Add "Copper"
   SP.Add "Zinc"
   SP.Add "Potassium"
   SP.Add "Phosphorus"
   SP.Add "Magnesium"
   SP.Add "Folate"
   
   Dim nutinfo As New exCollection
   Dim junk2 As String, junk3 As String
   Dim j As Long, i As Long, k As Long
   Parts = Split(longjunk, vbCrLf)
   For j = 1 To SP.Count
   junk = SP(j)
     For i = 0 To UBound(Parts)
       k = InStr(1, Parts(i), junk, vbTextCompare)
       If k <> 0 Then
          Parts(i) = Right(Parts(i), Len(Parts(i)) - k)
          junk2 = ""
          If Val(StripLetters(Parts(i))) <> 0 Then
            junk2 = Val(StripLetters(Parts(i)))
            Parts(i) = ""
          ElseIf Val(StripLetters(Parts(i + 1))) <> 0 Then
            junk2 = Val(StripLetters(Parts(i + 1)))
            Parts(i) = ""
          End If
          If junk2 <> "" And junk2 <> "0" Then
             junk3 = ""
             junk = Replace(junk, "-", "")
             junk = Trim$(Replace(junk, "total", "", , , vbTextCompare))
             junk = Replace(junk, "Carbohydrates", "Carbohydrate", , , vbTextCompare)
             junk = Replace(junk, "dietary fiber", "fiber", , , vbTextCompare)
             junk = Replace(junk, "sugars", "sugar", , , vbTextCompare)
             junk = Replace(junk, "Monounsat. Fat", "Monounsaturated fat", , , vbTextCompare)
             junk = Replace(junk, "Polyunsat. Fat", "polyunsaturated fat", , , vbTextCompare)
             junk3 = nutinfo(junk)
             If junk3 = "" Then nutinfo.Add junk2, junk
          End If
       End If
     Next i
   Next j
   
   junk = ""
   junk = nutinfo("calories")
   If junk = "" Then
      i = 1
g:
      j = InStr(i, longjunk, "calories", vbTextCompare)
      If j <> 0 Then
         junk = Right$(longjunk, Len(longjunk) - (j + 8))
         If Val(junk) <> 0 Then
            nutinfo.Add Val(junk), "calories"
         Else
            i = j + 1
            GoTo g
         End If
      End If
   
   End If
   
   j = InStr(1, longjunk, "serving", vbTextCompare)
   If j <> 0 Then
      i = InStr(j + 1, longjunk, vbCrLf)
      junk = Mid$(longjunk, j, i - j + 1)
      junk = Replace(junk, "serving:", "", , , vbTextCompare)
      junk = Replace(junk, "serving size", "", , , vbTextCompare)
      junk = Replace(junk, "size", "", , , vbTextCompare)
      junk = Trim$(junk)
      If Val(junk) <> 0 Then
         nutinfo.Add Val(junk), "amount"
         junk = Trim$(junk)
         j = InStr(1, junk, " ")
         junk = Trim$(Right$(junk, Len(junk) - j))
         'junk = Trim$(Replace(junk, Val(junk), "", 1, 1))
      End If
      
      j = InStr(1, junk, "(")
      If j <> 0 Then
        i = InStr(j, junk, ")")
        junk3 = Mid$(junk, j, i - j + 1)
        junk2 = Replace(Replace(Mid$(junk, j, i - j + 1), "(", " "), ")", " ")
        If InStr(1, junk2, "gm", vbTextCompare) <> 0 Or InStr(1, junk2, "g", vbTextCompare) <> 0 Then
           If Val(junk2) <> 0 Then
              nutinfo.Add Val(junk2), "grams"
           End If
        End If
      End If
      junk = Trim$(Replace(junk, junk3, ""))
      If junk <> "" Then
         nutinfo.Add junk, "unit"
      End If
   End If
   nutinfo.Add SearchText, "foodname"
   
   Dim html2
   
   Set html2 = WB.document
   Set ele = html2.getElementsByTagName("input")
   
   For i = 1 To nutinfo.Count
     junk = nutinfo.ItemName(i)
     junk2 = LCase$(Replace(junk, " ", "_"))
     For Each e In ele
        If junk2 = LCase(e.name) Then
          e.Value = nutinfo(i)
        End If
        
     Next
   Next i
   
   
   
 
'End If
End Sub

Private Function StripLetters(ByVal inPhrase As String) As String


    On Error GoTo Err_Proc
   Dim j As String, jj As String, i As Long
   For i = 1 To Len(inPhrase)
      jj = Mid$(inPhrase, i, 1)
      If (jj >= "0" And jj <= "9") Then
        
           j = Right$(inPhrase, Len(inPhrase) - i + 1) ' j + jj
           Exit For
      End If
      
   Next i
   StripLetters = j
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "frmBrowser", "StripLetters", Err.Description
    Resume Exit_Proc


End Function



Private Sub WB3_DocumentComplete(ByVal pDisp As Object, URL As Variant)


    On Error GoTo Err_Proc
  Call GetFoodData
  Label2.Visible = False
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmBrowser", "WB3_DocumentComplete", Err.Description
    Resume Exit_Proc


End Sub



Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
