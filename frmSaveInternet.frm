VERSION 5.00
Begin VB.Form frmSaveInternet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save New Food to Internet"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      Begin VB.VScrollBar VScroll1 
         Height          =   6975
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin CalorieBalance.RDADisplay RDADisplay1 
         Height          =   13620
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   24024
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Never Save      (do not ask)    "
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do not save to internet"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save to Internet Database"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
End
Attribute VB_Name = "frmSaveInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Public AbbrevID As Long
Public SaveFood As Boolean
Public HostForm As FNewFood
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage
Private Template As String

Private Sub Command1_Click()
   OpenURL App.path & "\resources\temp\newfood.html"
    Unload Me
    Me.Hide
End Sub

Private Sub Command2_Click()


    On Error GoTo Err_Proc
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSaveIntern", "Command2_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Command3_Click()


    On Error GoTo Err_Proc
If SaveFood Then
   Call SaveSetting(App.Title, "settings", "SaveIntenetfood", 2)
Else
   Call SaveSetting(App.Title, "settings", "SaveIntenetRecipe", 2)
   
End If
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSaveIntern", "Command3_Click", Err.Description
    Resume Exit_Proc


End Sub

Private Sub Form_Load()


    On Error GoTo Err_Proc
Unload Me
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "frmSaveIntern", "Form_Load", Err.Description
    Resume Exit_Proc


End Sub

Private Sub SaveInternetFood(AbbrevID As Long)
'On Error GoTo errhandl

   
    


Dim temp2 As Recordset
Dim temp As Recordset

Dim ideals As Recordset


Set temp = DB.OpenRecordset("select * from abbrev where index=" & AbbrevID & ";", dbOpenDynaset)

Set temp2 = DB.OpenRecordset("SELECT *" _
                              & " From Weight " _
                              & " WHERE (((index)=" & AbbrevID & "));", dbOpenDynaset)
Set ideals = DB.OpenRecordset("Select * from ideals where user='AnyUser';", dbOpenDynaset)

Call RDADisplay1.DisplayFood(temp, ideals, temp2, temp2.Fields("Amount").Value, temp2.Fields("msre_desc").Value, temp2("gm_wgt"))

    Dim ff As Long
    Dim Template As String
    Dim junk As String
    ff = FreeFile
    Open App.path & "\resources\daily\newfood.html" For Input As #ff
    While Not EOF(ff)
       Line Input #ff, junk
       Template = Template + junk + vbCrLf
    Wend
    Close #ff
    
    Template = Replace(Template, "~foodname~", temp("foodname"), , , vbTextCompare)
    
    junk = "value=""" + Right$("0000" & temp("foodgroup"), 4) + """"
    Template = Replace(Template, junk, junk & " selected ", , , vbTextCompare)
    
    Template = Replace(Template, "~amount~", ConvertDecimalToFraction(temp2.Fields("Amount").Value), , , vbTextCompare)
    Template = Replace(Template, "~unit~", temp2.Fields("msre_desc").Value, , , vbTextCompare)
    Template = Replace(Template, "~grams~", temp2("gm_wgt"), , , vbTextCompare)
    
    Dim BC As Double, i As Long
    If Not (temp2.EOF And temp2.BOF) Then
      
        BC = temp2.Fields("gm_wgt").Value / 100
        Set temp2 = Nothing
    Else
       BC = 1
    End If
    
    'some of the nutrients are listed by raw values and others by percents.  The template will have a p before
    'the ones that need percents.  just loop through all the information and try it and then do the same with
    'the percents

    Dim junk2 As String
 
    Switch = False
    On Error Resume Next
    For i = 0 To temp.Fields.Count - 1
        junk2 = 0
        junk = MakeWebName(temp.Fields(i).Name)
          
        If VarType(temp(i)) <> vbString Then
           junk2 = Round(temp(i) * BC, 2)
        Else
           junk2 = temp(i)
        End If
        junk = "name=""" + junk + """"
        Template = Replace(Template, junk, junk & " value= " + junk2 + " ", , , vbTextCompare)
         
    Next
    For i = 0 To temp.Fields.Count - 1
        junk2 = 0
        junk = MakeWebName(temp.Fields(i).Name)
          
        If VarType(temp(i)) <> vbString Then
           junk2 = Round(temp(i) * BC / ideals(temp.Fields(i).Name) * 100, 1)
        Else
           junk2 = temp(i)
        End If
        junk = "name=""p" + junk + """"
        Template = Replace(Template, junk, junk & " value= " + junk2 + " ", , , vbTextCompare)
         
    Next
    
    ff = FreeFile
    Open App.path & "\resources\temp\newfood.html" For Output As #ff
    Print #ff, , Template
    Close #ff
ideals.Close
errhandl:
Set ideals = Nothing
Set temp = Nothing


End Sub

Private Function MakeWebName(ListItem As String) As String
  On Error Resume Next
  Dim junk As String
  junk = Trim$(ListItem)
  junk = Replace(junk, "-", "__")
  junk = Replace(junk, " ", "_")
 
  If junk = "Fiber" Then junk = "Dietary fiber"
  If LCase$(junk) = "Carbs" Then junk = "carbohydrate"
  MakeWebName = junk
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

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function

Private Sub Form_Resize()
 VScroll1.Move Picture2.Width - VScroll1.Width, 0, VScroll1.Width, Picture2.Height
End Sub

Private Sub VScroll1_Change()
RDADisplay1.Top = (Height - RDADisplay1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub

Private Sub VScroll1_Scroll()
RDADisplay1.Top = (Height - RDADisplay1.Height * 1.1) * (VScroll1.Value / VScroll1.Max)
End Sub
