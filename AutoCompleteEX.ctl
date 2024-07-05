VERSION 5.00
Begin VB.UserControl AutoCompleteEX 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   ScaleHeight     =   3300
   ScaleWidth      =   3090
   Begin VB.ListBox LSuggest 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Width           =   3120
   End
   Begin VB.TextBox TData 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3105
   End
End
Attribute VB_Name = "AutoCompleteEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
'Event Declarations:
Event ExitFocus() 'MappingInfo=UserControl,UserControl,-1,ExitFocus
Event ItemSelected(SelectedID As Long)
Event NoneSelected()
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TData,TData,-1,KeyDown
Event TabPress(Shift As Boolean)
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TData,TData,-1,KeyUp
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event change() 'MappingInfo=TData,TData,-1,Change
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
'Default Property Values:
Const m_def_SelectedID = 0
Const m_def_SuggestedHeight = 0
'Const m_def_SuggestedHeight = 0
Const m_def_SearchTable = "abbrevExercise"
Const m_def_SearchField = "Exercisename"
'Property Variables:
Dim ShiftPressed As Boolean
Dim m_SelectedID As Long
Dim m_SuggestedHeight As Single
'Dim m_SuggestedHeight As Single
Dim m_SearchTable As String
Dim m_SearchField As String

Dim sDB As Database


Dim firstID As Long  'holds the ID for the first item on the list.  If they type out the whole phrase this
            'will hold the selected item or a -1111
'Modes

Dim WhoGotFocus As Integer ' 0-textbox  1 - listbox

Dim SelectedEvent As Integer '0-Nothing has been selected respond to everything
                             '1-Selected, stop giving choices
Dim M_ListBox As Object

Dim LosingFocus As Boolean

Public RecipeOnly As Boolean

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function GetNutrients() As Collection
On Error GoTo errhandl
Dim sTemp As Recordset
Dim temp As Collection
Dim junk As String, JunkName As String
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & m_SearchTable & " WHERE (index = " & _
                                  m_SelectedID & ");", dbOpenDynaset)
     If Err = 3075 Then
    'Here we got a bug!!
        Exit Function
        '
     End If
     Set temp = New Collection
     If Not sTemp.RecordCount = 0 Then
          sTemp.MoveFirst
          Dim i As Long
          On Error Resume Next
          For i = 0 To sTemp.Fields.Count - 1
            JunkName = sTemp.Fields(i).Name
            junk = sTemp.Fields(i)
            If Err.Number <> 0 Then
              junk = ""
              Err.Clear
            End If
            temp.Add junk, JunkName
          Next i
     End If
     Set GetNutrients = temp
     
    sTemp.Close
    Set sTemp = Nothing
errhandl:
End Function
Public Sub SetListBox(tListBox As Object)
On Error Resume Next
  Set M_ListBox = tListBox
End Sub

Private Function AutoComplete(SText As String) As Boolean
On Error GoTo errhandl
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
Dim cc As Long
Dim FirstHold As String
Dim SLen As Long
Dim Query As String
Dim Parts() As String, Req As String
Dim i As Long, j As Long
'Set AutoComplete function to FALSE
firstID = -1111
AutoComplete = False
If (Not SText = "") And Len(SText) > 1 Then
'Set OldLen as the sTextbox lenght
   OldLen = Len(SText)
   SText = Replace(SText, "'", "''")
   SText = Replace(SText, ",", "")
   Parts = Split(SText, " ")
   If UBound(Parts) = 0 Then
      Req = "(" & m_SearchField & " LIKE '*" & SText & "*')"
   Else
      Req = ""
      For i = 0 To UBound(Parts) - 1
         Req = Req & "(" & m_SearchField & " LIKE '*" & Parts(i) & "*') and "
      Next i
      Req = Req & "(" & m_SearchField & " LIKE '*" & Parts(i) & "*')"
   End If
      
   If RecipeOnly Then
      Req = Req & " and (ndb_no = '-100')"
   End If
   
   Query = "SELECT * FROM " & m_SearchTable & " WHERE ((" & Req & ") and ([formula]<>''))" & _
           " ORDER BY " & m_SearchTable & ".Usage DESC," & _
           m_SearchField & ";"
           
   SLen = Len(Query)
 
   
   Set sTemp = sDB.OpenRecordset(Query, dbOpenDynaset)
   If Err = 3075 Then
    'Here we got a bug!!
     '
   End If

   If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            LSuggest.Clear
            cc = 15
            UserControl.Height = LSuggest.Top + (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(4, vbPixels, UserControl.ScaleMode)
            LSuggest.Height = (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
            Exit Function
        Else
            sTemp.MoveFirst
            LSuggest.Clear
            cc = 0
          '  LSuggest.AddItem "Big Search"
          '  LSuggest.ItemData(cc) = -999
          '  cc = cc + 1
            
            Do While Not sTemp.EOF
                If (sTemp("formula") = "0" Or Trim$(sTemp("formula") = "")) Then sTemp.MoveNext
                If cc = 1 Then
                   FirstHold = sTemp.Fields(m_SearchField).Value
                   firstID = sTemp.Fields("index").Value
                ElseIf cc = 2 Then
                   LSuggest.AddItem FirstHold
                   LSuggest.ItemData(1) = firstID
                   LSuggest.AddItem sTemp.Fields(m_SearchField).Value
                   LSuggest.ItemData(2) = sTemp.Fields("index").Value
                Else
                   LSuggest.AddItem sTemp.Fields(m_SearchField).Value
                   LSuggest.ItemData(cc) = sTemp.Fields("index").Value + 0
                End If
                sTemp.MoveNext
                cc = cc + 1
            Loop
            
            If cc = 2 Then
              If Len(SText) <= Len(FirstHold) Then
                LSuggest.AddItem FirstHold
                LSuggest.ItemData(1) = firstID
              Else
                cc = 0
              End If
            End If
            
            
            If cc > 10 Then cc = 10
            
            If cc = 0 Then
'              UserControl.Height = LSuggest.Top
            Else
 '             UserControl.Height = LSuggest.Top + (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(4, vbPixels, UserControl.ScaleMode)
            End If
             cc = 15
            UserControl.Height = LSuggest.Top + (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(4, vbPixels, UserControl.ScaleMode)
            LSuggest.Height = (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
        End If
         AutoComplete = True
       Else
        m_SelectedID = -1111
        UserControl.Height = LSuggest.Top
        LSuggest.Clear
        'LSuggest.AddItem "Big Search"
        'LSuggest.ItemData(0) = -999
            cc = 15
            UserControl.Height = LSuggest.Top + (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(4, vbPixels, UserControl.ScaleMode)
            LSuggest.Height = (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
        If Not M_ListBox Is Nothing Then Call M_ListBox.LoadListBox(m_SelectedID)
    End If
Else
    m_SelectedID = -1111
            cc = 15
            UserControl.Height = LSuggest.Top + (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(4, vbPixels, UserControl.ScaleMode)
            LSuggest.Height = (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
    LSuggest.Clear
    'LSuggest.AddItem "Big Search"
    'LSuggest.ItemData(cc) = -999
    If Not M_ListBox Is Nothing Then Call M_ListBox.LoadListBox(m_SelectedID)
End If

Set sTemp = Nothing
Exit Function
errhandl:
  'MsgBox Err.Description, vbOKOnly, ""
  'Resume
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
                   
On Error Resume Next
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
On Error Resume Next
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Private Sub UpLoadtoTextbox()
On Error Resume Next
  Dim i As Long
  
  For i = 0 To LSuggest.ListCount - 1
   If LSuggest.Selected(i) Then
     SelectedEvent = 1
     If LCase$(LSuggest.List(i)) = "big search" Then
       If LCase$(TData.Text) <> "big search" Then
        Text = TData.Text
       Else
        Text = ""
       End If
     Else
        TData.Text = LSuggest.List(i)
     End If
     m_SelectedID = LSuggest.ItemData(i)
     firstID = m_SelectedID
     Call RegisterSelected
     Exit For
   End If
  Next i

End Sub

Private Sub LSuggest_GotFocus()


    On Error GoTo Err_Proc
LosingFocus = False

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "AutoComplete", "LSuggest_GotFocus", Err.Description
    Resume Exit_Proc


End Sub


Private Sub LSuggest_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Long
If WhoGotFocus = 1 And KeyAscii = 13 Or KeyAscii = 9 Then
    For i = 0 To LSuggest.ListCount - 1
      If LSuggest.Selected(i) Then
        m_SelectedID = LSuggest.ItemData(i)
      End If
    Next i
    If m_SelectedID = -1111 Or m_SelectedID = 0 Then
       RaiseEvent NoneSelected
       KeyAscii = 0
       Exit Sub
    Else
       Call UpLoadtoTextbox
    End If
    KeyAscii = 0
End If

End Sub

Private Sub LSuggest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    Call UpLoadtoTextbox
    If m_SelectedID = -1111 Or m_SelectedID = 0 Then
       RaiseEvent NoneSelected
       Exit Sub
    End If
End Sub


Private Sub TData_GotFocus()
On Error Resume Next
LosingFocus = False

End Sub

Private Sub TData_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 16 Then
  ShiftPressed = False
End If
If KeyCode = 13 Or KeyCode = 9 Then KeyCode = 0
End Sub



Private Sub UserControl_Click()


    On Error GoTo Err_Proc

    RaiseEvent Click
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "AutoComplete", "UserControl_Click", Err.Description
    Resume Exit_Proc


End Sub



Private Sub TData_KeyPress(KeyAscii As Integer)
On Error Resume Next
   If KeyAscii = 9 Or KeyAscii = 13 Then
    
    m_SelectedID = firstID
    If firstID = -1111 Or firstID = 0 Or LCase$(Trim$(TData.Text)) <> LCase$(Trim$(LSuggest.List(0))) Then
       m_SelectedID = 0
       RaiseEvent NoneSelected
       KeyAscii = 0
       Exit Sub
    Else
       KeyAscii = 0
       Call RegisterSelected
    End If
   End If
   
End Sub

Private Sub TData_Keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 If KeyCode = 16 And Shift = 1 Then
  ShiftPressed = True
 End If
 If KeyCode = 40 Then
    If LSuggest.ListCount >= 1 Then
      WhoGotFocus = 1
      LSuggest.SetFocus
      LSuggest.Selected(0) = True
    End If
 End If
 If KeyCode = 13 Then
    'Call TData_KeyPress(13)
    KeyCode = 0
 End If
 RaiseEvent KeyUp(KeyCode, Shift)
 If SelectedEvent = 1 Then
   Call AutoComplete(TData.Text)
   RaiseEvent change
 End If
End Sub

Private Sub RegisterSelected()
On Error Resume Next
   LSuggest.Clear
   UserControl.Height = TData.Height
   If m_SelectedID <> -1111 And m_SelectedID <> 0 Then
      If Not M_ListBox Is Nothing Then Call M_ListBox.LoadListBox(m_SelectedID)
      RaiseEvent ItemSelected(m_SelectedID)
   End If
End Sub

Private Sub TData_Change()
  On Error Resume Next
  If SelectedEvent = 0 Then
    Call AutoComplete(TData.Text)
    RaiseEvent change
  End If
End Sub

Public Sub CloseWithText()
 On Error Resume Next
    Call CloseBox
    Call RegisterSelected
    
End Sub


Private Sub UserControl_Resize()
On Error Resume Next
    TData.Left = 0
    TData.Top = 0
    TData.Width = UserControl.Width
  
    LSuggest.Left = 0
    LSuggest.Top = TData.Height + UserControl.ScaleX(3, vbPixels, UserControl.ScaleMode)
    LSuggest.Width = UserControl.Width
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TData,TData,-1,Text
Public Property Get Text() As String
On Error Resume Next
    Text = TData.Text
End Property

Public Property Let Text(ByVal New_Text As String)
On Error Resume Next
    SelectedEvent = 0
    If TData.Text = New_Text And New_Text <> "" Then
        Call AutoComplete(TData.Text)
        RaiseEvent change
    Else
        TData.Text = New_Text
    End If
    TData.SelStart = Len(New_Text)
    m_SelectedID = firstID
    If Not M_ListBox Is Nothing Then Call M_ListBox.LoadListBox(m_SelectedID)
    PropertyChanged "Text"
End Property

Public Sub SetTextAndID(Text As String, ID As Long)
On Error Resume Next
   SelectedEvent = 1
   TData.Text = Text
   TData.SelStart = Len(Text)
   m_SelectedID = ID
   firstID = m_SelectedID
   Call AutoComplete(TData.Text)
   If Not M_ListBox Is Nothing And Text <> "" Then Call M_ListBox.LoadListBox(m_SelectedID)
   SelectedEvent = 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    Set UserControl.Font = Ambient.Font
    m_SearchTable = m_def_SearchTable
    m_SearchField = m_def_SearchField
'    m_SuggestedHeight = m_def_SuggestedHeight
    m_SuggestedHeight = m_def_SuggestedHeight
    m_SelectedID = m_def_SelectedID
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Dim Index As Integer

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)

    TData.Text = PropBag.ReadProperty("Text", "")
    m_SearchTable = PropBag.ReadProperty("SearchTable", m_def_SearchTable)
    m_SearchField = PropBag.ReadProperty("SearchField", m_def_SearchField)
'    m_SuggestedHeight = PropBag.ReadProperty("SuggestedHeight", m_def_SuggestedHeight)
    m_SuggestedHeight = PropBag.ReadProperty("SuggestedHeight", m_def_SuggestedHeight)
    m_SelectedID = PropBag.ReadProperty("SelectedID", m_def_SelectedID)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next

Set sDB = Nothing
Set M_ListBox = Nothing

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
Dim Index As Integer

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", TData.Text, "")
    Call PropBag.WriteProperty("SearchTable", m_SearchTable, m_def_SearchTable)
    Call PropBag.WriteProperty("SearchField", m_SearchField, m_def_SearchField)
    Call PropBag.WriteProperty("SuggestedHeight", m_SuggestedHeight, m_def_SuggestedHeight)
    Call PropBag.WriteProperty("SelectedID", m_SelectedID, m_def_SelectedID)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function AddDataBase(tDB As Database) As Boolean
  On Error GoTo errhandl
  Set sDB = tDB
  AddDataBase = True
  
  Exit Function
errhandl:
  AddDataBase = False
  
  
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SearchTable() As String
On Error Resume Next
    SearchTable = m_SearchTable
End Property

Public Property Let SearchTable(ByVal New_SearchTable As String)
On Error Resume Next
    m_SearchTable = New_SearchTable
    PropertyChanged "SearchTable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SearchField() As String
On Error Resume Next
    SearchField = m_SearchField
End Property

Public Property Let SearchField(ByVal New_SearchField As String)
On Error Resume Next
    m_SearchField = New_SearchField
    PropertyChanged "SearchField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,1,1,0
Public Property Get SuggestedHeight() As Single
On Error Resume Next
    SuggestedHeight = TData.Height
End Property

Public Property Let SuggestedHeight(ByVal New_SuggestedHeight As Single)
On Error Resume Next
    m_SuggestedHeight = New_SuggestedHeight
    TData.Height = m_SuggestedHeight
    UserControl.Height = TData.Height
    PropertyChanged "SuggestedHeight"
End Property

Public Function Translate(ID As Long) As String
On Error Resume Next
  Dim temp As Recordset, first As Boolean
  Set temp = sDB.OpenRecordset("Select " & m_SearchField & " from " & m_SearchTable & _
          " where index = " & ID & ";", dbOpenDynaset)
If temp Is Nothing Then Exit Function
  While (Not temp.EOF) And (Not first)
     Translate = temp.Fields(m_SearchField)
     temp.MoveNext
     first = True
  Wend
  
  temp.Close
  Set temp = Nothing

End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get SelectedID() As Long
On Error Resume Next

    SelectedID = m_SelectedID
End Property

Public Property Let SelectedID(ByVal New_SelectedID As Long)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 387
  
    m_SelectedID = New_SelectedID
    firstID = m_SelectedID
    If m_SelectedID <> 0 Then
      Dim temp As String
      temp = Translate(m_SelectedID)
      SelectedEvent = 1
      TData.Text = temp
      SelectedEvent = 0
      'need to look up text for id here
      If Not M_ListBox Is Nothing Then Call M_ListBox.LoadListBox(m_SelectedID)
    Else
      TData.Text = ""
    End If
    
    PropertyChanged "SelectedID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function NotifyOfFocus() As Variant
On Error Resume Next
    m_SelectedID = firstID
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Reset() As Variant
On Error Resume Next
   SelectedEvent = 0
   
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function PutAtEnd() As Variant
On Error Resume Next
TData.SelStart = Len(TData.Text)
TData.SelLength = 0

End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function CloseBox() As Variant
On Error Resume Next
  UserControl.Height = TData.Height
End Function

Private Sub UserControl_ExitFocus()
On Error Resume Next
  Dim i As Long
    For i = 0 To LSuggest.ListCount - 1
      If LSuggest.Selected(i) Then
        m_SelectedID = LSuggest.ItemData(i)
      End If
    Next i
    If m_SelectedID = -1111 Or m_SelectedID = 0 Then
       RaiseEvent NoneSelected
       Exit Sub
    Else
       Call UpLoadtoTextbox
    End If
    
    RaiseEvent ExitFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
On Error Resume Next
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
On Error Resume Next
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

       Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function


