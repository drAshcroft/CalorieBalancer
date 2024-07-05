VERSION 5.00
Begin VB.UserControl UnitListBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "UnitListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Dim ListBoxConversions() As Single
Dim DB As Database


'Default Property Values:
Const m_def_SearchTable = "Weight"
Const m_def_SearchField = "index"
Const m_def_ItemConversion = 0
Const m_def_GramsSelected = 0
Const m_def_Text = ""
'Property Variables:
Dim m_SearchTable As String
Dim m_SearchField As String

Dim m_GramsSelected As Single
Dim m_Text As String
'Event Declarations:
Event ItemSelected(Item As String, Conversion As Single)
Dim ShiftPressed As Boolean
Event TabPress(ShiftPressed As Boolean)


Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function Translate(FoodID As Long, Unit As String) As Single


    On Error GoTo Err_Proc
    Translate = Module1.TranslateUnitToGrams(FoodID, Unit)
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "UnitListBox", "Translate", Err.Description
    Resume Exit_Proc


End Function


Public Sub LoadListBox(FoodID As Long)
On Error GoTo errhandl
   Dim temp As Recordset
   Dim cc As Long
   Dim BiggestWidth As Single
   Dim tempSize As Single
   Dim junkString As String
   List1.Clear
   m_Text = ""
   m_GramsSelected = 0
   
   If FoodID <> -1111 And FoodID <> 0 And FoodID > -200 Then
      Set temp = DB.OpenRecordset("SELECT *" _
                              & " From " & m_SearchTable _
                              & " WHERE (((" & m_SearchField & ")=" & FoodID & "));", dbOpenDynaset)
      If Not temp.RecordCount = 0 Then
        If temp.EOF = True And temp.BOF = True Then
            Exit Sub
        Else
            temp.MoveFirst
       
            cc = 0
            BiggestWidth = 0
            ReDim ListBoxConversions(cc)
            Dim alreadyGrams As Boolean
            Do While Not temp.EOF
                junkString = temp.Fields("msre_desc").Value
                'If LCase$(junkString) = "grams" Then alreadyGrams = True
                List1.AddItem junkString
                ListBoxConversions(cc) = temp.Fields("gm_wgt").Value / temp.Fields("amount").Value
                tempSize = UserControl.TextWidth(junkString)
                If tempSize > BiggestWidth Then BiggestWidth = tempSize
                temp.MoveNext
                cc = cc + 1
                ReDim Preserve ListBoxConversions(cc)
            Loop
            Call Module1.ConvertUnits(List1, ListBoxConversions, cc)
            If (Not alreadyGrams) And ListBoxConversions(0) <> 100 Then
                List1.AddItem "OZ."
                ListBoxConversions(cc) = 28.3495231
                cc = cc + 1
                ReDim Preserve ListBoxConversions(cc)
                List1.AddItem "Grams"
                ListBoxConversions(cc) = 1
                cc = cc + 1
                ReDim Preserve ListBoxConversions(cc)
            End If
            
            If cc > 10 Then cc = 10
            UserControl.Height = (cc + 1) * UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
        End If
        
       End If
       m_Text = List1.List(0)
       m_GramsSelected = ListBoxConversions(0)
   Else
       List1.Clear
       UserControl.Height = UserControl.TextHeight("^~$_)") + ScaleX(2, vbPixels, UserControl.ScaleMode)
   End If
   
'   temp.Close
   Set temp = Nothing
    'boxwidth = BiggestWidth + UserControl.ScaleX(15, vbPixels, UserControl.ScaleMode)
    'UserControl.Width = boxwidth
    Exit Sub
errhandl:
    MsgBox "Unable to find unit." & vbCrLf & Err.Description, vbOKOnly, ""
    If DoDebug Then Resume
End Sub
Private Sub SelectText()
On Error GoTo errhandl
    Dim i As Long
    Dim junk As String
    If List1.ListCount = 0 Then
      m_Text = ""
      m_GramsSelected = 0
    Else
      For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          junk = List1.List(i)
          Exit For
        End If
       Next i
       m_GramsSelected = ListBoxConversions(i)
       m_Text = junk
    End If
    
Exit Sub
errhandl:
MsgBox Err.Description, vbOKOnly, ""
End Sub
Private Sub CloseList1()
    On Error Resume Next
    If Left$(m_Text, 2) <> "--" Then RaiseEvent ItemSelected(m_Text, m_GramsSelected)
End Sub


Private Sub List1_Click()
On Error Resume Next
   Call SelectText
End Sub

Private Sub List1_GotFocus()
Dim Selected As Boolean, i As Long
On Error Resume Next
Selected = False
For i = 0 To List1.ListCount - 1
  Selected = List1.Selected(i) Or Selected
Next i
If Not Selected Then
  List1.Selected(0) = True
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = 9 Then
    Call SelectText
    
    If Left$(m_Text, 2) <> "--" Then RaiseEvent TabPress(ShiftPressed)
    KeyAscii = 0
  End If
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 16 Then
  ShiftPressed = False
End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Call SelectText
    Call CloseList1
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 16 And Shift = 1 Then
    ShiftPressed = True
    Exit Sub
  End If
  Call SelectText
  If KeyCode = 13 Then
     Call CloseList1
     KeyCode = 0
  End If
End Sub

Public Sub UserControl_ExitFocus()
On Error Resume Next
    Call SelectText
    Call CloseList1
   
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  Call SelectText
  If KeyCode = 13 Then
     Call CloseList1
     KeyCode = 0
  End If

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,List
Public Property Get List(ByVal Index As Integer) As String
On Error Resume Next
    List = List1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
On Error Resume Next
    List1.List(Index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,ListCount
Public Property Get ListCount() As Integer
On Error Resume Next
    ListCount = List1.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,ListIndex
Public Property Get ListIndex() As Integer
On Error Resume Next
    ListIndex = List1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
On Error Resume Next
    List1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Selected
Public Property Get Selected(ByVal Index As Integer) As Boolean
On Error Resume Next
    Selected = List1.Selected(Index)
End Property

Public Property Let Selected(ByVal Index As Integer, ByVal New_Selected As Boolean)
On Error Resume Next
    List1.Selected(Index) = New_Selected
    PropertyChanged "Selected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Clear
Public Sub Clear()
On Error Resume Next
    List1.Clear
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddDataBase(sDB As Database) As Variant
On Error Resume Next
   Set DB = sDB
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,1,2,0
Public Property Get GramsSelected() As Single
On Error Resume Next
    GramsSelected = m_GramsSelected
End Property

Public Property Let GramsSelected(ByVal New_GramsSelected As Single)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_GramsSelected = New_GramsSelected
    PropertyChanged "GramsSelected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get Text() As String
On Error Resume Next
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
On Error Resume Next
    Dim i As Long
    If Ambient.UserMode = False Then Err.Raise 387
    m_Text = New_Text
    For i = 0 To List1.ListCount - 1
       If List1.List(i) = m_Text Then
         List1.Selected(i) = True
       End If
    Next i
    Call SelectText
    PropertyChanged "Text"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    m_GramsSelected = m_def_GramsSelected
    m_Text = m_def_Text
    
    m_SearchTable = m_def_SearchTable
    m_SearchField = m_def_SearchField
End Sub



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Dim Index As Integer
    m_GramsSelected = PropBag.ReadProperty("GramsSelected", m_def_GramsSelected)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    
    m_SearchTable = PropBag.ReadProperty("SearchTable", m_def_SearchTable)
    m_SearchField = PropBag.ReadProperty("SearchField", m_def_SearchField)
    Set List1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = List1.Font
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
  List1.Top = 0
  List1.Left = 0
  List1.Width = UserControl.Width
  List1.Height = UserControl.Height
  DoEvents
  UserControl.Height = List1.Height
  
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next

Set DB = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Dim Index As Integer
Call PropBag.WriteProperty("GramsSelected", m_GramsSelected, m_def_GramsSelected)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    
    Call PropBag.WriteProperty("SearchTable", m_SearchTable, m_def_SearchTable)
    Call PropBag.WriteProperty("SearchField", m_SearchField, m_def_SearchField)
    Call PropBag.WriteProperty("Font", List1.Font, Ambient.Font)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,2,0
Public Property Get ItemConversion(ByVal Index As Long) As Single
On Error Resume Next
    ItemConversion = ListBoxConversions(Index)
End Property

Public Property Let ItemConversion(ByVal Index As Long, ByVal New_ItemConversion As Single)
On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 387
    ListBoxConversions(Index) = New_ItemConversion
    PropertyChanged "ItemConversion"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
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
'MemberInfo=13,0,0,0
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
'MemberInfo=14
Public Function NotifyOfFocus() As Variant
On Error Resume Next
    Call CloseList1
End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = List1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set List1.Font = New_Font
    Set UserControl.Font = List1.Font
    PropertyChanged "Font"
End Property


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

       Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
