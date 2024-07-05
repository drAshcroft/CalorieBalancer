VERSION 5.00
Begin VB.UserControl ExerciseBox 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4170
   ScaleWidth      =   4800
   Begin VB.CommandButton Command3 
      Caption         =   "Gray Out"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Black out "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "ExerciseBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
'Default Property Values:
Const m_def_Text = ""
'Property Variables:
Dim m_Text As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1(0),Text1,0,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1(0),Text1,0,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1(0),Text1,0,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1(0),Text1,0,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event change() 'MappingInfo=Text1(0),Text1,0,Change
Attribute change.VB_Description = "Occurs when the contents of a control have changed."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."


Dim Labels() As String

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1(0),Text1,0,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."

    ForeColor = Text1(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  Dim i As Long
  For i = 0 To Text1.UBound
    Text1(i).ForeColor() = New_ForeColor
  Next i
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1(0),Text1,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Dim i As Long
  For i = 0 To Text1.UBound

    Set Text1(i).Font = New_Font
  Next
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Command1_Click()
Call Text1_KeyUp(0, 13, 0)
End Sub

Private Sub Command2_Click()
Text1(0).Text = "-" & Text1(0).Text
Call Text1_KeyUp(0, 13, 0)
End Sub

Private Sub Command3_Click()
Text1(0).Text = "*" & Text1(0).Text
Call Text1_KeyUp(0, 13, 0)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Text1_Change(Index As Integer)
    RaiseEvent change
End Sub

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
On Error Resume Next
 Dim i As Long
 m_Text = Text1(0)
 For i = 1 To Text1.UBound
   m_Text = m_Text & "/" & Text1(i)
 Next i
 
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
On Error Resume Next
  Dim i As Long
  Dim Parts() As String
  If InStr(1, New_Text, "/") = 0 Then
    ReDim Parts(0)
    Parts(0) = New_Text
  Else
    Parts = Split(New_Text, "/")
  End If
  For i = 1 To UBound(Parts)
    Text1(i) = Parts(i)
  Next i
  
  
  Text1(0).Text = Parts(0)
    Text1(0).SelStart = Len(Parts(0))
    Text1(0).SelLength = 0
    m_Text = New_Text
    PropertyChanged "Text"
End Property

Public Sub SetLabels(FoodItem As String)
On Error Resume Next
  Dim junk As String
  
  
  junk = Trim$(StrReverse(FoodItem))
  Dim i As Long
  i = InStr(1, junk, "(")
  If i = 0 Then
    junk = "(min)"
    junk = StrReverse(junk)
    i = InStr(1, junk, "(")
  End If
  
  If i <> 0 Then
     Dim junk2 As String
     junk2 = Mid$(junk, 2, i - 2)
     junk2 = StrReverse(junk2)
     Dim Parts() As String
     
     junk2 = Replace(junk2, "min,", "Minutes,", , , vbTextCompare)
     junk2 = Replace(junk2, "min)", "Minutes)", , , vbTextCompare)
     If LCase$(junk2) = "min" Then junk2 = "Minutes"
     Labels = Split(junk2, "/")
     
     For i = 1 To Label1.UBound
        Unload Text1(i)
        Unload Label1(i)
     Next i
     
     For i = 1 To UBound(Labels)
        Load Text1(i)
        Load Label1(i)
    
        Text1(i).Top = i * Text1(0).Height * 1.2
        Label1(i).Top = i * Text1(0).Height * 1.2
    
        Text1(i).ZOrder
        Text1(i).TabIndex = i
        Label1(i).ZOrder
        Text1(i) = ""
        Label1(i).Caption = Labels(i)
        Label1(i).Visible = True
        Text1(i).Visible = True
      Next i
      Text1(0).Text = ""
      Label1(0).Caption = Labels(0) & ":"
      Text1(0).Top = 0
      Text1(0).TabIndex = 0
      Label1(0).Top = 0
      UserControl.Height = i * Text1(0).Height * 1.2 + Command1.Height + 50
      
      UserControl.Width = Text1(0).Left + Text1(0).Width + 100
     Dim tp As Single
     tp = i * Text1(0).Height * 1.2 + 50
     Command1.Left = Text1(0).Left + Text1(0).Width + 100 - Command1.Width
     Command1.Top = tp
     Command2.Top = tp
     Command3.Top = tp
     Command2.Left = 0
     Command3.Left = Command2.Width
     
  End If
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
      Dim i As Long
  For i = 0 To Text1.UBound
    Text1(i).ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set Text1(i).Font = PropBag.ReadProperty("Font", Ambient.Font)
  Next
        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Text1(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Text1(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
End Sub

