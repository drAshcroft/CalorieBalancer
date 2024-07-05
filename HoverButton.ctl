VERSION 5.00
Begin VB.UserControl HoverButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   DefaultCancel   =   -1  'True
   ForeColor       =   &H00808000&
   MaskColor       =   &H00FF00FF&
   Picture         =   "HoverButton.ctx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   2430
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   480
   End
End
Attribute VB_Name = "HoverButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type



Dim Clicked As Boolean '0-is unclicked;1-clicked


Private Declare Function WindowFromPoint Lib "USER32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long

'Property Variables:
Dim m_Skin_Up As StdPicture
Dim m_Skin_Hover As StdPicture
Dim m_Skin_Down As StdPicture
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub Timer1_Timer()


       Dim PT As POINTAPI
       
       'where the mouse is now
       GetCursorPos PT
       
       'is this control under the mouse?
       If WindowFromPoint(PT.X, PT.Y) <> UserControl.hWnd Then
          Timer1.Enabled = False
          If Clicked Then
             Call ChangePicture(m_Skin_Down)
          Else
             Call ChangePicture(m_Skin_Up)
          End If
       
       End If


End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As Long
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Release() As Variant
  Clicked = False
  Call ChangePicture(m_Skin_Up)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Up() As Picture
    Set Skin_Up = m_Skin_Up
End Property

Public Property Set Skin_Up(ByVal New_Skin_Up As Picture)
    Set m_Skin_Up = New_Skin_Up
    Call ChangePicture(m_Skin_Up)
    PropertyChanged "Skin_Up"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Hover() As Picture
    Set Skin_Hover = m_Skin_Hover
End Property

Public Property Set Skin_Hover(ByVal New_Skin_Hover As Picture)
    Set m_Skin_Hover = New_Skin_Hover
    
    PropertyChanged "Skin_Hover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Down() As Picture
    Set Skin_Down = m_Skin_Down
End Property

Public Property Set Skin_Down(ByVal New_Skin_Down As Picture)
    Set m_Skin_Down = New_Skin_Down
    PropertyChanged "Skin_Down"
End Property

Private Sub UserControl_ExitFocus()
Call Timer1_Timer
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Skin_Up = LoadPicture("")
    Set m_Skin_Hover = LoadPicture("")
    Set m_Skin_Down = LoadPicture("")
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer1.Enabled = True
If Not Clicked Then Call ChangePicture(m_Skin_Hover)

End Sub
Private Sub ChangePicture(Pic As Picture)
   UserControl.Width = ScaleX(Pic.Width, vbHimetric, vbTwips)
   UserControl.Height = ScaleY(Pic.Height, vbHimetric, vbTwips)
   Set UserControl.Picture = Pic
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set m_Skin_Up = PropBag.ReadProperty("Skin_Up", Nothing)
    Set m_Skin_Hover = PropBag.ReadProperty("Skin_Hover", Nothing)
    Set m_Skin_Down = PropBag.ReadProperty("Skin_Down", Nothing)
    Call ChangePicture(m_Skin_Up)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("Skin_Up", m_Skin_Up, Nothing)
    Call PropBag.WriteProperty("Skin_Hover", m_Skin_Hover, Nothing)
    Call PropBag.WriteProperty("Skin_Down", m_Skin_Down, Nothing)
End Sub


