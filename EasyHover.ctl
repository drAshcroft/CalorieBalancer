VERSION 5.00
Begin VB.UserControl EasyHover 
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   ScaleHeight     =   3240
   ScaleWidth      =   3645
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1200
      Top             =   1800
   End
End
Attribute VB_Name = "EasyHover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
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
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  On Error Resume Next
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
 On Error Resume Next
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
On Error Resume Next
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
On Error Resume Next
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub Timer1_Timer()

On Error Resume Next
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
On Error Resume Next

    Clicked = True
    Call ChangePicture(m_Skin_Down)
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
On Error Resume Next
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As Long
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
On Error Resume Next
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
On Error Resume Next
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
On Error Resume Next
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
On Error Resume Next
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Release() As Variant
On Error Resume Next
  Clicked = False
  Call ChangePicture(m_Skin_Up)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Up() As Picture
On Error Resume Next
    Set Skin_Up = m_Skin_Up
End Property

Public Property Set Skin_Up(ByVal New_Skin_Up As Picture)
On Error Resume Next
    Set m_Skin_Up = New_Skin_Up
    Call ChangePicture(m_Skin_Up)
    PropertyChanged "Skin_Up"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Hover() As Picture
On Error Resume Next
    Set Skin_Hover = m_Skin_Hover
End Property

Public Property Set Skin_Hover(ByVal New_Skin_Hover As Picture)
On Error Resume Next
    Set m_Skin_Hover = New_Skin_Hover
    
    PropertyChanged "Skin_Hover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Skin_Down() As Picture
On Error Resume Next
    Set Skin_Down = m_Skin_Down
End Property

Public Property Set Skin_Down(ByVal New_Skin_Down As Picture)
On Error Resume Next
    Set m_Skin_Down = New_Skin_Down
    PropertyChanged "Skin_Down"
End Property

Private Sub UserControl_ExitFocus()
On Error Resume Next
Call Timer1_Timer
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    Set m_Skin_Up = LoadPicture("")
    Set m_Skin_Hover = LoadPicture("")
    Set m_Skin_Down = LoadPicture("")
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Timer1.Enabled = False
Timer1.Enabled = True
If Not Clicked Then Call ChangePicture(m_Skin_Hover)

End Sub
Private Sub ChangePicture(Pic As Picture)
On Error Resume Next
   If Not Pic Is Nothing Then
      UserControl.Width = ScaleX(Pic.Width, vbHimetric, vbTwips)
      UserControl.Height = ScaleY(Pic.Height, vbHimetric, vbTwips)
      Set UserControl.Picture = Pic
   End If
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
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
On Error Resume Next
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("Skin_Up", m_Skin_Up, Nothing)
    Call PropBag.WriteProperty("Skin_Hover", m_Skin_Hover, Nothing)
    Call PropBag.WriteProperty("Skin_Down", m_Skin_Down, Nothing)
End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
