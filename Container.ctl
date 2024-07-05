VERSION 5.00
Begin VB.UserControl IoxContainer 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   1740
   ToolboxBitmap   =   "Container.ctx":0000
   Begin VB.Timer FocusTimer 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Scrollpic 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1500
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1500
      Width           =   240
   End
   Begin VB.HScrollBar HScroll 
      Height          =   240
      Left            =   0
      Max             =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1500
      Width           =   1500
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1500
      Left            =   1500
      Max             =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "IoxContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////  IOX Container  \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'This is the most advanced container control in PSC.
'Did you love controls that don't have much "client-side" code...

'Ok, this is a 0 line client-side code control, It´s a REAL container,
'You just put controls inside and the contros do the dirty job.
'Allows to put many controls in the panel, an use scrollbars to acces them
'This panel suport Mouse Wheel, WITHOUT SUBCLASS, just uses the free time
'to Peek Messages, so its IDE safe.
'It has a lot of features, like:
'Use ScrollBarConstants (vbBoth, vbVertical, vbSBNone, vbHorizontal)
'Use ScrollBar sensibility
'Vertical and horizontal margin to ajust the contained controls
'Change the potion of the scrolls when a contained controls got focus
'Is aligneable

'I got the idea from IsPanel, By DavidJ, but this is a complete rewrite, so
'it uses diferent and more eficient programing techniques.


' Created by Ivan Tellez
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\  IOX Container  //////////////////////////////



' API Declarations
' ==================================
Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function PeekMessage Lib "USER32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "USER32" () As Long



' API Constants
' ==================================
Private Const PM_REMOVE = &H1

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private bCancel As Boolean
Private Const WM_MOUSEWHEEL = 522


Private Const SM_CYVSCROLL = 20

Private Type CurrentControlType
        Name As String
        Index As Long
End Type

Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type




' Control properties
' ==================================
Public Enum BorderStyleEnum
    [None]
    [Fixed Single]
End Enum

Public Enum ScrollBehaviorEnum
    [Normal]
    [Middle]
    [Reverse]
End Enum

Public Enum SensibilityEnum
    [Highest]
    [High]
    [Medium]
    [Low]
End Enum

Public Enum MarginEnum
    [None] = 0
    [5 pixels] = 5
    [10 pixels] = 10
    [15 pixels] = 15
    [20 pixels] = 20
    [25 pixels] = 25
    [30 pixels] = 30
End Enum

'Default Property Values:
Const m_def_BorderStyle = 1
Const m_def_Enabled = True
Const m_def_ScrollBars = 3
Const m_def_Sensibility = 2
Const m_def_MarginV = 5
Const m_def_MarginH = 5
Const m_def_ScrollBehavior = 0

'Property Variables:
Dim m_BorderStyle As BorderStyleEnum
Dim m_Enabled As Boolean
Dim m_ScrollBars As ScrollBarConstants
Dim m_Sensibility As SensibilityEnum
Dim m_MarginV As MarginEnum
Dim m_MarginH As MarginEnum
Dim m_ScrollBehavior As ScrollBehaviorEnum

Private HPrevValue As Long
Private VPrevValue As Long
Private TempControl As Control

Private CurrCtrl As CurrentControlType
Private LastCtrl As CurrentControlType





'Initialize properties for a new user control
' ==================================
Private Sub UserControl_InitProperties()
    m_BorderStyle = m_def_BorderStyle
    m_Enabled = m_def_Enabled
    m_ScrollBars = m_def_ScrollBars
    m_Sensibility = m_def_Sensibility
    m_MarginH = m_def_MarginH
    m_MarginV = m_def_MarginV
    m_ScrollBehavior = m_def_ScrollBehavior
End Sub

'Load properties from the PropBag
' ==================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Sensibility = PropBag.ReadProperty("Sensibility", m_def_Sensibility)
    m_MarginH = PropBag.ReadProperty("MarginH", m_def_MarginH)
    m_MarginV = PropBag.ReadProperty("MarginV", m_def_MarginV)
    m_ScrollBehavior = PropBag.ReadProperty("ScrollBehavior", m_def_ScrollBehavior)
End Sub

'Save properties from the PropBag
' ==================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ScrollBars", m_ScrollBars, m_def_ScrollBars)
    Call PropBag.WriteProperty("Sensibility", m_Sensibility, m_def_Sensibility)
    Call PropBag.WriteProperty("MarginH", m_MarginH, m_def_MarginH)
    Call PropBag.WriteProperty("MarginV", m_MarginV, m_def_MarginV)
    Call PropBag.WriteProperty("ScrollBehavior", m_ScrollBehavior, m_def_ScrollBehavior)
End Sub





'User control properties
' ==================================
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Sensibility() As SensibilityEnum
    Sensibility = m_Sensibility
End Property
Public Property Let Sensibility(ByVal New_Sensibility As SensibilityEnum)
    If Ambient.UserMode Then Err.Raise 382
    m_Sensibility = New_Sensibility
    PropertyChanged "Sensibility"
End Property

Public Property Get ScrollBehavior() As ScrollBehaviorEnum
    ScrollBehavior = m_ScrollBehavior
End Property
Public Property Let ScrollBehavior(ByVal New_ScrollBehavior As ScrollBehaviorEnum)
    If Ambient.UserMode Then Err.Raise 382
    m_ScrollBehavior = New_ScrollBehavior
    PropertyChanged "ScrollBehavior"
End Property

Public Property Get MarginV() As MarginEnum
    MarginV = m_MarginV
End Property
Public Property Let MarginV(ByVal New_MarginV As MarginEnum)
    If Ambient.UserMode Then Err.Raise 382
    m_MarginV = New_MarginV
    PropertyChanged "MarginV"
    UserControl_Resize
End Property

Public Property Get MarginH() As MarginEnum
    MarginH = m_MarginH
End Property
Public Property Let MarginH(ByVal New_MarginH As MarginEnum)
    If Ambient.UserMode Then Err.Raise 382
    m_MarginH = New_MarginH
    PropertyChanged "MarginH"
    UserControl_Resize
End Property

Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ScrollBars() As ScrollBarConstants
    ScrollBars = m_ScrollBars
End Property
Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarConstants)
    m_ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
    UserControl_Resize
End Property



'User control methods
' ==================================
Private Sub UserControl_Terminate()
MsgBox "iox"
bCancel = True
FocusTimer = True
End Sub

Private Sub UserControl_Show()
UserControl_Resize
FocusTimer.Enabled = True
ProcessMessages
End Sub

Private Sub UserControl_Initialize()
    CurrCtrl.Name = ""
    CurrCtrl.Index = 0
    LastCtrl.Name = ""
    LastCtrl.Index = 0
    HPrevValue = 0
    VPrevValue = 0
    HScroll.Value = 0
    VScroll.Value = 0
End Sub

Private Sub UserControl_Paint()
If Not Ambient.UserMode Then UserControl_Resize
End Sub

















'Private methods
' ==================================
Private Sub UserControl_Resize()
On Error Resume Next
Dim ScrollWidth As Long
Dim ScrollHeight As Long

ScrollWidth = GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelX
ScrollHeight = GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelY
HScroll.Enabled = False
VScroll.Enabled = False

With UserControl
    Select Case m_ScrollBars        'Position the controls
        Case vbSBNone
            HScroll.Visible = False
            VScroll.Visible = False
            Scrollpic.Visible = False
        Case vbHorizontal
            HScroll.Visible = True
            VScroll.Visible = False
            Scrollpic.Visible = False
            HScroll.Move 0, .ScaleHeight - ScrollHeight, .ScaleWidth, ScrollHeight
            HScroll.Enabled = (CalcClientWidth > UserControl.ScaleWidth)
        Case vbVertical
            HScroll.Visible = False
            VScroll.Visible = True
            Scrollpic.Visible = False
            VScroll.Move .ScaleWidth - ScrollWidth, 0, ScrollWidth, .ScaleHeight
            VScroll.Enabled = (CalcClientHeight > UserControl.ScaleHeight)
        Case vbBoth
            HScroll.Visible = True
            VScroll.Visible = True
            Scrollpic.Visible = True
            VScroll.Move .ScaleWidth - ScrollWidth, 0, ScrollWidth, .ScaleHeight + (ScrollHeight * -1)
            HScroll.Move 0, .ScaleHeight - ScrollHeight, .ScaleWidth + (ScrollWidth * -1), ScrollHeight
            Scrollpic.Move .ScaleWidth - ScrollWidth, .ScaleHeight - ScrollHeight, ScrollWidth, ScrollHeight
            HScroll.Enabled = (CalcClientWidth > (UserControl.ScaleWidth - VScroll.Width))
            VScroll.Enabled = (CalcClientHeight > (UserControl.ScaleHeight - HScroll.Height))
    End Select
    
    If Ambient.UserMode Then
        CalcScrollValues
    End If
    
End With

HScroll.ZOrder
VScroll.ZOrder
Scrollpic.ZOrder
End Sub

Private Function CalcClientWidth() As Long
On Error Resume Next
Dim MaxWidth As Long
Dim ctrl As Object
For Each ctrl In UserControl.ContainedControls
    If MaxWidth < (ctrl.Left + ctrl.Width + (m_MarginH * Screen.TwipsPerPixelX)) Then
        MaxWidth = (ctrl.Left + ctrl.Width + (m_MarginH * Screen.TwipsPerPixelX))
    End If
Next
If MaxWidth <> 0 Then
    CalcClientWidth = MaxWidth    'Maximo valor hacia la derecha
Else
    CalcClientWidth = 0 'UserControl.ScaleWidth
End If
End Function

Private Function CalcClientHeight() As Long
On Error Resume Next
Dim MaxHeight As Long
Dim ctrl As Object
For Each ctrl In UserControl.ContainedControls
    If MaxHeight < (ctrl.Top + ctrl.Height + (m_MarginV * Screen.TwipsPerPixelX)) Then
        MaxHeight = (ctrl.Top + ctrl.Height + (m_MarginV * Screen.TwipsPerPixelX))
    End If
Next
If MaxHeight <> 0 Then
    CalcClientHeight = MaxHeight    'Max down val
Else
    CalcClientHeight = 0
End If
End Function


Private Sub CalcScrollValues()
Dim NewMaxVal As Long
Select Case m_ScrollBars 'Position the controls
    Case vbSBNone
    Case vbHorizontal    'Only  horizontal scroll
        With HScroll
            NewMaxVal = CalcClientWidth() - UserControl.ScaleWidth
            If NewMaxVal > 32767 Then NewMaxVal = 32767
            .Max = NewMaxVal
            If .Value > .Max Then .Value = .Max
            .LargeChange = .Max
            Select Case m_Sensibility
                Case [Highest]
                    .SmallChange = .LargeChange / 20
                Case [High]
                    .SmallChange = .LargeChange / 15
                Case [Medium]
                    .SmallChange = .LargeChange / 10
                Case [Low]
                    .SmallChange = .LargeChange / 5
            End Select

        End With
        VScroll.Max = 0
        VScroll.Value = 0
        
    Case vbVertical   'Only  Vertical scroll
        With VScroll
            Dim test As Variant
            NewMaxVal = CalcClientHeight - UserControl.ScaleHeight
            If NewMaxVal > 32767 Then NewMaxVal = 32767
            .Max = NewMaxVal
            If .Value > .Max Then .Value = .Max
            .LargeChange = .Max
            Select Case m_Sensibility
                Case [Highest]
                    .SmallChange = .LargeChange / 20
                Case [High]
                    .SmallChange = .LargeChange / 15
                Case [Medium]
                    .SmallChange = .LargeChange / 10
                Case [Low]
                    .SmallChange = .LargeChange / 5
            End Select
            If .Value > .Max Then .Value = .Max
        End With
        HScroll.Max = 0
        HScroll.Value = 0
        
    Case vbBoth     'Both scrolls
        With HScroll
            NewMaxVal = CalcClientWidth - (UserControl.ScaleWidth - VScroll.Width)
            If NewMaxVal > 32767 Then NewMaxVal = 32767
            .Max = NewMaxVal
            If .Value > .Max Then .Value = .Max
            .LargeChange = .Max / 2
            Select Case m_Sensibility
                Case [Highest]
                    .SmallChange = .LargeChange / 20
                Case [High]
                    .SmallChange = .LargeChange / 15
                Case [Medium]
                    .SmallChange = .LargeChange / 10
                Case [Low]
                    .SmallChange = .LargeChange / 5
            End Select
        End With
        With VScroll
            NewMaxVal = CalcClientHeight - (UserControl.ScaleHeight - HScroll.Height)
            If NewMaxVal > 32767 Then NewMaxVal = 32767
            .Max = NewMaxVal
            If .Value > .Max Then .Value = .Max
            .LargeChange = .Max
            Select Case m_Sensibility
                Case [Highest]
                    .SmallChange = .LargeChange / 20
                Case [High]
                    .SmallChange = .LargeChange / 15
                Case [Medium]
                    .SmallChange = .LargeChange / 10
                Case [Low]
                    .SmallChange = .LargeChange / 5
            End Select
        End With
End Select
    If HScroll.Width >= (400 * Screen.TwipsPerPixelX) Then  'More than 400 px
        'HScroll.SmallChange = HScroll.SmallChange * 0.5
    ElseIf HScroll.Width < (400 * Screen.TwipsPerPixelX) And HScroll.Width >= (150 * Screen.TwipsPerPixelX) Then
        HScroll.SmallChange = HScroll.SmallChange * 2
    Else
        HScroll.SmallChange = HScroll.SmallChange * 4
    End If
    
    If VScroll.Height >= (400 * Screen.TwipsPerPixelY) Then  'More than 400 px
        'VScroll.SmallChange = VScroll.SmallChange * 0.5
    ElseIf VScroll.Height < (400 * Screen.TwipsPerPixelY) And VScroll.Height >= (150 * Screen.TwipsPerPixelX) Then
        VScroll.SmallChange = VScroll.SmallChange * 2
    Else
        VScroll.SmallChange = VScroll.SmallChange * 4   ' les than 150
    End If
    

End Sub



Private Sub HScroll_Change()
PositionContainedControls
End Sub
Private Sub HScroll_Scroll()
PositionContainedControls
End Sub
Private Sub VScroll_Change()
PositionContainedControls
End Sub
Private Sub VScroll_Scroll()
PositionContainedControls
End Sub


Private Sub PositionContainedControls()
On Error Resume Next
Dim ctrl As Control

Select Case m_ScrollBars
    Case vbSBNone

    Case vbHorizontal
        For Each ctrl In UserControl.ContainedControls
            ctrl.Move (ctrl.Left + HPrevValue) - HScroll.Value
        Next
        HPrevValue = HScroll.Value

    Case vbVertical
        For Each ctrl In UserControl.ContainedControls
            ctrl.Move ctrl.Left, (ctrl.Top + VPrevValue) - VScroll.Value
        Next
        VPrevValue = VScroll.Value
    
    Case vbBoth
        For Each ctrl In UserControl.ContainedControls
            ctrl.Move (ctrl.Left + HPrevValue) - HScroll.Value, (ctrl.Top + VPrevValue) - VScroll.Value
        Next
        HPrevValue = HScroll.Value
        VPrevValue = VScroll.Value
End Select
End Sub


Private Function OnArea() As Boolean
    Dim mpos As POINTAPI
    Dim oRect As Rect
    GetCursorPos mpos
    GetWindowRect Me.hWnd, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        OnArea = True
    Else
        OnArea = False
   End If
End Function


Public Sub ScrollUp()
If OnArea = True Then
    Select Case m_ScrollBehavior
        Case [Normal]
            If VScroll.Value >= VScroll.SmallChange Then
                VScroll.Value = VScroll.Value - VScroll.SmallChange
            ElseIf VScroll.Value = VScroll.Min Then
                If HScroll.Value >= HScroll.SmallChange Then
                    HScroll.Value = HScroll.Value - HScroll.SmallChange
                Else
                    HScroll.Value = HScroll.Min
                End If
            Else
                VScroll.Value = VScroll.Min
            End If
        Case [Middle], [Reverse]
            If HScroll.Value >= HScroll.SmallChange Then
                HScroll.Value = HScroll.Value - HScroll.SmallChange
            Else
                HScroll.Value = HScroll.Min
                If VScroll.Value >= VScroll.SmallChange Then
                    VScroll.Value = VScroll.Value - VScroll.SmallChange
                Else
                    VScroll.Value = VScroll.Min
                End If
            End If
    End Select
End If
End Sub


Public Sub ScrollDown()
If OnArea = True Then
    Select Case m_ScrollBehavior
        Case [Normal], [Middle]
            If VScroll.Value <= VScroll.Max - VScroll.SmallChange Then
                VScroll.Value = VScroll.Value + VScroll.SmallChange
            ElseIf VScroll.Value = VScroll.Max Then
                If HScroll.Value <= HScroll.Max - HScroll.SmallChange Then
                    HScroll.Value = HScroll.Value + HScroll.SmallChange
                Else
                    HScroll.Value = HScroll.Max
                End If
            Else
                VScroll.Value = VScroll.Max
            End If
        Case [Reverse]
            If HScroll.Value <= HScroll.Max - HScroll.SmallChange Then
                HScroll.Value = HScroll.Value + HScroll.SmallChange
            Else
                HScroll.Value = HScroll.Max
                If VScroll.Value <= VScroll.Max - VScroll.SmallChange Then
                    VScroll.Value = VScroll.Value + VScroll.SmallChange
                Else
                    VScroll.Value = VScroll.Max
                End If
            End If

    End Select
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    bCancel = True              'IDE SAFE method, if go to debug mode,
    FocusTimer.Enabled = False  'Stop ProcessMessages
End Sub


Private Sub ProcessMessages()
Dim Message As msg
On Error GoTo Err_Site
Dim ctrl As Control
Dim sec As Integer
sec = Second(Now)

Do While Not bCancel
    If Not UserControl.Ambient.UserMode = True Then Exit Do
    WaitMessage 'Wait For message
    'if the mousewheel is used:
    If PeekMessage(Message, UserControl.Parent.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then
        If Message.wParam < 0 Then 'scroll up
            For Each ctrl In UserControl.Parent.Controls
                If TypeOf ctrl Is IoxContainer Then
                    ctrl.ScrollDown
                End If
            Next
        Else        'scroll down
            For Each ctrl In UserControl.Parent.Controls
                If TypeOf ctrl Is IoxContainer Then
                    ctrl.ScrollUp
                End If
            Next
        End If
    End If
    DoEvents
    If (Second(Now) - sec) > 1 Then bCancel = True
    Loop
Err_Site:
    If Err.Number = 398 Then bCancel = True
End Sub





Private Sub FocusTimer_timer()
On Error Resume Next
If Not UserControl.Ambient.UserMode = True Then
    FocusTimer.Enabled = False  'If in design view
    Exit Sub
End If

CurrCtrl.Name = UserControl.Parent.ActiveControl.Name
CurrCtrl.Index = UserControl.Parent.ActiveControl.Index

'Determine if control name or index has changed
If (CurrCtrl.Name <> LastCtrl.Name) Or (CurrCtrl.Index <> LastCtrl.Index) Then
    If CurrCtrl.Name <> LastCtrl.Name Then LastCtrl.Name = CurrCtrl.Name
    If CurrCtrl.Index <> LastCtrl.Index Then LastCtrl.Index = CurrCtrl.Index
    
    Dim LastHwnd As Long
    Dim CtrlContainer As Object
    Dim CtrlPositionx  As Long
    Dim CtrlPositionY  As Long

    CtrlPositionY = 0
    
    Set TempControl = UserControl.Parent.ActiveControl
    Set CtrlContainer = TempControl.Container   'Parent of focused control
    
    Do
        LastHwnd = CtrlContainer.hWnd
        If LastHwnd = UserControl.hWnd Then Exit Do      'Parent is IoxContainer
        CtrlPositionY = CtrlPositionY + CtrlContainer.Top
        CtrlPositionx = CtrlPositionx + CtrlContainer.Left
        Err.Clear
        Set CtrlContainer = CtrlContainer.Container
        If Err.Number <> 0 Then Exit Do         'No parent
    Loop
End If

If Not LastHwnd = Me.hWnd Then Exit Sub          'Active Control is't IoxContainer
    
Dim TempValue As Long, CtrlTop As Long, CtrlLeft As Long

'Determine if control is out of vertical viewing range

CtrlTop = CtrlPositionY + TempControl.Top '- 50
TempValue = TempControl.Height
If TempValue > VScroll.Height Then TempValue = VScroll.Height - 175
CtrlPositionY = CtrlPositionY + TempControl.Top + TempValue
'If the Control is outside of the Vertical viewing area, change the VScroll
If CtrlTop < 0 Then
    VScroll.Value = VScroll.Value + CtrlTop
ElseIf CtrlPositionY > VScroll.Height Then
    VScroll.Value = VScroll.Value + (CtrlPositionY - (VScroll.Height))
End If

'Determine if control is out of horizontal viewing range
CtrlLeft = CtrlPositionx + TempControl.Left '- 50
TempValue = TempControl.Width
If TempValue > HScroll.Width Then TempValue = HScroll.Width - 175
CtrlPositionx = CtrlPositionx + TempControl.Left + TempValue
'If the Control is outside of the Horizontal viewing area, change the HScroll
If CtrlLeft < 0 Then
    HScroll.Value = HScroll.Value + CtrlLeft
ElseIf CtrlPositionx > HScroll.Width Then
    HScroll.Value = HScroll.Value + (CtrlPositionx - (HScroll.Width))
End If

End Sub
