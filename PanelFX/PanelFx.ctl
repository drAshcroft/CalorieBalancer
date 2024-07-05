VERSION 5.00
Begin VB.UserControl PanelFx 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   ControlContainer=   -1  'True
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   139
End
Attribute VB_Name = "PanelFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Private Declare Function SetRect Lib "USER32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "USER32" ()
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum GRADIENT_DIR
    Horizontal = &H0
    Vertical = &H1
End Enum

Enum BackGroundStyle
    bSoildColor = 0
    bGradient = 1
    bBitmap = 2
End Enum

Enum BoxRect
    OutLine = 0
    Filled = 1
End Enum

Enum Alignment
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Enum m_PanelArea
    Title = 0
    Panel = 1
End Enum

Const m_def_PanelBorderColor = &H400000
Dim m_PanelBorderColor As OLE_COLOR
Dim m_TitleHeight As Long
'Default Property Values:
Const m_def_HideTitleIcon = 0
Const m_def_TitleIconWidth = 32
Const m_def_TitleIconHeight = 32
Const m_def_TitleIconXPos = 8
Const m_def_TitleIconYPos = 6
Const m_def_gCPanelStart = &HDE9A39
Const m_def_gCPanelEnd = &HFFFFFF
Const m_def_gCTitleStart = &HDE9A39
Const m_def_gCTitleEnd = &HFFFFFF
Const m_def_AllowDraging = 0
Const m_def_CanCollapse = 0
Const m_def_RoundEdge = 0
Const m_def_PanelBackColor = &H80000005
Const m_def_TileBackColor = vbHighlight
Const m_def_TitleForeColor = vbWhite
Const m_def_TitleCaption = "Panel"
'Property Variables:
Dim m_HideTitleIcon As Boolean
Dim m_TitleIconWidth As Integer
Dim m_TitleIconHeight As Integer
Dim m_TitleIconXPos As Integer
Dim m_TitleIconYPos As Integer
Dim m_TitleIcon As StdPicture
Dim m_TitleBmp As StdPicture, m_PanelBmp As StdPicture
Dim m_gCPanelStart As OLE_COLOR
Dim m_gCPanelEnd As OLE_COLOR
Dim m_gCTitleStart As OLE_COLOR
Dim m_gCTitleEnd As OLE_COLOR
Dim m_gTitleDir As GRADIENT_DIR
Dim m_gPanelDir As GRADIENT_DIR
Dim m_AllowDraging As Boolean
Dim m_AllowParentDraging As Boolean
Dim m_CanCollapse As Boolean
Dim m_RoundEdge As Long
Dim m_PanelBackColor As OLE_COLOR
Dim m_TileBackColor As OLE_COLOR
Dim m_TitleForeColor As OLE_COLOR
Dim m_TitleCaption As String
Dim m_TitleAlignment As Alignment
Dim m_TextHeight As Long
Dim m_BackGroundStyle As BackGroundStyle
Dim m_HasTitleIcon As Boolean
'
Dim Old_Height As Long
'Events
Event TileClick()
Event PanelClick()
'Event Declarations:

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub

Public Sub Collapse(mCollapse As Boolean)
    'Expand or Collapse the panel
    If (m_CanCollapse = False) Then Exit Sub
    If (mCollapse) Then
        'Expand Panel
        Old_Height = (UserControl.ScaleHeight * Screen.TwipsPerPixelX)
        UserControl.Height = (m_TitleHeight * Screen.TwipsPerPixelX)
    Else
        'Collapse Panel
        UserControl.Height = Old_Height
    End If
    
End Sub

Private Function GDI_TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, GDI_TranslateColor) Then
        GDI_TranslateColor = &HFFFF&
    End If
End Function

Private Sub GDI_Box(hdc As Long, lRect As Rect, hColor As OLE_COLOR, hStyle As BoxRect)
Dim hBrush As Long

    hBrush = CreateSolidBrush(GDI_TranslateColor(hColor))
    
    If hStyle = Filled Then
        FillRect hdc, lRect, hBrush    'Draw a filled box
    Else
        FrameRect hdc, lRect, hBrush  'Draw outline for the box
    End If
    
    DeleteObject hBrush
End Sub

Private Sub GDI_GradientFill(hdc As Long, mRect As Rect, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), GDI_TranslateColor(mEndColor)
    tTV(0).X = mRect.Left
    tTV(0).Y = mRect.Top
    
    setTriVertexColor tTV(0), GDI_TranslateColor(mStartColor)
    tTV(1).X = mRect.Right
    tTV(1).Y = mRect.Bottom
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Private Function PrintTitle()
Dim y_pos As Long, x_pos As Long

    With UserControl
        m_TextHeight = .TextHeight(m_TitleCaption)
        .ForeColor = m_TitleForeColor
        y_pos = (m_TitleHeight * 0.4) - (.TextHeight(m_TitleCaption) * 0.4)
        If (m_TitleAlignment = aLeft) Then
            If (m_HasTitleIcon) Then
                x_pos = 4
                x_pos = m_TitleIconWidth + x_pos
            Else
                x_pos = 4
            End If
            
        ElseIf (m_TitleAlignment = aCenter) Then
            x_pos = (.ScaleWidth - .TextWidth(m_TitleCaption)) \ 2 - 4
        ElseIf (m_TitleAlignment = aRight) Then
            x_pos = .ScaleWidth - .TextWidth(m_TitleCaption) - 4
        End If
        
        .CurrentX = x_pos
        .CurrentY = y_pos
        UserControl.Print m_TitleCaption
        
    End With
    
End Function

Function PanelOrTitle(Y As Single) As m_PanelArea
    If (Y <= 0) Or (Y < m_TitleHeight) Then
        PanelOrTitle = Title
    Else
        PanelOrTitle = Panel
    End If
End Function

Private Sub RenderPanel()
Dim Center_Offset As Long, hBrush As Long
Dim rc As Rect
Dim rgn As Long
On Error Resume Next

    With UserControl
        .Cls
        .BackColor = 0
        rgn = CreateRoundRectRgn(0, 0, .ScaleWidth, .ScaleHeight, m_RoundEdge, m_RoundEdge)
        SetWindowRgn .hWnd, rgn, True
        'Draw the title box outline
        SetRect rc, 1, 1, .ScaleWidth, m_TitleHeight
        GDI_Box .hdc, rc, m_PanelBorderColor, OutLine
        'Draw the filled area inside the tile
        SetRect rc, 0, 1, .ScaleWidth - 1, m_TitleHeight - 1
        If (BackGroundStyle = bSoildColor) Then
            'Fill with soild color
            GDI_Box .hdc, rc, m_TileBackColor, Filled
        ElseIf (BackGroundStyle = bGradient) Then
            'Fill with Gradient effect
            GDI_GradientFill .hdc, rc, m_gCTitleStart, m_gCTitleEnd, gCTitleDir
        ElseIf (BackGroundStyle = bBitmap) Then
            'Fill with Texture
            hBrush = CreatePatternBrush(m_TitleBmp.handle)
            FillRect .hdc, rc, hBrush
            DeleteObject hBrush
        End If
        '
        'Draw the bottom part of the panel outline
        SetRect rc, 0, m_TitleHeight - 1, .ScaleWidth, .ScaleHeight
        GDI_Box .hdc, rc, m_PanelBorderColor, OutLine
        'Draw the filled bottom panel area
        SetRect rc, 1, m_TitleHeight, .ScaleWidth - 1, .ScaleHeight - 1
        
        If (BackGroundStyle = bSoildColor) Then
            'Fill with soild color
            GDI_Box .hdc, rc, m_PanelBackColor, Filled
        ElseIf (BackGroundStyle = bGradient) Then
            'Fill with Gradient effect
            GDI_GradientFill .hdc, rc, m_gCPanelStart, m_gCPanelEnd, m_gPanelDir
        ElseIf (BackGroundStyle = bBitmap) Then
            'Fill with Texture
            hBrush = CreatePatternBrush(m_PanelBmp.handle)
            FillRect .hdc, rc, hBrush
            DeleteObject hBrush
        End If

        'Update controls forcolor with the panel outline color
        .ForeColor = m_PanelBorderColor
        'Used to add some roundness to the panel
        RoundRect .hdc, 0, 0, .ScaleWidth - 1, .ScaleHeight - 1, m_RoundEdge, m_RoundEdge
        
        'Hide a panels icon
        m_HasTitleIcon = Not m_HideTitleIcon
        
        'Check that the panel has an picture and show icon is enabled
        If Not m_HideTitleIcon Then
            If m_TitleIcon Is Nothing Then
                m_HasTitleIcon = False
            Else
                'Paint on the panel titles picture
                m_HasTitleIcon = True
                UserControl.PaintPicture m_TitleIcon, m_TitleIconXPos, m_TitleIconYPos, m_TitleIconWidth, m_TitleIconHeight
            End If
        End If
        
        'Print on the Title
        If Len(m_TitleCaption) <> 0 Then
            Call PrintTitle
        End If
        
        'Update the control
        .Refresh
    End With
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y, PanelOrTitle(Y))
    If (Button <> vbLeftButton) Then Exit Sub
    
    Select Case PanelOrTitle(Y)
        Case Title
            RaiseEvent TileClick
        Case Panel
            RaiseEvent PanelClick
    End Select
    
End Sub

Private Sub UserControl_Resize()
    RenderPanel
End Sub

Private Sub UserControl_Show()
    RenderPanel
End Sub

'Property Stuff

Public Property Get TileHeight() As Long
    TileHeight = m_TitleHeight
End Property

Public Property Let TileHeight(ByVal New_TileHeight As Long)
    m_TitleHeight = New_TileHeight
    PropertyChanged "TileHeight"
    Call RenderPanel
End Property

Private Sub UserControl_InitProperties()
    m_TitleHeight = 20
    m_PanelBorderColor = m_def_PanelBorderColor
    m_Caption = m_def_Caption
    m_TitleCaption = UserControl.Ambient.DisplayName
    m_TitleAlignment = aLeft
    m_BackGroundStyle = bSoildColor
    m_TitleForeColor = m_def_TitleForeColor
    Set UserControl.Font = Ambient.Font
    m_TileBackColor = m_def_TileBackColor
    m_PanelBackColor = m_def_PanelBackColor
    m_RoundEdge = m_def_RoundEdge
    m_CanCollapse = m_def_CanCollapse
    m_AllowDraging = m_def_AllowDraging
    m_AllowParentDraging = False
    m_gCTitleStart = m_def_gCTitleStart
    m_gCTitleEnd = m_def_gCTitleEnd
    m_gTitleDir = Vertical
    m_gPanelDir = Vertical
    
    m_gCPanelStart = m_def_gCPanelStart
    m_gCPanelEnd = m_def_gCPanelEnd
    Set m_TitleIcon = Nothing
    
    Set m_TitleBmp = Nothing
    Set m_PanelBmp = Nothing
    
    m_TitleIconWidth = m_def_TitleIconWidth
    m_TitleIconHeight = m_def_TitleIconHeight
    m_TitleIconXPos = m_def_TitleIconXPos
    m_TitleIconYPos = m_def_TitleIconYPos
    m_HideTitleIcon = m_def_HideTitleIcon
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TitleHeight = PropBag.ReadProperty("TileHeight", 20)
    m_PanelBorderColor = PropBag.ReadProperty("PanelBorderColor", m_def_PanelBorderColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_TitleCaption = PropBag.ReadProperty("TitleCaption", m_def_TitleCaption)
    m_TitleForeColor = PropBag.ReadProperty("TitleForeColor", m_def_TitleForeColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TitleAlignment = PropBag.ReadProperty("TitleAlignment", aLeft)
    m_TileBackColor = PropBag.ReadProperty("TileBackColor", m_def_TileBackColor)
    m_PanelBackColor = PropBag.ReadProperty("PanelBackColor", m_def_PanelBackColor)
    m_RoundEdge = PropBag.ReadProperty("RoundEdge", m_def_RoundEdge)
    m_CanCollapse = PropBag.ReadProperty("CanCollapse", m_def_CanCollapse)
    m_AllowDraging = PropBag.ReadProperty("AllowDraging", m_def_AllowDraging)
    m_BackGroundStyle = PropBag.ReadProperty("BackGroundStyle", bSoildColor)
    m_gCTitleStart = PropBag.ReadProperty("gCTitleStart", m_def_gCTitleStart)
    m_gCTitleEnd = PropBag.ReadProperty("gCTitleEnd", m_def_gCTitleEnd)
    m_gTitleDir = PropBag.ReadProperty("gCTitleDir", Vertical)
    m_gCPanelStart = PropBag.ReadProperty("gCPanelStart", m_def_gCPanelStart)
    m_gCPanelEnd = PropBag.ReadProperty("gCPanelEnd", m_def_gCPanelEnd)
    m_gPanelDir = PropBag.ReadProperty("gCPanelDir", Vertical)
    Set m_TitleIcon = PropBag.ReadProperty("TitleIcon", Nothing)
    m_TitleIconWidth = PropBag.ReadProperty("TitleIconWidth", m_def_TitleIconWidth)
    m_TitleIconHeight = PropBag.ReadProperty("TitleIconHeight", m_def_TitleIconHeight)
    m_TitleIconXPos = PropBag.ReadProperty("TitleIconXPos", m_def_TitleIconXPos)
    m_TitleIconYPos = PropBag.ReadProperty("TitleIconYPos", m_def_TitleIconYPos)
    m_HideTitleIcon = PropBag.ReadProperty("HideTitleIcon", m_def_HideTitleIcon)
    m_AllowParentDraging = PropBag.ReadProperty("AllowParentDraging", False)
    Set m_TitleBmp = PropBag.ReadProperty("TitleBitmap", Nothing)
    Set m_PanelBmp = PropBag.ReadProperty("PanelBitmap", Nothing)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TileHeight", m_TitleHeight, 20)
    Call PropBag.WriteProperty("PanelBorderColor", m_PanelBorderColor, m_def_PanelBorderColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("TitleCaption", m_TitleCaption, m_def_TitleCaption)
    Call PropBag.WriteProperty("TitleForeColor", m_TitleForeColor, m_def_TitleForeColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleAlignment", m_TitleAlignment, aLeft)
    Call PropBag.WriteProperty("TileBackColor", m_TileBackColor, m_def_TileBackColor)
    Call PropBag.WriteProperty("PanelBackColor", m_PanelBackColor, m_def_PanelBackColor)
    Call PropBag.WriteProperty("RoundEdge", m_RoundEdge, m_def_RoundEdge)
    Call PropBag.WriteProperty("CanCollapse", m_CanCollapse, m_def_CanCollapse)
    Call PropBag.WriteProperty("AllowDraging", m_AllowDraging, m_def_AllowDraging)
    Call PropBag.WriteProperty("BackGroundStyle", m_BackGroundStyle, bSoildColor)
    Call PropBag.WriteProperty("gCTitleStart", m_gCTitleStart, m_def_gCTitleStart)
    Call PropBag.WriteProperty("gCTitleEnd", m_gCTitleEnd, m_def_gCTitleEnd)
    Call PropBag.WriteProperty("gCTitleDir", m_gTitleDir, Vertical)
    Call PropBag.WriteProperty("gCPanelStart", m_gCPanelStart, m_def_gCPanelStart)
    Call PropBag.WriteProperty("gCPanelEnd", m_gCPanelEnd, m_def_gCPanelEnd)
    Call PropBag.WriteProperty("gCPanelDir", m_gPanelDir, Vertical)
    Call PropBag.WriteProperty("TitleIcon", m_TitleIcon, Nothing)
    Call PropBag.WriteProperty("TitleIconWidth", m_TitleIconWidth, m_def_TitleIconWidth)
    Call PropBag.WriteProperty("TitleIconHeight", m_TitleIconHeight, m_def_TitleIconHeight)
    Call PropBag.WriteProperty("TitleIconXPos", m_TitleIconXPos, m_def_TitleIconXPos)
    Call PropBag.WriteProperty("TitleIconYPos", m_TitleIconYPos, m_def_TitleIconYPos)
    Call PropBag.WriteProperty("HideTitleIcon", m_HideTitleIcon, m_def_HideTitleIcon)
    Call PropBag.WriteProperty("AllowParentDraging", m_AllowParentDraging, False)
    Call PropBag.WriteProperty("TitleBitmap", m_TitleBmp, Nothing)
    Call PropBag.WriteProperty("PanelBitmap", m_PanelBmp, Nothing)
    
End Sub

Public Property Get PanelBorderColor() As OLE_COLOR
    PanelBorderColor = m_PanelBorderColor
End Property

Public Property Let PanelBorderColor(ByVal New_PanelBorderColor As OLE_COLOR)
    m_PanelBorderColor = New_PanelBorderColor
    PropertyChanged "PanelBorderColor"
    Call RenderPanel
End Property

Public Property Get TitleCaption() As String
    TitleCaption = m_TitleCaption
End Property

Public Property Let TitleCaption(ByVal New_TitleCaption As String)
    m_TitleCaption = New_TitleCaption
    PropertyChanged "TitleCaption"
    Call RenderPanel
End Property

Public Property Get TitleForeColor() As OLE_COLOR
    TitleForeColor = m_TitleForeColor
End Property

Public Property Let TitleForeColor(ByVal New_TitleForeColor As OLE_COLOR)
    m_TitleForeColor = New_TitleForeColor
    PropertyChanged "TitleForeColor"
    Call RenderPanel
End Property

Public Property Get TitleAlignment() As Alignment
   TitleAlignment = m_TitleAlignment
End Property

Public Property Let TitleAlignment(ByVal New_Align As Alignment)
    m_TitleAlignment = New_Align
    PropertyChanged "TitleAlignment"
    Call RenderPanel
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call RenderPanel
End Property

Public Property Get TileBackColor() As OLE_COLOR
    TileBackColor = m_TileBackColor
End Property

Public Property Let TileBackColor(ByVal New_TileBackColor As OLE_COLOR)
    m_TileBackColor = New_TileBackColor
    PropertyChanged "TileBackColor"
    Call RenderPanel
End Property

Public Property Get PanelBackColor() As OLE_COLOR
    PanelBackColor = m_PanelBackColor
End Property

Public Property Let PanelBackColor(ByVal New_PanelBackColor As OLE_COLOR)
    m_PanelBackColor = New_PanelBackColor
    PropertyChanged "PanelBackColor"
    Call RenderPanel
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iHwnd As Long
    RaiseEvent MouseMove(Button, Shift, X, Y, PanelOrTitle(Y))
    
    If Button <> vbLeftButton Then Exit Sub
    
    If (PanelOrTitle(Y) = Title) And (AllowDraging) Then
        'Check for parent draging
        'If part draging is enabled only move the controls parent
        
        If (m_AllowParentDraging) Then
            iHwnd = UserControl.Parent.hWnd
        Else
            iHwnd = UserControl.hWnd
        End If
        
        Call ReleaseCapture
        Call SendMessage(iHwnd, &HA1, 2, 0&)
        
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y, PanelOrTitle(Y))
End Sub

Public Property Get RoundEdge() As Long
    RoundEdge = m_RoundEdge
End Property

Public Property Let RoundEdge(ByVal New_RoundEdge As Long)
    m_RoundEdge = New_RoundEdge
    PropertyChanged "RoundEdge"
    Call RenderPanel
End Property

Public Property Get CanCollapse() As Boolean
    CanCollapse = m_CanCollapse
End Property

Public Property Let CanCollapse(ByVal New_CanCollapse As Boolean)
    m_CanCollapse = New_CanCollapse
    PropertyChanged "CanCollapse"
End Property

Public Property Get AllowDraging() As Boolean
    AllowDraging = m_AllowDraging
End Property

Public Property Let AllowDraging(ByVal New_AllowDraging As Boolean)
    m_AllowDraging = New_AllowDraging
    PropertyChanged "AllowDraging"
End Property

Public Property Get BackGroundStyle() As BackGroundStyle
    BackGroundStyle = m_BackGroundStyle
End Property

Public Property Let BackGroundStyle(ByVal vNew_Style As BackGroundStyle)
    m_BackGroundStyle = vNew_Style
    PropertyChanged "BackGroundStyle"
    Call RenderPanel
End Property

Public Property Get gCTitleStart() As OLE_COLOR
    gCTitleStart = m_gCTitleStart
End Property

Public Property Let gCTitleStart(ByVal New_gCTitleStart As OLE_COLOR)
    m_gCTitleStart = New_gCTitleStart
    PropertyChanged "gCTitleStart"
    Call RenderPanel
End Property

Public Property Get gCTitleEnd() As OLE_COLOR
    gCTitleEnd = m_gCTitleEnd
End Property

Public Property Let gCTitleEnd(ByVal New_gCTitleEnd As OLE_COLOR)
    m_gCTitleEnd = New_gCTitleEnd
    PropertyChanged "gCTitleEnd"
    Call RenderPanel
End Property

Public Property Get gCTitleDir() As GRADIENT_DIR
    gCTitleDir = m_gTitleDir
End Property

Public Property Let gCTitleDir(ByVal New_DirT As GRADIENT_DIR)
    m_gTitleDir = New_DirT
    PropertyChanged "gCTitleDir"
    Call RenderPanel
End Property
'
Public Property Get gCPanelDir() As GRADIENT_DIR
    gCPanelDir = m_gPanelDir
End Property

Public Property Let gCPanelDir(ByVal New_DirT As GRADIENT_DIR)
    m_gPanelDir = New_DirT
    PropertyChanged "gCPanelDir"
    Call RenderPanel
End Property

Public Property Get gCPanelStart() As OLE_COLOR
    gCPanelStart = m_gCPanelStart
End Property

Public Property Let gCPanelStart(ByVal New_gCPanelStart As OLE_COLOR)
    m_gCPanelStart = New_gCPanelStart
    PropertyChanged "gCPanelStart"
    Call RenderPanel
End Property

Public Property Get gCPanelEnd() As OLE_COLOR
    gCPanelEnd = m_gCPanelEnd
End Property

Public Property Let gCPanelEnd(ByVal New_gCPanelEnd As OLE_COLOR)
    m_gCPanelEnd = New_gCPanelEnd
    PropertyChanged "gCPanelEnd"
    Call RenderPanel
End Property

Public Property Get TitleIcon() As StdPicture
    Set TitleIcon = m_TitleIcon
End Property

Public Property Set TitleIcon(ByVal New_TitleIcon As StdPicture)
    Set m_TitleIcon = New_TitleIcon
    PropertyChanged "TitleIcon"
    Call RenderPanel
End Property

Public Property Get TitleBitmap() As StdPicture
    Set TitleBitmap = m_TitleBmp
End Property

Public Property Set TitleBitmap(ByVal NewTBmp As StdPicture)
    Set m_TitleBmp = NewTBmp
    PropertyChanged "TitleBitmap"
    Call RenderPanel
End Property

Public Property Get PanelBitmap() As StdPicture
    Set PanelBitmap = m_PanelBmp
End Property

Public Property Set PanelBitmap(ByVal NewPanel As StdPicture)
    Set m_PanelBmp = NewPanel
    PropertyChanged "PanelBitmap"
    Call RenderPanel
End Property

Public Property Get TitleIconWidth() As Integer
    TitleIconWidth = m_TitleIconWidth
End Property

Public Property Let TitleIconWidth(ByVal New_TitleIconWidth As Integer)
    m_TitleIconWidth = New_TitleIconWidth
    PropertyChanged "TitleIconWidth"
    Call RenderPanel
End Property

Public Property Get TitleIconHeight() As Integer
    TitleIconHeight = m_TitleIconHeight
End Property

Public Property Let TitleIconHeight(ByVal New_TitleIconHeight As Integer)
    m_TitleIconHeight = New_TitleIconHeight
    PropertyChanged "TitleIconHeight"
    Call RenderPanel
End Property

Public Property Get TitleIconXPos() As Integer
    TitleIconXPos = m_TitleIconXPos
End Property

Public Property Let TitleIconXPos(ByVal New_TitleIconXPos As Integer)
    m_TitleIconXPos = New_TitleIconXPos
    PropertyChanged "TitleIconXPos"
    Call RenderPanel
End Property

Public Property Get TitleIconYPos() As Integer
    TitleIconYPos = m_TitleIconYPos
End Property

Public Property Let TitleIconYPos(ByVal New_TitleIconYPos As Integer)
    m_TitleIconYPos = New_TitleIconYPos
    PropertyChanged "TitleIconYPos"
    Call RenderPanel
End Property

Public Property Get HideTitleIcon() As Boolean
    HideTitleIcon = m_HideTitleIcon
End Property

Public Property Let HideTitleIcon(ByVal New_HideTitleIcon As Boolean)
    m_HideTitleIcon = New_HideTitleIcon
    PropertyChanged "HideTitleIcon"
    Call RenderPanel
End Property

Public Property Get AllowParentDraging() As Boolean
    AllowParentDraging = m_AllowParentDraging
End Property

Public Property Let AllowParentDraging(ByVal New_Drag As Boolean)
    m_AllowParentDraging = New_Drag
    PropertyChanged "AllowParentDraging"
    Call RenderPanel
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

