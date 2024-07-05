VERSION 5.00
Begin VB.UserControl ProgressBars 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox bar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      Picture         =   "ProgressBars.ctx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Background 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   660
      Picture         =   "ProgressBars.ctx":0A62
      ScaleHeight     =   1365
      ScaleWidth      =   1365
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "ProgressBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private Type LinesType
   Caption As String
   Percent As Single
   Value As String
End Type

Private lines() As LinesType
Private BarH As Long, StartY As Long, SpacingY As Long, StartX As Long

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Dim SHeight As Single, SWidth As Single

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Const GRADIENT_FILL_OP_FLAG As Long = &HFF


Private Const two8 = 2 ^ 8
Private Const two16 = 2 ^ 16

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal DY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

 Const PIDEG As Double = 1.74532925199433E-02
'Default Property Values:
Const m_def_HeadWidth = 12
Const m_def_Blend = 0
'Property Variables:
Dim m_HeadWidth As Long
Dim m_Blend As Single
'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."


'This code's base have been made by Branco Medeiros.
'As he wrote, he just reworked it from a Java color class.
'But I think he was too accurate: his class's color model
'used 360 values for hue, 100 saturation, and 100 for
'brightness. Using these maximal values, one can use only
'360*100*100 = 3,600,000 unique colors. Using RGB, with a
'maximal value of 255, there are 16,777,216 possibilities.
'It's far more than the original color resolution of the
'class, so I decided to use a maximum of 255 for each color
'component. This way, You can pack an HSB value to a Long
'with the RGB() function. This may cause some disturbance
'in your code, but you can always write a comment for
'yourself about the content of a Long.

'The way You can pack an HSB value to a Long with RGB:
'  HSB values to Long: RGB(Hue, Saturation, Brightness)

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Sub SplitLong(ByVal inVal As Long, ByRef Val1 As Long, ByRef Val2 As Long, ByRef Val3 As Long)
On Error Resume Next
Val1 = (inVal And &HFF)
Val2 = (inVal And &HFF00&) \ &HFF
Val3 = (inVal And &HFF0000) \ &H10000

End Sub

Private Function SplitLong_Place(inVal As Long, Place As Byte) As Long
On Error Resume Next
Select Case Place
Case 1
   SplitLong_Place = (inVal And &HFF)
Case 2
   SplitLong_Place = (inVal And &HFF00&) \ 256
Case 3
   SplitLong_Place = (inVal And &HFF0000) \ &H10000
End Select

End Function

Private Function HSBToRGB(H As Long, s As Long, L As Long) As Long
On Error Resume Next
Dim r As Long, g As Long, B As Long
'Dim H As Long, S As Long, L As Long
Dim nH As Single, nS As Single, nL As Single
Dim nF As Single, nP As Single, nQ As Single, nT As Single
Dim lH As Long

'   SplitLong HSBValue, H, S, L
   
   If s > 0 Then
      nH = H / 42.666666
     
      nL = L / 256
      nS = s / 256
      
      lH = Int(nH)
      nF = nH - lH
      nP = nL * (1 - nS)
      nQ = nL * (1 - nS * nF)
      nT = nL * (1 - nS * (1 - nF))
      
      Select Case lH
      Case 0
         r = nL * 255
         g = nT * 255
         B = nP * 255
      Case 1
         r = nQ * 255
         g = nL * 255
         B = nP * 255
      Case 2
         r = nP * 255
         g = nL * 255
         B = nT * 255
      Case 3
         r = nP * 255
         g = nQ * 255
         B = nL * 255
      Case 4
         r = nT * 255
         g = nP * 255
         B = nL * 255
      Case 5
         r = nL * 255
         g = nP * 255
         B = nQ * 255
      End Select
   Else
      r = L
      g = r
      B = r
   End If
     
   HSBToRGB = RGB(r, g, B)
  
End Function

Private Function RGBToHSB(ByVal RGBValue As Long, H As Long, s As Long, L As Long) As Long
On Error Resume Next
Dim nTemp As Single
Dim lMin As Long, lMax As Long, lDelta As Long
Dim r As Long, g As Long, B As Long


   SplitLong RGBValue, r, g, B
  
   If r > g Then
      If r > B Then
         lMax = r
      Else
         lMax = B
      End If
   Else
      If g > B Then
         lMax = g
      Else
         lMax = B
      End If
   End If
      
   If r < g Then
      If r < B Then
         lMin = r
      Else
         lMin = B
      End If
   Else
      If g < B Then
         lMin = g
      Else
         lMin = B
      End If
   End If
      
  
  lDelta = lMax - lMin
  
  L = lMax
  
   If lMax > 0 Then
      s = (lDelta / lMax) * 255
      If lDelta > 0 Then
         If lMax = r Then
            nTemp = (g - B) / lDelta
         ElseIf lMax = g Then
            nTemp = 2 + (B - r) / lDelta
         Else
            nTemp = 4 + (r - g) / lDelta
         End If
         H = nTemp * 42.666666
         If H < 0 Then H = H + 256
      End If
   End If
  
  RGBToHSB = RGB(H, s, L)

End Function


Private Function MiddleStretch(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, _
   ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcW As Long, ByVal SrcH As Long, clr As Long, ByVal Cut1 As Long, ByVal Cut2 As Long, AveClr) As Long
   On Error GoTo errhandl
    If DstW = 0 Or DstH = 0 Then Exit Function
    Dim B As Long, H As Long, f As Long, i As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As Long, Data2() As Long
    Dim Info As BITMAPINFO
    Dim Info2 As BITMAPINFO
    
    TmpDC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    ReDim Data1(DstW * DstH * 4 - 1)
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info.bmiHeader.biWidth = DstW
    Info.bmiHeader.biHeight = DstH
    Info.bmiHeader.biPlanes = 1
    Info.bmiHeader.biBitCount = 32
    Info.bmiHeader.biCompression = 0
    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    
    
    Sr2DC = CreateCompatibleDC(SrcDC)
    Sr2Bmp = CreateCompatibleBitmap(SrcDC, SrcW, SrcH)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data2(SrcW * SrcH * 4 - 1)
    Info2.bmiHeader.biSize = Len(Info.bmiHeader)
    Info2.bmiHeader.biWidth = SrcW
    Info2.bmiHeader.biHeight = SrcH
    Info2.bmiHeader.biPlanes = 1
    Info2.bmiHeader.biBitCount = 32
    Info2.bmiHeader.biCompression = 0
    BitBlt Sr2DC, 0, 0, SrcW, SrcH, SrcDC, SrcX, SrcY, vbSrcCopy
    GetDIBits Sr2DC, Sr2Bmp, 0, SrcH, Data2(0), Info2, 0
    
    Dim r As Single
    Dim Red As Single, Green As Single, Blue As Single
    Dim Src As Long
    Dim m_Blend As Single
    Dim i2 As Long
    Dim f2 As Long
    Dim j As Single, Jl As Long
    Dim B_clr As Long
    Dim NClr As Long
    B_clr = UserControl.BackColor
    m_Blend = 0.1
    If Cut1 + Cut2 > DstW Then
       Cut1 = DstW / 2
       Cut2 = DstW / 2
    
    End If
    NClr = 0
    For H = 0 To DstH - 1
        f = H * DstW
        j = (SrcH / DstH)
        f2 = Int(H * j) * SrcW
        For B = 0 To DstW - 1
            i = f + B
               If B < Cut1 Then
                  i2 = f2 + B
               ElseIf (DstW - B) < Cut2 Then
                  Jl = DstW - B
                  i2 = f2 + SrcW - Jl
               End If
               If Data2(i2) <> vbWhite Then
                Dim HS As Long, s As Long, L As Long
                Call RGBToHSB(Data2(i2), HS, s, L)
                HS = clr '(HS + clr) Mod 250
                AveClr = AveClr + HS
                NClr = NClr + 1
                Data1(i) = HSBToRGB(HS, s, L)
            Else
             ' Data1(i) = &HC0C0C0
            End If
        Next B
    Next H
    If NClr = 0 Then NClr = 1
    AveClr = AveClr / NClr
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0
errhandl:
    Erase Data1
    Erase Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    DeleteDC TmpDC
    DeleteDC Sr2DC
End Function



Private Sub DrawBar(StartY As Long, ByVal Percent As Single, BarH As Long, Caption As String, ByVal Value As String)
  On Error GoTo errhandl
  If Percent < 0 Then Percent = 0
  Dim BarW As Long
  Dim bColor As Long
  
  BarW = Percent / 100 * (SWidth - StartX) * 0.85
  
  UserControl.FillStyle = 1
 
  'Debug.Print Percent
  If Percent < 75 Then
    bColor = 3 'Percent * 0.7
  ElseIf Percent < 90 Then
    bColor = 24 '(Percent - 100) * 4 + 100
  ElseIf Percent > 110 Then
    bColor = 180 '175 + (Percent - 130)
  Else
    bColor = 26
  End If
 ' bColor = Percent * 4
  'Debug.Print bColor
  Dim AveClr

  MiddleStretch UserControl.hdc, StartX, StartY, BarW, BarH, _
  bar.hdc, 0, 0, bar.ScaleWidth, bar.ScaleHeight, bColor, m_HeadWidth, m_HeadWidth, AveClr
  Dim TW As Long
  TW = UserControl.TextWidth(Caption)
  UserControl.CurrentX = 5
  UserControl.CurrentY = StartY + BarH / 2 - TextHeight("!") / 2
  UserControl.ForeColor = 0
  UserControl.Print Caption
  
  Dim r As Long, g As Long, B As Long
  Call SplitLong(CLng(AveClr), r, g, B)
  UserControl.ForeColor = 0 'RGB(255 - r, 255 - g, 255 - B)
    
  If Percent <= 95 Then
     UserControl.CurrentX = StartX + BarW
     UserControl.CurrentY = StartY + BarH / 2 - UserControl.TextHeight(Value) / 2
     UserControl.Print Value
  Else
     UserControl.CurrentX = StartX + (SWidth - StartX) * 0.85 - UserControl.TextWidth(Value & " ")
     UserControl.CurrentY = StartY + BarH / 2 - UserControl.TextHeight(Value) / 2
     UserControl.Print Value
  End If
errhandl:
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
On Error Resume Next
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
On Error Resume Next
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
On Error Resume Next
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_DblClick()
On Error Resume Next
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get Blend() As Single
On Error Resume Next
    Blend = m_Blend
End Property

Public Property Let Blend(ByVal New_Blend As Single)
On Error Resume Next
    m_Blend = New_Blend
    PropertyChanged "Blend"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Background,Background,-1,Picture
Public Property Get WholeBackground() As Picture
Attribute WholeBackground.VB_Description = "Returns/sets a graphic to be displayed in a control."
On Error Resume Next
    Set WholeBackground = Background.Picture
End Property

Public Property Set WholeBackground(ByVal New_WholeBackground As Picture)
On Error Resume Next
    Set Background.Picture = New_WholeBackground
    PropertyChanged "WholeBackground"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=bar,bar,-1,Picture
Public Property Get Bars() As Picture


    On Error GoTo Err_Proc
   ' Set Bars = bar.Picture
Exit_Proc:
    Exit Property


Err_Proc:
    Err_Handler "ProgressBars", "Bars", Err.Description
    Resume Exit_Proc


End Property

Public Property Set Bars(ByVal New_Bars As Picture)


    On Error GoTo Err_Proc
On Error Resume Next
  '  Set bar.Picture = New_Bars
    PropertyChanged "Bars"
Exit_Proc:
    Exit Property


Err_Proc:
    Err_Handler "ProgressBars", "Bars", Err.Description
    Resume Exit_Proc


End Property

Public Sub Clear()
On Error Resume Next
   Erase lines
   ReDim lines(0)
End Sub

Public Function Rows() As Long
On Error Resume Next
   Rows = UBound(lines)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddLine(Percent As Single, Caption As String, Value As String) As Variant
On Error Resume Next
ReDim Preserve lines(UBound(lines) + 1)
With lines(UBound(lines))
  .Caption = " " & Caption
  .Percent = Percent
  .Value = Value
End With
End Function
Public Sub UpdateLine(Index As Long, Percent As Single, Value As String)
On Error Resume Next
   With lines(Index + 1)
     .Percent = Percent
     .Value = Value
   End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Draw() As Variant
On Error GoTo errhandl
   'UserControl.ScaleMode = vbPixels
   'UserControl.FillStyle = 0
   'UserControl.FillColor = vbWhite
   
   'Call RoundRect(UserControl.hdc, 0, 0, SWidth, SHeight, StartY / 2, StartY / 2)
      Dim BW As Long, BH As Long
      Dim BorW As Long
      BW = Background.Width
      BH = Background.Height
      BorW = 15
      
   UserControl.PaintPicture Background.Picture, BorW, 0, SWidth - 2 * BorW, BorW, BorW, 0, BW - 2 * BorW, BorW
   UserControl.PaintPicture Background.Picture, BorW, SHeight - BorW, SWidth - 2 * BorW, BorW, BorW, BH - BorW, BW - 2 * BorW, BorW
   UserControl.PaintPicture Background.Picture, 0, BorW, BorW, SHeight - BorW * 2, 0, BorW, BorW, BH - BorW * 2
   UserControl.PaintPicture Background.Picture, SWidth - BorW, BorW, BorW, SHeight - BorW * 2, BW - BorW, BorW, BorW, BH - BorW * 2
   
   UserControl.PaintPicture Background.Picture, 0, 0, BorW, BorW, 0, 0, BorW, BorW
   UserControl.PaintPicture Background.Picture, 0, SHeight - BorW, BorW, BorW, 0, BH - BorW, BorW, BorW
   UserControl.PaintPicture Background.Picture, SWidth - BorW, 0, BorW, BorW, BW - BorW, 0, BorW, BorW
   UserControl.PaintPicture Background.Picture, SWidth - BorW, SHeight - BorW, BorW, BorW, BW - BorW, BH - BorW, BorW, BorW
   
   UserControl.PaintPicture Background.Picture, BorW, BorW, SWidth - BorW * 2, SHeight - BorW * 2, BorW, BorW, BW - BorW * 2, BH - BorW * 2
   
   
      
  UserControl.DrawWidth = 3
  UserControl.Line (SWidth * 0.85, 3)-(SWidth * 0.85, SHeight - 6), RGB(0, 200, 100)
  
  
  Dim i As Long, p As Single
  Dim j As Long
  StartX = 0
  For i = 1 To UBound(lines)
    j = UserControl.TextWidth(lines(i).Caption)
    If j > StartX Then StartX = j
    
  Next i
  StartX = StartX + 3
  If StartX = 3 Then StartX = SWidth * 0.1
  UserControl.Line (StartX, 3)-(StartX, SHeight - 6), RGB(0, 200, 100)
  UserControl.DrawWidth = 1
  
  StartX = StartX + 3
  
  For i = 1 To UBound(lines)
     p = lines(i).Percent
     DrawBar StartY + SpacingY * (i - 1), p, BarH, lines(i).Caption, lines(i).Value
  Next i
errhandl:
  UserControl.Refresh
End Function

Private Sub UserControl_Initialize()
On Error Resume Next
  ReDim lines(0)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    Set UserControl.Font = Ambient.Font
    m_Blend = m_def_Blend
    m_HeadWidth = m_def_HeadWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
   ' Set Picture = PropBag.ReadProperty("WholeBackground", Nothing)
   ' Set bar.Picture = PropBag.ReadProperty("Bars", Nothing)
    m_HeadWidth = PropBag.ReadProperty("HeadWidth", m_def_HeadWidth)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
SHeight = UserControl.ScaleHeight
SWidth = UserControl.ScaleWidth
StartY = 12
BarH = SHeight / 10
SpacingY = BarH + 3
Call Draw
UserControl.Refresh
UserControl.MaskColor = Background.Point(1, 1)
Set UserControl.MaskPicture = UserControl.Image
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
    Call PropBag.WriteProperty("WholeBackground", Picture, Nothing)
    Call PropBag.WriteProperty("Bars", Picture, Nothing)
    Call PropBag.WriteProperty("HeadWidth", m_HeadWidth, m_def_HeadWidth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,9
Public Property Get HeadWidth() As Long
On Error Resume Next
    HeadWidth = m_HeadWidth
End Property

Public Property Let HeadWidth(ByVal New_HeadWidth As Long)
On Error Resume Next
    m_HeadWidth = New_HeadWidth
    PropertyChanged "HeadWidth"
End Property


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

     Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
