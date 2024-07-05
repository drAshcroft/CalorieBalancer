VERSION 5.00
Begin VB.UserControl PieChart 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   5340
      Picture         =   "PieChart.ctx":0000
      ScaleHeight     =   3780
      ScaleWidth      =   3765
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4890
      Width           =   3765
   End
   Begin VB.PictureBox Display 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   375
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4650
      Visible         =   0   'False
      Width           =   4440
   End
End
Attribute VB_Name = "PieChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private Declare Function Pie Lib "gdi32" (ByVal hdc As _
    Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As _
    Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As _
    Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

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


Private Type Slice
    Angle As Single
    Caption As String
    Color As Long
End Type
Const Pi = 3.1415
Dim Slices() As Slice
Dim MidX As Single, midY As Single, W As Single, H As Single, RX As Single, RY As Single
'Default Property Values:
Const m_def_Blend = 1
'Property Variables:
Dim m_Blend As Single


Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Private Function FoxAlphaBlend(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long) As Long
On Error GoTo errhandl
    If DstW = 0 Or DstH = 0 Then Exit Function
    Dim B As Long, H As Long, f As Long, i As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As Long, Data2() As Long
    Dim Info As BITMAPINFO
    
    
    
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 4 - 1)
    ReDim Data2(DstW * DstH * 4 - 1)
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info.bmiHeader.biWidth = DstW
    Info.bmiHeader.biHeight = DstH
    Info.bmiHeader.biPlanes = 1
    Info.bmiHeader.biBitCount = 32
    Info.bmiHeader.biCompression = 0

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, SrcX, SrcY, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
    
    Dim r As Single
    Dim Red As Single, Green As Single, Blue As Single
    Dim Src As Long
    
    
    For H = 0 To DstH - 1
        f = H * DstW
        For B = 0 To DstW - 1
            i = f + B
            r = (Data1(i) / vbWhite) ^ m_Blend
            Src = Data2(i)
            Red = (Src And &HFF) * r
            Green = (Src And &HFF00&) / two8 * r
            Blue = (Src And &HFF0000) / two16 * r
            Data1(i) = RGB(Red, Green, Blue)
            
            'Data1(I) = ShadeColors(Data1(I), Data2(I), m_Blend)
        Next B
    Next H

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0
    
    Erase Data1
    Erase Data2
    
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
errhandl:
    Erase Data1
    Erase Data2
    
    DeleteDC TmpDC
    DeleteDC Sr2DC
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=display,display,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   On Error Resume Next
    BackColor = Display.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    UserControl.BackColor = New_BackColor
    Display.BackColor() = New_BackColor
    Call DrawGraph
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
On Error Resume Next
     Call DrawGraph
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Reset()
On Error Resume Next
   Erase Slices
   ReDim Slices(0)
   UserControl.Cls
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddSlice(Angle As Single, Caption As String, Color As Long) As Long
  On Error GoTo errhandl
  ReDim Preserve Slices(UBound(Slices) + 1)
  With Slices(UBound(Slices))
     .Angle = Angle / 100 * Pi * 2
     .Caption = Caption
     .Color = Color
  End With
  Exit Function
errhandl:
If Err.Number = 6 Then Slices(UBound(Slices)).Angle = 0
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DrawGraph() As Variant
On Error GoTo errhandl
   UserControl.Cls
   'UserControl.Line (0, 0)-(ScaleWidth, ScaleHeight), BackColor, BF
   UserControl.PaintPicture Mask.Picture, 0, 0, ScaleWidth, ScaleHeight
   
   Display.FillStyle = 0
   Display.Cls
   Dim sum As Single, i As Long
   
   Dim Mids() As Single
   Dim theta2 As Single
   sum = 0
   ReDim Mids(UBound(Slices))
   For i = 1 To UBound(Slices)
     theta2 = sum + Slices(i).Angle
     Display.FillColor = Slices(i).Color
     Display.ForeColor = Slices(i).Color
     Pie Display.hdc, 0, 1, W - 2, H - 2, _
        RX + RX * Cos(theta2), RY + RY * Sin(theta2), _
        RX + RX * Cos(sum), RY + RY * Sin(sum)
     Mids(i) = (sum + theta2) / 2
     sum = theta2
   Next i
    If UBound(Slices) = 0 Then
      Display.FillColor = vbRed
      Display.ForeColor = vbRed
      sum = 0
      theta2 = 6.2
      Pie Display.hdc, 0, 1, W - 2, H - 2, _
        RX + RX * Cos(theta2), RY + RY * Sin(theta2), _
        RX + RX * Cos(sum), RY + RY * Sin(sum)
      
    End If
    
    FoxAlphaBlend UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
    Display.hdc, 0, 0 'Check1
    
    For i = 1 To UBound(Slices)
       UserControl.CurrentX = RX + RX * Cos(Mids(i)) / 2 - TextWidth(Slices(i).Caption) / 2
       UserControl.CurrentY = RY + RY * Sin(Mids(i)) / 2 - TextHeight(Slices(i).Caption) / 2
       UserControl.Print Slices(i).Caption
    Next i
    
    
    UserControl.Refresh
    
errhandl:
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Mask,Mask,-1,Picture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  On Error Resume Next
    Set MaskPicture = Mask.Picture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
On Error Resume Next
    Set Mask.Picture = New_MaskPicture
    Call DrawGraph
    PropertyChanged "MaskPicture"
End Property

Private Sub UserControl_Initialize()
On Error Resume Next
   ReDim Slices(0)
   Call UserControl_Resize
   
End Sub

Private Sub UserControl_Paint()


    On Error GoTo Err_Proc
Call DrawGraph
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "PieChart", "UserControl_Paint", Err.Description
    Resume Exit_Proc


End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Display.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Mask.Picture = PropBag.ReadProperty("MaskPicture", Mask.Picture)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
    Call DrawGraph
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Mask.Width = UserControl.ScaleWidth
Mask.Height = UserControl.ScaleHeight
Display.Width = UserControl.ScaleWidth
Display.Height = UserControl.ScaleHeight
W = Display.ScaleWidth
H = Display.ScaleHeight
RX = W / 2
RY = H / 2
Call DrawGraph
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("BackColor", Display.BackColor, &H8000000F)
    Call PropBag.WriteProperty("MaskPicture", Mask.Picture, Nothing)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
End Sub

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
    Call DrawGraph
    PropertyChanged "Font"
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
    Call DrawGraph
    PropertyChanged "ForeColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    Set UserControl.Font = Ambient.Font
    m_Blend = m_def_Blend
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get Blend() As Single
On Error Resume Next
    Blend = m_Blend
End Property

Public Property Let Blend(ByVal New_Blend As Single)
On Error Resume Next
    m_Blend = New_Blend
    Call DrawGraph
    PropertyChanged "Blend"
End Property


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
