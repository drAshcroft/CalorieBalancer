VERSION 5.00
Begin VB.UserControl uGraphSurface 
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   ScaleHeight     =   5955
   ScaleWidth      =   8535
   Begin VB.PictureBox Captions 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   8475
      ScaleHeight     =   330
      ScaleWidth      =   1740
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.PictureBox PicAxis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   -1665
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -210
      Width           =   8295
      Begin VB.PictureBox Canvas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   2400
         ScaleHeight     =   279
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   4695
         Begin VB.Shape Dot 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   2160
            Shape           =   3  'Circle
            Top             =   2760
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Box 
            BorderStyle     =   3  'Dot
            Height          =   1935
            Index           =   0
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Line Line1 
            Index           =   0
            Visible         =   0   'False
            X1              =   72
            X2              =   224
            Y1              =   152
            Y2              =   256
         End
      End
   End
End
Attribute VB_Name = "uGraphSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
'line stuff
Private Type ForceSections

  Dot(1) As Long
  LineI As Long
  SelectedDot As Long
End Type
Private lines() As ForceSections
Private SelectedLine As Long
Private EditLineMode As Boolean

'Box Stuff
Dim BoxX1 As Single, BoxY1 As Single, BoxX2 As Single, BoxY2 As Single

Dim cHdc As Long

'mode stuff
Dim MouseMode As Boolean
Dim NoEvents As Boolean

Dim Months(12) As String
'enumerations
Public Enum GraphTypes
  gLine = 1
  gScatter = 3
  gScatterError = 4
  gIntensity = 5
  gSurface = 6
  gHistogram = 7
  gHistogramGaussian = 8
  gLineError = 9
End Enum

Public Enum GraphProcesses
  gCount = 1
  gProbablity = 2
  gDensity = 3
End Enum

Private Type mDatas
  Filled As Boolean
  Data As Variant
  Color As Long
  Style As GraphTypes
  Caption As String
End Type
Private Datas() As mDatas
Private DrawCenteredText As String

'bounds
Dim MaximaX(1) As Single, MaximaY(1) As Single, MaximaZ(1) As Single
 

'api stuff
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type Size
    CX As Long
    cy As Long
End Type
Const PS_SOLID = 0
Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
'used with fnWeight
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
'used with SetBkMode
Const OPAQUE = 2
Const TRANSPARENT = 1

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

Private Const PIDEG As Double = 1.74532925199433E-02

   Private Const LF_FACESIZE = 32

   Private Type LOGFONT
      lfHeight As Long
      lfWidth As Long
      lfEscapement As Long
      lfOrientation As Long
      lfWeight As Long
      lfItalic As Byte
      lfUnderline As Byte
      lfStrikeOut As Byte
      lfCharSet As Byte
      lfOutPrecision As Byte
      lfClipPrecision As Byte
      lfQuality As Byte
      lfPitchAndFamily As Byte
      lfFaceName As String * LF_FACESIZE
   End Type


Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long



Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long


Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetSysColorBrush Lib "USER32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "USER32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long


' Used to create the metafile
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
' Used for creating the temporary WMF file
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long


'This Enum is needed to set the "Mapping" property for EMF images
Public Enum MMETRIC
        MM_HIMETRIC = 3
        MM_LOMETRIC = 2
        MM_LOENGLISH = 4
        MM_ISOTROPIC = 7
        MM_HIENGLISH = 5
        MM_ANISOTROPIC = 8
        MM_ADLIB = 9
End Enum



Private Type PLACEABLEWMFHEADER
    hdrData(1 To 10) As Integer
    checksum As Integer
End Type


'Event Declarations:
Event SelectLineAdded(Index As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
Event SelectLineChanged(Index As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
Event SelectLineDeleted(Index As Long)
Event SelectBoxSet(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
Event Click() 'MappingInfo=canvas,canvas,-1,Click
Event DblClick() 'MappingInfo=canvas,canvas,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=canvas,canvas,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=canvas,canvas,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=canvas,canvas,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=canvas,canvas,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=canvas,canvas,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=canvas,canvas,-1,MouseUp
Event Paint() 'MappingInfo=canvas,canvas,-1,Paint
'Default Property Values:
Const m_def_AutoScaleXMin = True
Const m_def_AutoScaleXMax = True
Const m_def_MonthAxis = True
Const m_def_CopyGraph = 0
Const m_def_SelectMode = 0
Const m_def_ColorCoded = False

Const m_def_BorderColor = 0
Const m_def_Shape = 0
'Const m_def_AutoScaleX = True
Const m_def_AutoScaleY = True
Const m_def_DisplayXfromZero = False
Const m_def_CaptionName = ""
Const m_def_XAxisName = ""
Const m_def_YAxisName = ""
Const m_def_ZAxisName = ""
Const m_def_OverSizeY = True
'Property Variables:
Dim m_AutoScaleXMin As Boolean
Dim m_AutoScaleXMax As Boolean
Dim m_MonthAxis As Boolean
Dim m_CopyGraph As Variant
Dim m_SelectMode As Long
Dim m_ColorCoded As Boolean
Dim m_ShowLegend As Boolean

Dim m_BorderStyle As Integer
Dim m_BorderColor As Long
Dim m_Shape As Integer
Dim m_BorderWidth As Integer
'Dim m_AutoScaleX As Boolean
Dim m_AutoScaleY As Boolean
Dim m_DisplayXfromZero As Boolean
Dim m_AxisFont As Font
Dim m_CaptionFont As Font
Dim m_NumberFont As Font
Dim m_CaptionName As String
Dim m_XAxisName As String
Dim m_YAxisName As String
Dim m_ZAxisName As String
Dim m_OverSizeY As Boolean
Dim m_OverSizeX As Boolean


Dim mProcess As Boolean
Dim mProcessType As GraphProcesses
Dim mTotalCurveCount As Long
Dim mnBins(1) As Integer


Dim mMetaFile As Boolean
Dim MetaFile As String

Dim SX As Single, WX As Single, Sy As Single, WY As Single, RX As Long, RY As Long
Dim papi As POINTAPI

Dim DigitsToRoundX As Long, DigitsToRoundY As Long
Private GraphObject As PictureBox

Dim OldMapMode As Long

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function GetYMax(Index As Integer)
On Error Resume Next
  GetYMax = MaximaY(Index)
End Function

Public Function GetXMax(Index As Integer)
On Error Resume Next
  GetXMax = MaximaX(Index)
End Function

Public Sub SetXMax(Index As Integer, ByVal Value)
On Error Resume Next
   MaximaX(Index) = Value
End Sub

Public Sub SetYMax(Index As Integer, ByVal Value)
On Error Resume Next
   MaximaY(Index) = Value
End Sub


Private Function MakeMetaFile() As Long
On Error GoTo errhandl
    Dim hMetaDC As Long
    Dim aPt As POINTAPI
    Dim aSize       As Size
    PicAxis.ScaleMode = vbPixels
    Dim X As Long, Y As Long
    X = PicAxis.ScaleX(PicAxis.ScaleWidth, PicAxis.ScaleMode, vbPixels)
    Y = Abs(PicAxis.ScaleY(PicAxis.ScaleHeight, PicAxis.ScaleMode, vbPixels)) + 40
    
    hMetaDC = CreateMetaFile(App.path & "\resources\temp\Temp.wmf")
    ' Set the map mode to MM_ANISOTROPIC
    OldMapMode = SetMapMode(hMetaDC, MM_ANISOTROPIC)
    ' Set the metafile origin as 0, 0
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    ' Set the metafile width and height
    SetWindowExtEx hMetaDC, X, Y, aSize
    ' save the new dimensions
'    SaveDC hMetaDC
    MakeMetaFile = hMetaDC
errhandl:
    
End Function
Private Sub EndMetaFile()
On Error Resume Next
    Dim hMeta As Long
 '   RestoreDC ChDC, True
    ' close it and get the metafile handle
    Call SetMapMode(cHdc, OldMapMode)
    hMeta = CloseMetaFile(cHdc)
    
'    GetObject hMeta, Len(aMetaHdr), aMetaHdr
    ' delete it from memory
    Dim ret As Long
    ret = DeleteMetaFile(hMeta)
    
    Dim i As Long
    PicAxis.ScaleMode = vbPixels
    Dim pmh As PLACEABLEWMFHEADER
    pmh.hdrData(1) = &HCDD7 'id code part1
    pmh.hdrData(2) = &H9AC6 'id code part2
    pmh.hdrData(3) = 0 'always zero
    pmh.hdrData(4) = 0 'left
    pmh.hdrData(5) = 0 'top
    pmh.hdrData(6) = PicAxis.ScaleWidth  'right
    pmh.hdrData(7) = PicAxis.ScaleHeight 'bottom
    pmh.hdrData(8) = 1440 \ Screen.TwipsPerPixelX 'units per inch. 1440 is twips
    pmh.hdrData(9) = 0 'always zero
    pmh.hdrData(10) = 0 'always zero
    pmh.checksum = 0
    For i = 1 To 10
      pmh.checksum = pmh.checksum Xor pmh.hdrData(i)
    Next i
    
    
    'now copy tempmeta.wmf to the actual filename, while prepending the pmh struct
    Dim fi As Long, fo As Long
    Dim lnwmf&
    fi = FreeFile
    Open App.path & "\resources\temp\Temp.wmf" For Binary As #fi
    fo = FreeFile
    Open MetaFile For Binary As #fo
    Put #fo, 1, pmh
    lnwmf& = LOF(fi)
    For i = 1024 To 1 Step -1
       If (lnwmf& Mod i) = 0 Then Exit For
    Next i
    Dim c$
    c$ = Space$(i)
    While Not EOF(fi)
       Get #fi, , c$
       Put #fo, , c$
    Wend
    Close #fi
    Close #fo


End Sub

Public Sub DrawGraph()
On Error GoTo errhandl
  Dim Data() As Double
  Dim TempAutoscale As Boolean
  Dim i As Long, j As Long, X As Double, Y As Double
  Dim ret As Boolean, n As Long

  Dim MinX As Double, MaxX As Double
  If m_AutoScaleXMax Or m_AutoScaleXMin Then
     MinX = 100000
     MaxX = -100000
     For i = 0 To UBound(Datas)
     '  On Error Resume Next
       If TypeName(Datas(i).Data) <> "Empty" Then
       On Error Resume Next
       n = UBound(Datas(i).Data, 2)
       
       If Err.Number = 0 Then
        On Error GoTo errhandl
        For j = 0 To n
          X = Datas(i).Data(0, j)
          If X > MaxX Then MaxX = X
          If X < MinX Then MinX = X
        Next j
       Else
        Err.Clear
        On Error GoTo errhandl
       End If
       End If
     Next i
     If m_AutoScaleXMin Then MaximaX(0) = MinX
     If m_AutoScaleXMax Then MaximaX(1) = MaxX
  End If
  If m_AutoScaleY Then
     MaximaY(0) = 100000
     MaximaY(1) = -100000
     Err.Clear
     For i = 0 To UBound(Datas)
       If TypeName(Datas(i).Data) <> "Empty" Then
       For j = 0 To UBound(Datas(i).Data, 2)
          
          Y = Datas(i).Data(1, j)
          If Err.Number = 0 Then
            If Y > MaximaY(1) Then MaximaY(1) = Y
            If Y < MaximaY(0) Then MaximaY(0) = Y
          Else
            Err.Clear
          End If
          
       Next j
       End If
     Next i
  End If
  If MaximaX(0) = 100000 Then MaximaX(0) = 0
  If MaximaY(0) = 100000 Then MaximaY(0) = 0
  If MaximaX(0) >= MaximaX(1) Then MaximaX(1) = MaximaX(0) + 1
  If MaximaY(0) >= MaximaY(1) Then MaximaY(1) = MaximaY(0) + 1
         
  
  If mMetaFile Then
    'On Error
    cHdc = MakeMetaFile
  Else
    Canvas.Cls
    PicAxis.Cls
    cHdc = PicAxis.hdc
  End If
  ret = DrawAxes
         
  Dim DataColor As Long, GraphType As GraphTypes
  If ret Then
     For i = 0 To UBound(Datas)
       DataColor = Datas(i).Color
       GraphType = Datas(i).Style
       On Error Resume Next
       DrawLines Datas(i).Data, DataColor, GraphType
       Err.Clear
     Next i
  End If
  On Error GoTo errhandl
  Set PicAxis.Font = m_NumberFont
  Set Canvas.Font = m_NumberFont
  Canvas.FontBold = False
  If m_ShowLegend And ret Then
     Call DrawLegend
  End If
  If mMetaFile Then
    Call EndMetaFile
  End If
  
errhandl:
End Sub
Private Sub DrawLegend()
On Error GoTo errhandl
  Dim H As Single, c As Long
  Dim MW As Single, WW As Single
  Dim CX As Single
  Dim i As Long, LCaption  As String
  Dim Sy As Single
  c = 0
  MW = 0
  For i = 0 To UBound(Datas)
     If Datas(i).Caption <> LCaption Then
        c = c + 1
        
     End If
     LCaption = Datas(i).Caption
     WW = Canvas.TextWidth(LCaption)
     If Abs(WW) > Abs(MW) Then MW = WW
  Next i
  
  CX = Canvas.ScaleLeft + Canvas.ScaleWidth * 0.9 - MW
  H = Canvas.TextHeight("!") * c
  Sy = Canvas.ScaleTop + Canvas.ScaleHeight / 10 '2 - H / 2
  Canvas.DrawWidth = 1
  Canvas.Line (CX, Sy)-(Canvas.ScaleLeft + Canvas.ScaleWidth, Sy + H), vbWhite, BF
  Canvas.Line (CX, Sy)-(Canvas.ScaleLeft + Canvas.ScaleWidth, Sy + H), 0, B
  Canvas.CurrentY = Sy
  'Canvas.CurrentX = Canvas.ScaleLeft
  'Canvas.CurrentY = Canvas.ScaleTop + Canvas.ScaleHeight
  On Error Resume Next
  
  For i = 0 To UBound(Datas)
     Canvas.ForeColor = Datas(i).Color
     Canvas.CurrentX = CX ' Canvas.ScaleLeft + Canvas.ScaleWidth - Canvas.TextWidth(Datas(i).Caption)
     If Datas(i).Caption <> LCaption Then
        Canvas.Print Datas(i).Caption
     End If
     LCaption = Datas(i).Caption
  Next i
errhandl:
End Sub
Private Function Log10(X As Double)
On Error Resume Next
  Log10 = Log(X) / Log(10)
  
End Function
Public Function DrawCentered(Text As String)


    On Error GoTo Err_Proc
DrawCenteredText = Text
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "uGraphSurface", "DrawCentered", Err.Description
    Resume Exit_Proc


End Function
Private Function DrawAxes() As Boolean
On Error Resume Next
  Dim i As Long, clr As Long
  Dim n As Long
  Dim Min As Double, Max As Double, MinX As Double, MaxX As Double
  Dim minY As Double, MaxY As Double
  Dim X As Double, Y As Double
        Dim Lside As Double
        Dim Bside As Double
        Dim Rside As Double
        Dim NNy As Double
        Dim RealStepY As Double
        Dim GraphMinY As Double
        Dim RealStepX As Double
        Dim GraphMinX As Double
        Dim Units As String
        Dim NNx As Double
        Dim ii As Long
        Dim ALside As Double
        Dim ABside
        Dim THeight As Double
        Dim RealMax As Double
        Dim MetaX As Single
        Dim MetaY As Single
        Dim NDiv As Long
        
        
        Set GraphObject = PicAxis
        
        If mMetaFile Then
          MetaX = 1
          MetaY = 1
        Else
          MetaX = 1
          MetaY = 1
        End If
        
        
  DrawAxes = True
  On Error Resume Next
        Dim LengthX As Single
        MinX = MaximaX(0)
        MaxX = MaximaX(1)
        minY = MaximaY(0)
        MaxY = MaximaY(1)
        
        If MaxX < MaximaX(0) Then
          MaxX = MaxX + 12
        End If
        If MaxY < minY Then
          DrawAxes = False
          Exit Function
        End If
        
        
        If m_MonthAxis Then
            MaxX = Int(MaxX) + 1
            MinX = Int(MinX)
            NDiv = Int(MaxX) - Int(MinX)
            If NDiv = 0 Then NDiv = 1
        Else
            NDiv = 10
        End If
        
        If m_OverSizeX Then
            LengthX = MaxX - MinX
            MinX = MinX - LengthX / 20
            MaxX = MaxX + LengthX / 20
        End If
        If OverSizeY Then
           
           MaxY = MaxY + 0.5
           minY = minY - 0.5
           
        End If
        
        AClr vbBlack, 1
        RealMax = MaxY
        
        i = 0
        X = Fix(Log10(Abs((MaxX - MinX) / NDiv)))
        If X - 1 < -1 * DigitsToRoundX Then DigitsToRoundX = Abs(X - 1)
        Y = Fix(Log10(Abs((MaxY - minY) / NDiv)))
        If Y - 1 < -1 * DigitsToRoundY Then DigitsToRoundY = Abs(Y - 1)
        
        Dim ScienceX As Boolean, ScienceY As Boolean
        ScienceX = False
        ScienceY = False
        If DigitsToRoundX > 4 Then
          ScienceX = True
        End If
        If DigitsToRoundY > 4 Then
          ScienceY = True
        End If
        
        
        PicAxis.ScaleMode = vbPixels
        UserControl.ScaleMode = vbPixels
        SX = PicAxis.ScaleLeft
        WX = PicAxis.ScaleWidth
        RX = WX
        Sy = 0
        WY = Abs(PicAxis.ScaleHeight)
        RY = WY
        Dim TT As Single
        Set PicAxis.Font = m_AxisFont
        
        Lside = MetaX * PicAxis.TextHeight(m_YAxisName) * 4
        Rside = PicAxis.ScaleWidth
        
        
        Dim cy As Long
        
        Dim FirstFont As Long
        Dim Lf As LOGFONT, hFont As Long
        Dim MidX As Single
        Dim CX As Single
        cy = PicAxis.ScaleHeight / 2 '+ MetaX * PicAxis.TextWidth(m_YAxisName) / 4
        Bside = PicAxis.ScaleHeight - MetaY * PicAxis.TextHeight(m_XAxisName) - 5
        TT = MetaX * PicAxis.TextWidth(m_XAxisName)
        MidX = PicAxis.ScaleWidth / 2 - TT / 2
        'write the y axis name rotated
        Set Captions.Font = m_AxisFont
        CX = PicAxis.TextHeight("!")
        Captions.Height = CX
        Captions.Width = PicAxis.TextWidth(m_YAxisName)
'        Captions.ZOrder
        Captions.Cls
        Captions.CurrentX = 0
        Captions.CurrentY = 0
        Captions.Print m_YAxisName
        Captions.Refresh
        Call FoxRotate(cHdc, CX / 2, cy, Captions.hdc, Captions.Image.handle, _
                &HFF00FF, -90)
     '   PicAxis.Refresh
       ' TextOut cHdc, 0, cy, m_YAxisName, Len(m_YAxisName)
        
       'then write out the x axis name
        Set PicAxis.Font = m_AxisFont
        If mMetaFile Then
           FirstFont = SelectObject(cHdc, CreateMyFont(PicAxis.Font, 0))
        End If
        'DeleteObject SelectObject(ChDC, CreateMyFont(PicAxis.FontSize, 0))
        Call TextOut(cHdc, MidX, CLng(Bside), m_XAxisName, Len(m_XAxisName))
        
        
        Dim ATSide As Double
        If m_CaptionName <> "" Then
           Set PicAxis.Font = m_CaptionFont
           TT = MetaX * PicAxis.TextWidth(m_CaptionName)
           ATSide = MetaY * PicAxis.TextHeight(m_CaptionName)
           If mMetaFile Then
              DeleteObject SelectObject(cHdc, CreateMyFont(PicAxis.Font, 0))
           End If
           Call TextOut(cHdc, PicAxis.ScaleWidth / 2 - TT / 2, 0, m_CaptionName, Len(m_CaptionName))
        Else
           ATSide = 10
        End If
        
        Set PicAxis.Font = m_NumberFont
        If mMetaFile Then
          DeleteObject SelectObject(cHdc, CreateMyFont(PicAxis.Font, 0))
        End If
        Bside = Bside - MetaY * PicAxis.TextHeight("^_") 'subtract off the numbers
        
        Dim GraphMaxY As Double
        
        NNy = 10
        RealStepY = (MaxY - minY) / NNy
        If RealStepY = 0 Then
          DrawAxes = False
          Exit Function
        End If
        GraphMinY = minY / RealStepY
        If GraphMinY - Int(GraphMinY) <> 0 And m_OverSizeY Then
         If minY < 0 Then
          GraphMinY = (Int(Abs(GraphMinY)) + 1) * RealStepY * Sgn(GraphMinY)
         Else
          GraphMinY = (Int(Abs(GraphMinY))) * RealStepY * Sgn(GraphMinY)
         End If
        Else
          GraphMinY = minY
        End If
        
        If GraphMinY + RealStepY * NNy < MaxY And m_OverSizeY Then
          GraphMaxY = GraphMinY + RealStepY * (NNy + 1)
          NNy = NNy + 1
        Else
          GraphMaxY = MaxY
        End If
        
        
        If m_AutoScaleY Then
          minY = GraphMinY
          MaxY = GraphMaxY
        End If
        
        
        NNx = NDiv
        RealStepX = (MaxX - MinX) / NNx
        
        If RealStepX = 0 Then
          DrawAxes = False
          Exit Function
        End If
        THeight = MetaY * PicAxis.TextHeight("^_")
        ABside = Bside - MetaY * PicAxis.TextHeight("^_")
           
           
        Dim junk As String, YFormat As String, XFormat As String, j As Long
        Dim TWidth As Double, s1 As Single, s2 As Single
        s2 = 0
        For i = 0 To NNy
          If ScienceY Then
            YFormat = "00."
            For j = 1 To DigitsToRoundY
              YFormat = YFormat & "0"
            Next j
            YFormat = YFormat & "E+0"
            junk = Format(GraphMinY + RealStepY * i, YFormat) & " "
          Else
            junk = STR$(Round(GraphMinY + RealStepY * i, DigitsToRoundY + 1)) & " "
          End If
          s1 = 1.5 * PicAxis.TextWidth(junk)
          If s1 > s2 Then s2 = s1
        Next i
        
        
        
        
        TWidth = s2
        ALside = Lside + TWidth
        'NNx = (Rside - ALside) / NNx
           
           
        Canvas.Left = ALside
        Canvas.Width = Rside - ALside
        
        Canvas.Top = ATSide
        Canvas.Height = ABside - ATSide
           
           
        Dim HSide As Double
        Dim ConM As Double, ConInt As Double
        
           
        HSide = PicAxis.ScaleHeight
           
        ConM = (GraphMaxY - GraphMinY) / (ATSide - ABside)
        ConInt = GraphMaxY - ConM * ATSide
        
           
        PicAxis.ScaleTop = ConInt
        PicAxis.ScaleHeight = ConM * HSide
        X = Rside / (Rside - ALside) * (MaxX - MinX)
        PicAxis.ScaleLeft = MaxX - X
        PicAxis.ScaleWidth = X + RealStepX / 2
        Canvas.Width = (MaxX - MinX)
        'Canvas.Move Rside, PicAxis.ScaleTop, maxx - minx, GraphMaxY - GraphMinY
           
        Canvas.ScaleLeft = MinX
        Canvas.ScaleWidth = Canvas.Width
        Canvas.ScaleTop = GraphMaxY
        Canvas.ScaleHeight = GraphMinY - GraphMaxY
    
        'set  the scale factors for the api functions
        SX = PicAxis.ScaleLeft
        WX = PicAxis.ScaleWidth
        RX = PicAxis.ScaleX(WX, PicAxis.ScaleMode, vbPixels)
        Sy = PicAxis.ScaleTop + PicAxis.ScaleHeight
        WY = Abs(PicAxis.ScaleHeight)
        RY = Abs(PicAxis.ScaleY(WY, PicAxis.ScaleMode, vbPixels))
        
        Bside = GraphMinY - Abs(PicAxis.ScaleY(THeight, vbPixels, PicAxis.ScaleMode))
        
        Dim XpointWidth As Double, XPointWidth1 As Double, XPointWidth2 As Double, Skip As Long
        If Not ScienceX Then
          XPointWidth1 = 1.5 * PicAxis.TextWidth(STR$(Round(MinX, DigitsToRoundX)))
          XPointWidth2 = 1.5 * PicAxis.TextWidth(STR$(Round(MaxX, DigitsToRoundX)))
        Else
            XFormat = "00."
            For j = 1 To DigitsToRoundX
              XFormat = XFormat & "0"
            Next j
            XFormat = XFormat & "E+0"
        
          XPointWidth1 = 1.5 * PicAxis.TextWidth(STR$(Format(MinX, XFormat)))
          XPointWidth2 = 1.5 * PicAxis.TextWidth(STR$(Format(MaxX, XFormat)))
        End If
        If XPointWidth1 > XPointWidth2 Then XpointWidth = XPointWidth1 Else XpointWidth = XPointWidth2
        
        If XpointWidth >= RealStepX Then
          Skip = Round(XpointWidth * 1.5 / RealStepX + 0.5, 0)
        Else
          Skip = 1
        End If
        Dim Tx As String
        'first do the letters on the graph
        AClr vbBlack, 1
        If mMetaFile Then DeleteObject SelectObject(cHdc, CreateMyFont(PicAxis.Font, 0))
    
        For i = 0 To NNx
           X = RealStepX * i + MinX 'ALside + NNx * i
           If i Mod Skip = 0 Then
             If m_MonthAxis Then
               Tx = Months(WrapMonth(X))
             Else
               Tx = STR$(Round(X, 0))
             End If
             If ScienceX Then
               junk = Format(Tx, XFormat)
             Else
               junk = Tx
             End If
             aTextOut X - 1.5 * PicAxis.TextWidth(Tx) / 2, Bside, junk
           End If
        Next i
        
        Lside = MinX - PicAxis.ScaleX(Abs(TWidth), vbPixels, PicAxis.ScaleMode)
        Dim YPointWidth As Double
        YPointWidth = Abs(MetaY * PicAxis.TextHeight(MaxY))
        If YPointWidth >= RealStepY Then
          Skip = Round(YPointWidth / RealStepY + 0.5, 0)
        Else
          Skip = 1
        End If
           
        THeight = Abs(MetaY * PicAxis.TextHeight("^_") / 2)
        
        For i = 0 To NNy
           X = GraphMinY + RealStepY * i
           If i Mod Skip = 0 Then
             If ScienceY Then
               junk = Format(X, YFormat)
             Else
               junk = STR$(Round(X, DigitsToRoundY + 1)) & " "
             End If
             aTextOut Lside, X + THeight, " " & Round(X, DigitsToRoundY)
           End If
        Next i
       
        'now do the lines on the graph
        If Not mMetaFile Then
           cHdc = Canvas.hdc
           SX = Canvas.ScaleLeft
           WX = Canvas.ScaleWidth
           RX = Canvas.ScaleX(WX, Canvas.ScaleMode, vbPixels)
           Sy = Canvas.ScaleTop + Canvas.ScaleHeight
           WY = Abs(Canvas.ScaleHeight)
           RY = Abs(Canvas.ScaleY(WY, Canvas.ScaleMode, vbPixels))
           Set GraphObject = Canvas
        End If
        
        AClr RGB(200, 200, 200), 1
        For i = 0 To NNx
           X = RealStepX * i + MinX 'ALside + NNx * i
           ApiLine X, GraphMinY, X, GraphMaxY
        Next i
        ApiLine MaxX, GraphMinY, MaxX, GraphMaxY
        For i = 0 To NNy
           X = GraphMinY + RealStepY * i
           ApiLine MinX, X, MaxX, X
        Next i
           
        AClr vbBlack, 2
        ApiLine MinX, GraphMinY, MaxX, GraphMinY
        ApiLine MinX, GraphMinY, MinX, GraphMaxY
        PicAxis.DrawWidth = 1
        
        If mMetaFile Then DeleteObject SelectObject(cHdc, FirstFont)
End Function
Private Function FindU(X As Double, Y As Double, lx As Double, ly As Double, outX As Double, outY As Double, mode As Long) As Double
On Error GoTo errhandl
If mode = 1 Then
  FindU = (outY - ly) / (Y - ly)
ElseIf mode = 0 Then
  FindU = (outX - lx) / (X - lx)
ElseIf mode = 2 Then
  Dim u1 As Double, u2 As Double
  u1 = (outY - ly) / (Y - ly)
  u2 = (outX - lx) / (X - lx)
  If u1 < u2 Then FindU = u1 Else FindU = u2
End If
If FindU < 0 Then FindU = 0
If FindU > 1 Then FindU = 1

Exit Function
errhandl:
FindU = 1
End Function
Private Sub ApiLine(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
On Error Resume Next
  Dim tX1 As Long, tX2 As Long, tY1 As Long, tY2 As Long
   tX1 = (X1 - SX) / WX * RX
   tX2 = (X2 - SX) / WX * RX
   tY1 = RY - (Y1 - Sy) / WY * RY
   tY2 = RY - (Y2 - Sy) / WY * RY
   Call MoveToEx(cHdc, tX1, tY1, papi)
   Call LineTo(cHdc, tX2, tY2)
End Sub
Private Sub ALine(ByVal X1 As Double, ByVal Y1 As Double)
On Error Resume Next
   Dim tX1 As Long, tY1 As Long
   tX1 = (X1 - SX) / WX * RX
   tY1 = RY - (Y1 - Sy) / WY * RY
   Call LineTo(cHdc, tX1, tY1)
End Sub
Private Sub AMoveTo(ByVal X1 As Double, ByVal Y1 As Double)
On Error Resume Next
   Dim tX1 As Long, tY1 As Long
   tX1 = (X1 - SX) / WX * RX
   tY1 = RY - (Y1 - Sy) / WY * RY
   Call MoveToEx(cHdc, tX1, tY1, papi)
End Sub
Private Sub aTextOut(ByVal X1 As Double, ByVal Y1 As Double, junk As String)
On Error Resume Next
   Dim tX1 As Long, tY1 As Long
   tX1 = (X1 - SX) / WX * RX
   tY1 = RY - (Y1 - Sy) / WY * RY
   Call TextOut(cHdc, tX1, tY1, junk, Len(junk))
End Sub
Private Sub aCircle(ByVal X1 As Double, ByVal Y1 As Double, ByVal r As Double)
On Error Resume Next
   Dim tX1 As Long, tY1 As Long
   tX1 = (X1 - SX) / WX * RX
   tY1 = RY - (Y1 - Sy) / WY * RY
   r = Round(r / 2 + 0.5)
   Call Ellipse(cHdc, tX1 - r, tY1 - r, tX1 + r, tY1 + r)

End Sub
Private Sub AClr(clr As Long, Optional W As Long = 1)
On Error Resume Next
   Dim hRPen As Long
   hRPen = CreatePen(PS_SOLID, W, clr)
    'Select our pen into the form's device context and delete the old pen
   DeleteObject SelectObject(cHdc, hRPen)

End Sub
Private Sub DrawLines(Data As Variant, DataColor As Long, GraphType As GraphTypes)
On Error Resume Next
   Dim i As Long
   Dim n As Long
   Dim ColorCoded As Boolean
   Dim clr As Long
   Dim Y As Double, X As Double, lx As Double, ly As Double
    
   
       SX = GraphObject.ScaleLeft
       WX = GraphObject.ScaleWidth
       RX = GraphObject.ScaleX(WX, GraphObject.ScaleMode, vbPixels)
       Sy = GraphObject.ScaleTop + GraphObject.ScaleHeight
       WY = Abs(GraphObject.ScaleHeight)
       RY = Abs(GraphObject.ScaleY(WY, GraphObject.ScaleMode, vbPixels))
       n = UBound(Data, 2)
       If GraphType = gLineError Or GraphType = gScatterError Then
          Dim Barwidth As Double
          Barwidth = GraphObject.ScaleX(5, vbPixels, GraphObject.ScaleMode)
          AClr vbBlack, 1
          For i = 0 To n
            Y = Data(1, i)
            X = Data(0, i)
            If (Y - MaximaY(0)) * (MaximaY(1) - Y) >= 0 And (X - MaximaX(0)) * (MaximaX(1) - X) >= 0 Then
            
             ApiLine X, Y - Data(3, i), X, Y + Data(3, i) 'black
             ApiLine X - Barwidth, Y - Data(3, i), X + Barwidth, Y - Data(3, i) 'black
             ApiLine X - Barwidth, Y + Data(3, i), X + Barwidth, Y + Data(3, i) 'black
            End If
          Next i
       End If
       
       If GraphType = gLine Or GraphType = gLineError Then
          Call AMoveTo(Data(0, 0), Data(1, 0))
          ColorCoded = m_ColorCoded
          Dim InB As Boolean, LInB As Boolean
          
          LInB = True
          If Not ColorCoded Then
            AClr DataColor, 1
          End If
          For i = 0 To n
            Y = Data(1, i)
            X = Data(0, i)
            
            If (Y - MaximaY(0)) * (MaximaY(1) - Y) >= 0 And (X - MaximaX(0)) * (MaximaX(1) - X) >= 0 Then
              InB = True
            Else
              InB = False
            End If
            If i = 0 Then LInB = InB
            
            If ColorCoded Then
              clr = Data(2, i)
              AClr clr, 1
            End If
            Dim outX As Double, outY As Double
            
            If InB = False Or LInB = False Then
              Dim Um As Double
              If InB <> LInB Then
               If InB = False Then
                outX = vbWhite
                If X < MaximaX(0) Then
                  outX = MaximaX(0)
                ElseIf X > MaximaX(1) Then
                  outX = MaximaX(1)
                End If
                outY = vbWhite
                If Y < MaximaY(0) Then
                  outY = MaximaY(0)
                ElseIf Y > MaximaY(1) Then
                  outY = MaximaY(1)
                End If
                If outX = vbWhite And outY <> vbWhite Then
                  Um = FindU(X, Y, lx, ly, 0, outY, 1)
                ElseIf outX <> vbWhite And outY = vbWhite Then
                  Um = FindU(X, Y, lx, ly, outX, 0, 0)
                Else
                  Um = FindU(X, Y, lx, ly, outX, outY, 2)
                End If
                ApiLine lx, ly, lx + Um * (X - lx), ly + Um * (Y - ly) ', Clr
              ElseIf LInB = False Then
                outX = vbWhite
                If lx < MaximaX(0) Then
                  outX = MaximaX(0)
                ElseIf lx > MaximaX(1) Then
                  outX = MaximaX(1)
                End If
                outY = vbWhite
                If ly < MaximaY(0) Then
                  outY = MaximaY(0)
                ElseIf ly > MaximaY(1) Then
                  outY = MaximaY(1)
                End If
                If outX = vbWhite And outY <> vbWhite Then
                  Um = FindU(X, Y, lx, ly, 0, outY, 1)
                ElseIf outX <> vbWhite And outY = vbWhite Then
                  Um = FindU(X, Y, lx, ly, outX, 0, 0)
                Else
                  Um = FindU(X, Y, lx, ly, outX, outY, 2)
                End If
                ApiLine lx + Um * (X - lx), ly + Um * (Y - ly), X, Y ' Clr
              End If
             End If
            Else
              ALine Data(0, i), Data(1, i)  ' Clr
            End If
            LInB = InB
            lx = X
            ly = Y
          Next i
       End If
       If GraphType = gScatter Or GraphType = gScatterError Then
          Dim Radius As Double
          Radius = GraphObject.ScaleX(2, vbPixels, GraphObject.ScaleMode)
          ColorCoded = m_ColorCoded
          If Not ColorCoded Then
            AClr DataColor, 1
          End If
          For i = 0 To n
            Y = Data(1, i)
            X = Data(0, i)
            If (Y - MaximaY(0)) * (MaximaY(1) - Y) >= 0 And (X - MaximaX(0)) * (MaximaX(1) - X) >= 0 Then
              If ColorCoded Then
                clr = Data(2, i)
                AClr clr, 1
              End If
              If clr <> vbWhite Then
                aCircle X, Y, 2
              End If
            End If
          Next i
       End If
       
 
End Sub

Private Function CreateMyFont(Font As StdFont, nDegrees As Long) As Long
On Error Resume Next
    'Create a specified font
    Dim Weight As Long
    If Font.Bold Then
       Weight = FW_BOLD
    Else
       Weight = FW_NORMAL
    End If
    CreateMyFont = CreateFont(-MulDiv(Font.Size, GetDeviceCaps(cHdc, LOGPIXELSY), 72), 0, nDegrees * 10, 0, Weight, Font.Italic, Font.Underline, Font.Strikethrough, Font.Charset, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, Font.Name)
End Function

Private Sub picaxis_Resize()
On Error Resume Next
Call DrawGraph
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
SelectedLine = -1
ReDim Datas(0)
ReDim lines(0)

Dim d(1, 1) As Double
d(0, 0) = 0
d(1, 0) = 0
d(0, 1) = 10
d(1, 1) = 10
AddLine "", gScatter, vbWhite, d
Call DrawGraph

Months(1) = "Jan"
Months(2) = "Feb"
Months(3) = "Mar"
Months(4) = "Apr"
Months(5) = "May"
Months(6) = "Jun"
Months(7) = "Jul"
Months(8) = "Aug"
Months(9) = "Sep"
Months(10) = "Oct"
Months(11) = "Nov"
Months(12) = "Dec"

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
PicAxis.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
On Error Resume Next
    BackColor = Canvas.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    Canvas.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,BackColor
Public Property Get ShowLegend() As Boolean
On Error Resume Next
    ShowLegend = m_ShowLegend
End Property

Public Property Let ShowLegend(ByVal New_ShowLegend As Boolean)
On Error Resume Next
    m_ShowLegend = New_ShowLegend
   ' Call DrawGraph
    PropertyChanged "ShowLegend"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
On Error Resume Next
    Enabled = PicAxis.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next
    PicAxis.Enabled() = New_Enabled
  '  Canvas.Visible = New_Enabled
    
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = Canvas.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set Canvas.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
On Error Resume Next
    Canvas.Refresh
End Sub

Private Sub canvas_Click()
On Error Resume Next
    RaiseEvent Click
End Sub

Private Sub canvas_DblClick()
On Error Resume Next
    RaiseEvent DblClick
End Sub

Private Sub canvas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub canvas_KeyPress(KeyAscii As Integer)
On Error Resume Next
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub canvas_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
    If SelectedLine <> -1 Then
       On Error Resume Next
       With lines(SelectedLine)
          Unload Line1(.LineI)
          Unload Dot(.Dot(0))
          Unload Dot(.Dot(1))
          .LineI = -1
          .Dot(0) = -1
          .Dot(1) = -1
          .SelectedDot = 0
       End With
       RaiseEvent SelectLineDeleted(SelectedLine)
       If SelectedLine > 1 Then SelectedLine = SelectedLine - 1
    End If
End If
    
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
   MouseMode = True
   
   Dim i As Long, j As Long
   Dim X1 As Long, Y1 As Long
   If m_SelectMode = 1 Then
     'first check if they have clicked on a previous line
     Dim Pix As Single, dist As Single
     Dim LineT As Line
     Pix = Dot(0).Width ^ 2 * 4
     dist = Dot(0).Width
    
     For i = 1 To UBound(lines)
       With lines(i)
         If .LineI <> -1 Then
          Set LineT = Line1(.LineI)
        
          If LineMouseEvent(LineT, X, Y, 8) Then
             LineT.BorderColor = vbBlue
             SelectedLine = i
             .SelectedDot = -1
             EditLineMode = True
             If (LineT.X1 - X) ^ 2 + (LineT.Y1 - Y) ^ 2 < Pix Then
                .SelectedDot = 1
             End If
             If (LineT.X2 - X) ^ 2 + (LineT.Y2 - Y) ^ 2 < Pix Then
                .SelectedDot = 0
             End If
             Set LineT = Nothing
             GoTo Exitsub
          End If
          Set LineT = Nothing
         End If
       End With
     Next i
     
     'next add a new line
     SelectedLine = -1
     For i = 1 To UBound(lines)
        If lines(i).LineI = -1 Then
           SelectedLine = i
           Exit For
        End If
     Next i
     If SelectedLine = -1 Then
        i = UBound(lines) + 1
        ReDim Preserve lines(i)
        SelectedLine = i
     End If
     With lines(SelectedLine)
       On Error Resume Next
       Err.Clear
       For j = 0 To Line1.UBound + 1
         Load Line1(j)
         If Err.Number = 0 Then
           .LineI = j
           Exit For
         Else
           Err.Clear
         End If
       Next j
     
       Line1(.LineI).X1 = X
       Line1(.LineI).X2 = X
       Line1(.LineI).Y1 = Y
       Line1(.LineI).Y2 = Y
       Line1(.LineI).Visible = True
       For j = 0 To Dot.UBound + 1
         Load Dot(j)
         Dot(j).ZOrder
         If Err.Number = 0 Then
           .Dot(0) = j
           Exit For
         Else
           Err.Clear
         End If
       Next j
       Dot(.Dot(0)).Move X - Dot(.Dot(0)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(0)).Height / 2
       Dot(.Dot(0)).Visible = True
       For j = 0 To Dot.UBound + 1
         Load Dot(j)
         Dot(j).ZOrder
         If Err.Number = 0 Then
           .Dot(1) = j
           Exit For
         Else
           Err.Clear
         End If
       Next j
       Dot(.Dot(1)).Move X - Dot(.Dot(1)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(1)).Height / 2
       Dot(.Dot(1)).Visible = True
       .SelectedDot = 1
     End With
     
    ElseIf m_SelectMode = 2 Then
       Box(0).Visible = True
       Box(0).Move X, Y, 1, 1
       BoxX1 = X
       BoxY1 = Y
       
    End If
    
Exitsub:
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Dim i As Long
   If MouseMode And m_SelectMode = 1 Then
     If SelectedLine > UBound(lines) Then Exit Sub
     With lines(SelectedLine)
       If .SelectedDot > -1 Then
         If .SelectedDot = 0 Then
           Line1(.LineI).X2 = X
           Line1(.LineI).Y2 = Y
         Else
           Line1(.LineI).X1 = X
           Line1(.LineI).Y1 = Y
         End If
         Dot(.Dot(.SelectedDot)).Move X - Dot(.Dot(.SelectedDot)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(.SelectedDot)).Height / 2
       End If
     End With
   End If
   If MouseMode And m_SelectMode = 2 Then
     
      BoxX2 = X
      BoxY2 = Y
      If BoxX2 > BoxX1 Then
         Box(0).Width = BoxX2 - Box(0).Left
      Else
         Box(0).Left = BoxX2
         Box(0).Width = BoxX1 - BoxX2
      End If
      If Canvas.ScaleHeight > 0 Then
        If BoxY2 > BoxY1 Then
           Box(0).Height = Abs(BoxY2 - BoxY1)
        Else
           Box(0).Top = BoxY2
           Box(0).Height = Abs(BoxY1 - BoxY2)
        End If
      Else
        If BoxY2 < BoxY1 Then
           Box(0).Height = Abs(BoxY2 - BoxY1)
        Else
           Box(0).Top = BoxY2
           Box(0).Height = Abs(BoxY1 - BoxY2)
        End If
      End If
   End If

    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If m_SelectMode = 1 Then
  If SelectedLine > UBound(lines) Then Exit Sub
  With lines(SelectedLine)
    Line1(.LineI).BorderColor = vbBlack
    If EditLineMode Then
      If Not NoEvents Then RaiseEvent SelectLineChanged(SelectedLine, Line1(.LineI).X1, Line1(.LineI).Y1, Line1(.LineI).X2, Line1(.LineI).Y2)
    Else
      If Not NoEvents Then RaiseEvent SelectLineAdded(SelectedLine, Line1(.LineI).X1, Line1(.LineI).Y1, Line1(.LineI).X2, Line1(.LineI).Y2)
    End If
    EditLineMode = False
  End With
End If
If m_SelectMode = 2 Then
   If Not NoEvents Then RaiseEvent SelectBoxSet(BoxX1, BoxY1, BoxX2, BoxY2)
   Box(0).Visible = False
End If
   MouseMode = False
   If Not NoEvents Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
On Error Resume Next
    AutoRedraw = Canvas.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
On Error Resume Next
    Canvas.AutoRedraw() = New_AutoRedraw
    PicAxis.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property


'The Underscore following "Circle" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Circle
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
On Error Resume Next
    Canvas.Circle (X, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
On Error Resume Next
    Canvas.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
On Error Resume Next
    CurrentX = Canvas.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
On Error Resume Next
    Canvas.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
On Error Resume Next
    CurrentY = Canvas.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
On Error Resume Next
    Canvas.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
On Error Resume Next
    DrawMode = Canvas.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
On Error Resume Next
    Canvas.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
On Error Resume Next
    DrawStyle = Canvas.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
On Error Resume Next
    Canvas.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
On Error Resume Next
    DrawWidth = Canvas.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
On Error Resume Next
    Canvas.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
On Error Resume Next
    hdc = Canvas.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
On Error Resume Next
    Set Image = Canvas.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Line
Public Sub Line(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
Attribute Line.VB_Description = "Draws lines and rectangles on an object."
On Error Resume Next
    Canvas.Line (X1, Y1)-(X2, Y2), Color
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
On Error Resume Next
    MousePointer = Canvas.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
On Error Resume Next
    Canvas.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
On Error Resume Next
    Set MouseIcon = Canvas.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
On Error Resume Next
    Set Canvas.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub canvas_Paint()
On Error Resume Next
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
On Error Resume Next
    Set Picture = Canvas.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
On Error Resume Next
    Set Canvas.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
On Error Resume Next
    Point = Canvas.Point(X, Y)
End Function

'The Underscore following "PSet" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
On Error Resume Next
    Canvas.PSet Step(X, Y), Color
End Sub

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
On Error Resume Next
    Canvas.Scale (X1, Y1)-(X2, Y2)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
On Error Resume Next
    ScaleHeight = Canvas.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
On Error Resume Next
    Canvas.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
On Error Resume Next
    ScaleLeft = Canvas.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
On Error Resume Next
    Canvas.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
On Error Resume Next
    ScaleMode = Canvas.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
On Error Resume Next
    Canvas.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
On Error Resume Next
    ScaleTop = Canvas.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
On Error Resume Next
    Canvas.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=canvas,canvas,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
On Error Resume Next
    ScaleWidth = Canvas.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
On Error Resume Next
    Canvas.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Box(0),Box,0,Shape
'Public Property Get Shape() As Integer
'    Shape = Box(0).Shape
'End Property
'
'Public Property Let Shape(ByVal New_Shape As Integer)
'    Box(0).Shape() = New_Shape
'    PropertyChanged "Shape"
'End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Canvas.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Canvas.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Canvas.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Box(0).BorderStyle = PropBag.ReadProperty("BorderStyle", 2)
    Canvas.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    Line1(0).BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    Line1(0).BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    
    Canvas.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    Canvas.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    Canvas.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    Canvas.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Canvas.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 5595)
    Canvas.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    Canvas.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 8235)
    Canvas.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    Canvas.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
    
    
    Box(0).Shape = PropBag.ReadProperty("Shape", 0)

    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_Shape = PropBag.ReadProperty("Shape", m_def_Shape)
'    m_AutoScaleX = PropBag.ReadProperty("AutoScaleX", m_def_AutoScaleX)
    m_AutoScaleY = PropBag.ReadProperty("AutoScaleY", m_def_AutoScaleY)
    m_DisplayXfromZero = PropBag.ReadProperty("DisplayXfromZero", m_def_DisplayXfromZero)
    Set m_AxisFont = PropBag.ReadProperty("AxisFont", Ambient.Font)
    Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    Set m_NumberFont = PropBag.ReadProperty("NumberFont", Ambient.Font)
    m_CaptionName = PropBag.ReadProperty("CaptionName", m_def_CaptionName)
    m_XAxisName = PropBag.ReadProperty("XAxisName", m_def_XAxisName)
    m_YAxisName = PropBag.ReadProperty("YAxisName", m_def_YAxisName)
    m_ZAxisName = PropBag.ReadProperty("ZAxisName", m_def_ZAxisName)
    m_OverSizeY = PropBag.ReadProperty("OverSizeY", m_def_OverSizeY)
    m_ColorCoded = PropBag.ReadProperty("ColorCoded", m_def_ColorCoded)
    m_SelectMode = PropBag.ReadProperty("SelectMode", m_def_SelectMode)
    m_CopyGraph = PropBag.ReadProperty("CopyGraph", m_def_CopyGraph)
    m_MonthAxis = PropBag.ReadProperty("MonthAxis", m_def_MonthAxis)
    m_AutoScaleXMin = PropBag.ReadProperty("AutoScaleXMin", m_def_AutoScaleXMin)
    m_AutoScaleXMax = PropBag.ReadProperty("AutoScaleXMax", m_def_AutoScaleXMax)
    m_ShowLegend = PropBag.ReadProperty("Showlegend", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("ShowLegend", m_ShowLegend, False)
    Call PropBag.WriteProperty("BackColor", Canvas.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", Canvas.Enabled, True)
    Call PropBag.WriteProperty("Font", Canvas.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", Box(0).BorderStyle, 2)
    Call PropBag.WriteProperty("AutoRedraw", Canvas.AutoRedraw, True)
    Call PropBag.WriteProperty("BorderColor", Line1(0).BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderWidth", Line1(0).BorderWidth, 1)
    Call PropBag.WriteProperty("CurrentX", Canvas.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", Canvas.CurrentY, 0)
    Call PropBag.WriteProperty("DrawMode", Canvas.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", Canvas.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", Canvas.DrawWidth, 1)
    Call PropBag.WriteProperty("FillColor", Dot(0).FillColor, &HC000&)
    Call PropBag.WriteProperty("FillStyle", Dot(0).FillStyle, 0)
    Call PropBag.WriteProperty("MousePointer", Canvas.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", Canvas.ScaleHeight, 5595)
    Call PropBag.WriteProperty("ScaleLeft", Canvas.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", Canvas.ScaleMode, 3)
    Call PropBag.WriteProperty("ScaleTop", Canvas.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", Canvas.ScaleWidth, 8235)
    Call PropBag.WriteProperty("Shape", Box(0).Shape, 0)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("Shape", m_Shape, m_def_Shape)
'    Call PropBag.WriteProperty("AutoScaleX", m_AutoScaleX, m_def_AutoScaleX)
    Call PropBag.WriteProperty("AutoScaleY", m_AutoScaleY, m_def_AutoScaleY)
    Call PropBag.WriteProperty("DisplayXfromZero", m_DisplayXfromZero, m_def_DisplayXfromZero)
    Call PropBag.WriteProperty("AxisFont", m_AxisFont, Ambient.Font)
    Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
    Call PropBag.WriteProperty("NumberFont", m_NumberFont, Ambient.Font)
    Call PropBag.WriteProperty("CaptionName", m_CaptionName, m_def_CaptionName)
    Call PropBag.WriteProperty("XAxisName", m_XAxisName, m_def_XAxisName)
    Call PropBag.WriteProperty("YAxisName", m_YAxisName, m_def_YAxisName)
    Call PropBag.WriteProperty("ZAxisName", m_ZAxisName, m_def_ZAxisName)
    Call PropBag.WriteProperty("OverSizeY", m_OverSizeY, m_def_OverSizeY)
    Call PropBag.WriteProperty("ColorCoded", m_ColorCoded, m_def_ColorCoded)
    Call PropBag.WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
    Call PropBag.WriteProperty("CopyGraph", m_CopyGraph, m_def_CopyGraph)
    Call PropBag.WriteProperty("MonthAxis", m_MonthAxis, m_def_MonthAxis)
    Call PropBag.WriteProperty("AutoScaleXMin", m_AutoScaleXMin, m_def_AutoScaleXMin)
    Call PropBag.WriteProperty("AutoScaleXMax", m_AutoScaleXMax, m_def_AutoScaleXMax)
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
On Error Resume Next
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
On Error Resume Next
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
On Error Resume Next
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
On Error Resume Next
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Shape() As Integer
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
On Error Resume Next
    Shape = m_Shape
End Property

Public Property Let Shape(ByVal New_Shape As Integer)
On Error Resume Next
    m_Shape = New_Shape
    PropertyChanged "Shape"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
On Error Resume Next
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
On Error Resume Next
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get AutoScaleY() As Boolean
On Error Resume Next
    AutoScaleY = m_AutoScaleY
End Property

Public Property Let AutoScaleY(ByVal New_AutoScaleY As Boolean)
On Error Resume Next
    m_AutoScaleY = New_AutoScaleY
    PropertyChanged "AutoScaleY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get DisplayXfromZero() As Boolean
On Error Resume Next
    DisplayXfromZero = m_DisplayXfromZero
End Property

Public Property Let DisplayXfromZero(ByVal New_DisplayXfromZero As Boolean)
On Error Resume Next
    m_DisplayXfromZero = New_DisplayXfromZero
    PropertyChanged "DisplayXfromZero"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get AxisFont() As Font
On Error Resume Next
    Set AxisFont = m_AxisFont
End Property

Public Property Set AxisFont(ByVal New_AxisFont As Font)
On Error Resume Next
    Set m_AxisFont = New_AxisFont
    PropertyChanged "AxisFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get CaptionFont() As Font
On Error Resume Next
    Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)
On Error Resume Next
    Set m_CaptionFont = New_CaptionFont
    PropertyChanged "CaptionFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get NumberFont() As Font
On Error Resume Next
    Set NumberFont = m_NumberFont
End Property

Public Property Set NumberFont(ByVal New_NumberFont As Font)
On Error Resume Next
    Set m_NumberFont = New_NumberFont
    PropertyChanged "NumberFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get CaptionName() As String
On Error Resume Next
    CaptionName = m_CaptionName
End Property

Public Property Let CaptionName(ByVal New_CaptionName As String)
On Error Resume Next
    m_CaptionName = New_CaptionName
    PropertyChanged "CaptionName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get XAxisName() As String
On Error Resume Next
    XAxisName = m_XAxisName
End Property

Public Property Let XAxisName(ByVal New_XAxisName As String)
On Error Resume Next
    m_XAxisName = New_XAxisName
    PropertyChanged "XAxisName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get YAxisName() As String
On Error Resume Next
    YAxisName = m_YAxisName
End Property

Public Property Let YAxisName(ByVal New_YAxisName As String)
On Error Resume Next
    m_YAxisName = New_YAxisName & "  "
    PropertyChanged "YAxisName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ZAxisName() As String
On Error Resume Next
    ZAxisName = m_ZAxisName
End Property

Public Property Let ZAxisName(ByVal New_ZAxisName As String)
On Error Resume Next
    m_ZAxisName = New_ZAxisName
    PropertyChanged "ZAxisName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get OverSizeY() As Boolean
On Error Resume Next
    OverSizeY = m_OverSizeY
End Property

Public Property Let OverSizeY(ByVal New_OverSizeY As Boolean)
On Error Resume Next
    m_OverSizeY = New_OverSizeY
    PropertyChanged "OverSizeY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get OverSizeX() As Boolean
On Error Resume Next
    OverSizeX = m_OverSizeX
End Property

Public Property Let OverSizeX(ByVal New_OverSizeY As Boolean)
On Error Resume Next
    m_OverSizeX = New_OverSizeY
    PropertyChanged "OverSizeX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddLine(Caption As String, LineType As GraphTypes, Color As Long, Data) As Long
On Error Resume Next
Dim i As Long
For i = 0 To UBound(Datas)
  With Datas(i)
   If .Filled = False Then
      .Style = LineType
      .Color = Color
      .Data = Data
      .Filled = True
      .Caption = Caption
      Exit For
   End If
  End With
Next i
If i > UBound(Datas) Then
  ReDim Preserve Datas(i)
  With Datas(i)
      .Style = LineType
      .Color = Color
      .Data = Data
      .Filled = True
      .Caption = Caption
  End With
End If

AddLine = i
End Function
Public Function AddText(Caption As String, X As Single, Y As Single)
On Error Resume Next
X = Canvas.ScaleLeft + Canvas.ScaleWidth * X
Y = Canvas.ScaleTop + Canvas.ScaleHeight * Y
Canvas.CurrentX = X
Canvas.CurrentY = Y
Canvas.Print Caption

End Function
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    m_BorderColor = m_def_BorderColor
    m_Shape = m_def_Shape

    m_AutoScaleY = m_def_AutoScaleY
    m_DisplayXfromZero = m_def_DisplayXfromZero
    Set m_AxisFont = Ambient.Font
    Set m_CaptionFont = Ambient.Font
    Set m_NumberFont = Ambient.Font
    m_CaptionName = m_def_CaptionName
    m_XAxisName = m_def_XAxisName
    m_YAxisName = m_def_YAxisName
    m_ZAxisName = m_def_ZAxisName
    m_OverSizeY = m_def_OverSizeY
    m_ColorCoded = m_def_ColorCoded
    m_SelectMode = m_def_SelectMode
    m_CopyGraph = m_def_CopyGraph
    m_MonthAxis = m_def_MonthAxis
    m_AutoScaleXMin = m_def_AutoScaleXMin
    m_AutoScaleXMax = m_def_AutoScaleXMax
    m_ShowLegend = False
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get ColorCoded() As Boolean
On Error Resume Next
    ColorCoded = m_ColorCoded
End Property

Public Property Let ColorCoded(ByVal New_ColorCoded As Boolean)
On Error Resume Next
    m_ColorCoded = New_ColorCoded
    PropertyChanged "ColorCoded"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Reset() As Variant
m_MonthAxis = True
On Error Resume Next
m_AutoScaleXMin = True
m_AutoScaleXMax = True
m_AutoScaleY = True

Dim i As Long
For i = 0 To UBound(Datas)
  Datas(i).Data = ""
Next i
Erase Datas
ReDim Datas(0)

For i = 0 To Line1.UBound
  Unload Line1(i)
Next i
For i = 0 To Dot.UBound
  Unload Dot(i)
Next i
Erase lines
ReDim lines(0)
DrawCenteredText = ""
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SelectMode() As Long
On Error Resume Next
    SelectMode = m_SelectMode
End Property

Public Property Let SelectMode(ByVal New_SelectMode As Long)
On Error Resume Next
    m_SelectMode = New_SelectMode
    PropertyChanged "SelectMode"
End Property

Private Function LineMouseEvent(LineName As Line, ByVal X As Single, ByVal Y As Single, Optional LineWidth = 0, Optional LineHeight = 0) As Boolean
  On Error GoTo ExitF
  Dim XMin&, Xmax&, YMin&, YMax&, Gradient As Single, HalfBordWid As Single
  Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
  X = ChangeToPix(0, X)
  Y = ChangeToPix(1, Y)
  X1 = ChangeToPix(0, LineName.X1)
  X2 = ChangeToPix(0, LineName.X2)
  Y1 = ChangeToPix(1, LineName.Y1)
  Y2 = ChangeToPix(1, LineName.Y2)
  If X1 < X2 Then
     XMin = X1
     Xmax = X2
  Else
     XMin = X2
     Xmax = X1
  End If
  If Y1 < Y2 Then
     YMin = Y1
     YMax = Y2
  Else
     YMin = Y2
     YMax = Y1
  End If
  HalfBordWid = (LineName.BorderWidth + LineWidth) / 2

  If X >= XMin - HalfBordWid And X <= Xmax + HalfBordWid And Y >= YMin - HalfBordWid And Y <= YMax + HalfBordWid Then
       'calculate the line vector equation and check the Y values
       If X2 - X1 = 0 Then
          If Abs(X - X1) < LineWidth Then LineMouseEvent = True
       Else
          Gradient = (Y2 - Y1) / (X2 - X1)
          'Line Equation is: Y - Y2 = Gradient * (X - X2)
           LineMouseEvent = CBool(Abs(Gradient * (X - X2) - (Y - Y2)) < (LineName.BorderWidth + LineWidth))
       End If
  End If
ExitF:
 Exit Function
errhandl:
 MsgBox Err.Description, vbCritical, "Force Chart Editor"
 Err.Clear

End Function
Public Function ChangeToPix(axis As Long, ByVal X As Single) As Single
On Error Resume Next
   If axis = 0 Then
      ChangeToPix = Canvas.ScaleX(X - Canvas.ScaleLeft, Canvas.ScaleMode, vbPixels)
   Else
      
      ChangeToPix = Abs(Canvas.ScaleY(X - Canvas.ScaleTop, Canvas.ScaleMode, vbPixels))
   End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ClearSelectLines() As Variant
  On Error Resume Next
  Dim i As Long
  For i = 0 To UBound(lines)
    With lines(i)
       Unload Line1(.LineI)
       Unload Dot(.Dot(0))
       Unload Dot(.Dot(1))
       .LineI = 0
       .Dot(0) = 0
       .Dot(1) = 0
       .SelectedDot = 0
    End With
  Next i
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddSelectLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Long
On Error Resume Next
Dim T
T = m_SelectMode
NoEvents = True
m_SelectMode = 1
Call canvas_MouseDown(0, 0, X1, Y1)
AddSelectLine = SelectedLine
Call canvas_MouseMove(0, 0, X2, Y2)
Call canvas_MouseUp(0, 0, X2, Y2)
m_SelectMode = T
NoEvents = False
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function CopyData() As Variant
On Error GoTo errhandl
  Dim TMousepointer
  TMousepointer = Canvas.MousePointer
  Canvas.MousePointer = 11

  Clipboard.Clear
  Dim junk() As Variant
  Dim i As Long, j As Long, k As Long
  Dim Max As Long
  'find the max number of datapoint
 
  Max = 0
  For i = 0 To UBound(Datas)
     With Datas(i)
       If Not VarType(.Data) = vbEmpty Then
        j = UBound(.Data, 2)
        If j > Max Then Max = j
       End If
     End With
  Next i
  If Max = 0 Then Exit Function
  'make room for all the data
  ReDim junk(2 * UBound(Datas) + 1, Max)
  'now put it on a grid two rows per section
  For i = 0 To UBound(Datas)
     With Datas(i)
       For j = 0 To UBound(.Data, 2)
          junk(i * 2, j) = .Data(0, j)
          junk(i * 2 + 1, j) = .Data(1, j)
       Next j
     End With
  Next i
  'put it into text
  Dim junkString As String
  junkString = ""
  For j = 0 To UBound(junk, 2)
    For i = 0 To UBound(junk, 1) - 1
       junkString = junkString & junk(i, j) & vbTab
    Next
    junkString = junkString & junk(i, j) & vbCrLf
  Next
  Clipboard.SetText junkString
  Canvas.MousePointer = TMousepointer
  Exit Function
errhandl:
  MsgBox Err.Description
  Canvas.MousePointer = TMousepointer
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function CopyGraph() As Variant
On Error Resume Next
   Clipboard.Clear
   On Error Resume Next
   mMetaFile = True

   UserControl.ScaleMode = vbInches
   PicAxis.Width = 5
   PicAxis.Height = 3
   UserControl.ScaleMode = vbTwips
   MetaFile = App.path & "\resources\temp\Temp2.wmf"
   Call DrawGraph
   mMetaFile = False
   Call DrawGraph


   Clipboard.SetData LoadPicture(MetaFile), vbCFMetafile
   
   
   Dim Tx As Single, Ty As Single, TW As Single, th As Single
   Dim Ts
   Tx = PicAxis.ScaleLeft
   Ty = PicAxis.ScaleTop
   TW = PicAxis.ScaleWidth
   th = PicAxis.ScaleHeight
   Ts = PicAxis.ScaleMode
   PicAxis.ScaleMode = vbPixels
   PicAxis.PaintPicture Canvas.Image, Canvas.Left, Canvas.Top
   Canvas.Visible = False
   Dim X As Long, Y As Long, Y2 As Long
   X = Canvas.Left + Canvas.Width
   Y = Canvas.Top
   Y2 = Canvas.Top + Canvas.Height
   
   PicAxis.Line (X, Y)-(X, Y2), RGB(200, 200, 200)
   PicAxis.Refresh
   
   Clipboard.SetData PicAxis.Image, vbCFBitmap
   
   Canvas.Visible = True
   UserControl_Resize
   Call DrawGraph
   
End Function
Public Sub SaveBitmap(Filename As String)
On Error Resume Next
   PicAxis.AutoRedraw = True
   'Set PicAxis.Picture = PicAxis.Image
  
   PicAxis.ScaleMode = vbPixels
   PicAxis.PaintPicture Canvas.Image, Canvas.Left, Canvas.Top
   Canvas.Visible = False
   Dim X As Long, Y As Long, Y2 As Long
   X = Canvas.Left + Canvas.Width
   Y = Canvas.Top
   Y2 = Canvas.Top + Canvas.Height
   
   PicAxis.Line (X, Y)-(X, Y2), RGB(200, 200, 200)
   PicAxis.ScaleMode = vbPixels
   If DrawCenteredText <> "" Then
      Dim Parts() As String, i As Long
      Parts = Split(DrawCenteredText, vbCrLf)
      PicAxis.CurrentY = PicAxis.ScaleHeight / 2 - PicAxis.TextHeight(Parts(0))
      For i = 0 To UBound(Parts)
         PicAxis.CurrentX = Canvas.Left 'PicAxis.ScaleWidth / 2 - PicAxis.TextWidth(parts(i))
         PicAxis.Print Parts(i) & vbCrLf
      Next i
   End If
   Set PicAxis.Picture = PicAxis.Image
   PicAxis.Refresh
   
   SavePicture PicAxis.Picture, Filename
   'Clipboard.SetData PicAxis.Image, vbCFBitmap
   Set PicAxis.Picture = Nothing
   PicAxis.Cls
   
  ' PicAxis.AutoRedraw = False
   
   Canvas.Visible = True
   UserControl_Resize
   Call DrawGraph

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ZoomIn(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Variant
On Error Resume Next
   'record the location of all the lines before the scale changes
   Dim lX1(), lY1(), lX2(), lY2(), n As Long
   Dim i As Long, ind As Long
   On Error Resume Next
   If UBound(lines) > 0 Then
     n = UBound(lines)
     ReDim lX1(n)
     ReDim lX2(n)
     ReDim lY1(n)
     ReDim lY2(n)
     For i = 0 To n
        ind = lines(i).LineI
        lX1(i) = Line1(ind).X1
        lX2(i) = Line1(ind).X2
        lY1(i) = Line1(ind).Y1
        lY2(i) = Line1(ind).Y2
     Next i
   End If
   
   
   m_AutoScaleXMin = False
   m_AutoScaleXMax = False
   m_AutoScaleY = False
   If X1 < X2 Then
      MaximaX(0) = X1
      MaximaX(1) = X2
   Else
      MaximaX(0) = X2
      MaximaX(1) = X1
   End If
   If Y1 < Y2 Then
      MaximaY(0) = Y1
      MaximaY(1) = Y2
   Else
      MaximaY(0) = Y2
      MaximaY(1) = Y1
   End If
   Call DrawGraph
   
   'now return the lines to that location
   Dim X As Single, Y As Single
   If UBound(lines) > 0 Then
     n = UBound(lines)
     For i = 0 To n
        With lines(i)
          ind = lines(i).LineI
          Line1(ind).X1 = lX1(i)
          Line1(ind).X2 = lX2(i)
          Line1(ind).Y1 = lY1(i)
          Line1(ind).Y2 = lY2(i)
          X = lX1(i): Y = lY1(i)
          Dot(.Dot(0)).Move X - Dot(.Dot(0)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(0)).Height / 2
          X = lX2(i): Y = lY2(i)
          Dot(.Dot(1)).Move X - Dot(.Dot(1)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(1)).Height / 2
        End With
     Next i
   End If
   
   
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ZoomOut() As Variant
On Error Resume Next
   'record the location of all the lines before the scale changes
   Dim lX1(), lY1(), lX2(), lY2(), n As Long
   Dim i As Long, ind As Long
   On Error Resume Next
   If UBound(lines) > 0 Then
     n = UBound(lines)
     ReDim lX1(n)
     ReDim lX2(n)
     ReDim lY1(n)
     ReDim lY2(n)
     For i = 0 To n
        ind = lines(i).LineI
        lX1(i) = Line1(ind).X1
        lX2(i) = Line1(ind).X2
        lY1(i) = Line1(ind).Y1
        lY2(i) = Line1(ind).Y2
     Next i
   End If
   
   
   m_AutoScaleXMax = True
   m_AutoScaleXMin = True
   m_AutoScaleY = True
   Call DrawGraph
   
   'now put them in the correct location for the tew scale
   Dim X As Single, Y As Single
   If UBound(lines) > 0 Then
     n = UBound(lines)
     For i = 0 To n
        With lines(i)
          ind = lines(i).LineI
          Line1(ind).X1 = lX1(i)
          Line1(ind).X2 = lX2(i)
          Line1(ind).Y1 = lY1(i)
          Line1(ind).Y2 = lY2(i)
          X = lX1(i): Y = lY1(i)
          Dot(.Dot(0)).Move X - Dot(.Dot(0)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(0)).Height / 2
          X = lX2(i): Y = lY2(i)
          Dot(.Dot(1)).Move X - Dot(.Dot(1)).Width / 2, Y - Sgn(Canvas.ScaleHeight) * Dot(.Dot(1)).Height / 2
        End With
     Next i
   End If
   
End Function

Public Sub PrintGraph(Copies, Orientation)
On Error Resume Next
    Printer.Copies = Copies
    Printer.Orientation = Orientation
    Canvas.Refresh
    Printer.PaintPicture Canvas.Image, 0, 0
    Printer.EndDoc

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SaveMetaFile(Filename As String) As Variant
On Error Resume Next
mMetaFile = True

UserControl.ScaleMode = vbInches
PicAxis.Width = 5
PicAxis.Height = 3
UserControl.ScaleMode = vbTwips
MetaFile = Filename
Call DrawGraph
mMetaFile = False
Call DrawGraph
Call UserControl_Resize
End Function

Public Function CopyMetaFile()
On Error Resume Next
mMetaFile = True

UserControl.ScaleMode = vbInches
PicAxis.Width = 5
PicAxis.Height = 3
UserControl.ScaleMode = vbTwips
MetaFile = App.path & "\resources\temp\Temp2.wmf"
Call DrawGraph
mMetaFile = False
Call DrawGraph
Call UserControl_Resize
'Clipboard.Clear
Clipboard.SetData LoadPicture(MetaFile), vbCFMetafile
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get MonthAxis() As Boolean
On Error Resume Next
    MonthAxis = m_MonthAxis
End Property

Public Property Let MonthAxis(ByVal New_MonthAxis As Boolean)
On Error Resume Next
    m_MonthAxis = New_MonthAxis
    PropertyChanged "MonthAxis"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get AutoScaleXMin() As Boolean
On Error Resume Next
    AutoScaleXMin = m_AutoScaleXMin
End Property

Public Property Let AutoScaleXMin(ByVal New_AutoScaleXMin As Boolean)
On Error Resume Next
    m_AutoScaleXMin = New_AutoScaleXMin
    PropertyChanged "AutoScaleXMin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get AutoScaleXMax() As Boolean
On Error Resume Next
    AutoScaleXMax = m_AutoScaleXMax
End Property

Public Property Let AutoScaleXMax(ByVal New_AutoScaleXMax As Boolean)
On Error Resume Next
    m_AutoScaleXMax = New_AutoScaleXMax
    PropertyChanged "AutoScaleXMax"
End Property

Public Function FoxRotate(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal SrcDC As Long, ByVal SrcBmp As Long, ByVal TransColor As Long, ByVal Angle As Double, Optional ByVal Flags As Long) As Long
    Dim TmpDC As Long, TmpBmp As Long, OldObject As Long
    Dim BitCount As Long, BitCount2 As Long, LineWidth As Long, LineWidth2 As Long
    Dim retVal As Long
    Dim Width As Long, Height As Long, NewSize As Long
    Dim H As Long, B As Long, f As Long, d As Long, i As Long
    Dim dx1 As Double, dy1 As Double
    Dim TransR As Byte, TransG As Byte, TransB As Byte
    Dim TempAlpha As Byte
    Dim Info As BITMAPINFO, Info2 As BITMAPINFO
    Dim SrcBits() As Byte, TmpBits() As Byte
    On Error Resume Next
    TransR = TransColor And &HFF
    TransG = (TransColor And &HFF00&) / 255
    TransB = (TransColor And &HFF0000) / 65536
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info2.bmiHeader.biSize = Len(Info2.bmiHeader)
    retVal = GetDIBits(SrcDC, SrcBmp, 0, 0, ByVal 0, Info, 0)
    If retVal = 0 Then Exit Function
    TmpDC = CreateCompatibleDC(SrcDC)
    Width = Info.bmiHeader.biWidth
    Height = Info.bmiHeader.biHeight
    NewSize = Math.Sqr(Width ^ 2 + Height ^ 2) + 2
    
    TmpBmp = CreateCompatibleBitmap(SrcDC, NewSize, NewSize)
    If TmpBmp Then
        OldObject = SelectObject(TmpDC, TmpBmp)
        BitBlt TmpDC, 0, 0, NewSize, NewSize, DstDC, DstX - NewSize / 2, DstY - NewSize / 2, vbSrcCopy

        Info.bmiHeader.biBitCount = 24
        Info.bmiHeader.biCompression = 0
        Info2.bmiHeader.biBitCount = 24
        Info2.bmiHeader.biCompression = 0
        Info2.bmiHeader.biPlanes = 1
        Info2.bmiHeader.biHeight = NewSize
        Info2.bmiHeader.biWidth = NewSize
        
        LineWidth = Width * 3
        If (LineWidth Mod 4) Then LineWidth = LineWidth + 4 - (LineWidth Mod 4)
        BitCount = LineWidth * Height
        
        LineWidth2 = NewSize * 3
        If (LineWidth2 Mod 4) Then LineWidth2 = LineWidth2 + 4 - (LineWidth2 Mod 4)
        BitCount2 = LineWidth2 * NewSize
        
        ReDim SrcBits(BitCount - 1)
        ReDim TmpBits(BitCount2 - 1)
        GetDIBits SrcDC, SrcBmp, 0, Height, SrcBits(0), Info, 0
        GetDIBits TmpDC, TmpBmp, 0, NewSize, TmpBits(0), Info2, 0




        Dim CurOffset As Long
        Dim NewX As Double, NewY As Double
        Dim Xmm As Long, Ymm As Long
        Dim I1 As Long
        Dim v1 As Boolean
        dx1 = Cos(Angle * PIDEG)
        dy1 = Sin(Angle * PIDEG)
        
        For H = 0 To NewSize - 1
            CurOffset = LineWidth2 * H
            For B = 0 To NewSize - 1
                f = CurOffset + 3 * B
                NewX = Width / 2 + (B - NewSize / 2) * dx1 - (H - NewSize / 2) * dy1
                NewY = Height / 2 + (B - NewSize / 2) * dy1 + (H - NewSize / 2) * dx1
                
                Xmm = Int(NewX + 0.5)
                Ymm = Int(NewY + 0.5)
                If (Xmm >= 0) And (Xmm < Width) And (Ymm >= 0) And (Ymm < Height) Then
                    v1 = True
                    I1 = LineWidth * Ymm + 3 * Xmm
                    If Flags And &H1 Then
                        v1 = Not (SrcBits(I1 + 2) = TransR And SrcBits(I1 + 1) = TransG And SrcBits(I1) = TransB)
                    End If
                    If v1 Then For d = 0 To 2: TmpBits(f + d) = SrcBits(I1 + d): Next d
                End If
            Next B
        Next H
        
        SetDIBitsToDevice DstDC, DstX - NewSize / 2, DstY - NewSize / 2, NewSize, NewSize, 0, 0, 0, NewSize, TmpBits(0), Info2, 0
        Erase SrcBits
        Erase TmpBits
        DeleteObject SelectObject(TmpDC, OldObject)
    End If
    DeleteDC TmpDC
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
