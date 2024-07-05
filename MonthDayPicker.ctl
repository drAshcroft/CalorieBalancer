VERSION 5.00
Begin VB.UserControl MonthDayPicker 
   BackColor       =   &H0098CCD0&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00CEEFF7&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2400
      Picture         =   "MonthDayPicker.ctx":0000
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "MonthDayPicker.ctx":05D2
      Top             =   0
      Width           =   300
   End
   Begin VB.Label MonthL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label DoW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0098CCD0&
      Caption         =   "Sun"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "MonthDayPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Dim Months(12) As String
Dim cYear As Long
Dim cMonth As Long
Dim cDay As Long

Dim tMonth As Long, tDAy As Long, tYear As Long
   
Dim Index As Integer, FirstIndex As Integer, LastIndex As Long
Dim Curdate As Date
Event DateSelected(NewDate As Date)
Event LFocus()
Dim sDays(6) As String

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub SetDate(NewDate As Date)
On Error Resume Next
   cDay = Day(NewDate)
   cYear = Year(NewDate)
   cMonth = Month(NewDate)
   Curdate = NewDate 'cMonth & "/" & cDay & "/" & cYear
   Call FillMonth
End Sub

Public Function GetDate() As Date
On Error Resume Next
   GetDate = Curdate
End Function

Private Sub HScroll1_Change()


    On Error GoTo Err_Proc

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "MonthDayPicker", "HScroll1_Change", Err.Description
    Resume Exit_Proc


End Sub
Sub FillMonth()
On Error GoTo errhandl
   Dim StartDate As Double
   Dim theDay As Integer, RecNo As Integer
   
   Dim WW As Single, HH As Single, TT As Single, jj As String
   
   Picture2.Cls
   
   WW = Picture2.ScaleWidth / 7
   HH = Picture2.ScaleHeight / 6


   '.monthbox.Caption = Months(m_Month) & m_Year
   StartDate = DateSerial(cYear, cMonth, 1)
   
   FirstIndex = Weekday(StartDate) - 1
   Index = FirstIndex
   i = 1
   MonthL.Caption = Months(cMonth) & " " & cYear
   While cMonth = Month(StartDate + Index - FirstIndex)
       jj = " " & STR$(i)
       Index = Index + 1
       Picture2.CurrentX = (((Index - 1) Mod 7) + 1) * WW - Picture2.TextWidth(jj)
       Picture2.CurrentY = (Int((Index - 1) / 7)) * HH
       
       If cMonth = tMonth And cYear = tYear Then
         If (tDAy = i) Then
            Picture2.FontBold = True
            Picture2.ForeColor = vbBlue
            Picture2.CurrentX = (((Index - 1) Mod 7) + 1) * WW - Picture2.TextWidth(jj)
            Picture2.Print jj
            Picture2.FontBold = False
            Picture2.ForeColor = vbBlack
         Else
            Picture2.Print jj
         End If
       Else
         Picture2.Print jj
       End If
       If cDay = i Then
          Picture2.Line ((((Index - 1) Mod 7)) * WW + 1, (Int((Index - 1) / 7)) * HH)-((((Index - 1) Mod 7) + 1) * WW, (Int((Index - 1) / 7)) * HH + Picture2.TextHeight(j)), vbRed, B
       End If

       'Index = Index +1
       LastIndex = i
       i = i + 1
       
    Wend
    Exit Sub
errhandl:
    MsgBox Err.Description, vbOKOnly, ""
End Sub



Private Sub Image1_Click()
On Error Resume Next
cMonth = cMonth - 1
If cMonth = 13 Then
   cMonth = 1
   cYear = cYear + 1
End If
If cMonth = 0 Then
  cMonth = 12
  cYear = cYear - 1
End If
MonthL.Caption = Months(cMonth) & " " & cYear
Call FillMonth
Dim d As Date
d = IsoDateString(CInt(cMonth), CInt(cDay), CInt(cYear)) 'cMonth & "/" & cDay & "/" & cYear
RaiseEvent DateSelected(d)
End Sub

Private Sub Image2_Click()
On Error Resume Next
cMonth = cMonth + 1
If cMonth = 13 Then
   cMonth = 1
   cYear = cYear + 1
End If
If cMonth = 0 Then
  cMonth = 12
  cYear = cYear - 1
End If
MonthL.Caption = Months(cMonth) & " " & cYear
Call FillMonth
Dim d As Date
d = IsoDateString(CInt(cMonth), CInt(cDay), CInt(cYear)) 'cMonth & "/" & cDay & "/" & cYear
'd = cMonth & "/" & cDay & "/" & cYear
RaiseEvent DateSelected(d)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  Dim WW As Single, HH As Single
  Dim dd As Long
   WW = Picture2.ScaleWidth / 7
   HH = Picture2.ScaleHeight / 6
   
   dd = Int(Y / HH) * 7 + Fix(X / WW) - FirstIndex + 1
   If dd >= 1 And dd <= LastIndex Then
     cDay = dd
     Curdate = DateHandler.IsoDateString(CInt(cMonth), CInt(cDay), CInt(cYear)) 'cMonth & "/" & cDay & "/" & cYear
     Call FillMonth
   End If
   
  Dim d As Date
d = IsoDateString(CInt(cMonth), CInt(cDay), CInt(cYear))
RaiseEvent DateSelected(d)
   
End Sub

Private Sub UserControl_ExitFocus()
On Error Resume Next
 RaiseEvent LFocus
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
   cYear = Year(Date)
   cMonth = Month(Date)
   cDay = Day(Date)
   
   tYear = Year(Date)
   tMonth = Month(Date)
   tDAy = Day(Date)


   Curdate = IsoDateString(CInt(cMonth), CInt(cDay), CInt(cYear))
  
   
   Months(1) = "January"
   Months(2) = "February"
   Months(3) = "March"
   Months(4) = "April"
   Months(5) = "May"
   Months(6) = "June"
   Months(7) = "July"
   Months(8) = "August"
   Months(9) = "September"
   Months(10) = "October"
   Months(11) = "November"
   Months(12) = "December"
   
   
   MonthL.Caption = Months(cMonth) & " " & cYear
   sDays(0) = "Sun"
   sDays(1) = "Mon"
   sDays(2) = "Tue"
   sDays(3) = "Wed"
   sDays(4) = "Thu"
   sDays(5) = "Fri"
   sDays(6) = "Sat"
   Dim s As Single
   s = UserControl.ScaleWidth / 7
   
   For i = 2 To 7
    If DoW.Count < i Then
      Load DoW(i)
    End If
     DoW(i).Caption = sDays(i - 1)
     DoW(i).Visible = True
     DoW(i).Enabled = True
     DoW(i).Left = (i - 1) * s
   Next i
   
   Call FillMonth
End Sub

Private Sub UserControl_LostFocus()
On Error Resume Next
RaiseEvent LFocus
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
'Picture1.Width = UserControl.ScaleWidth
'Picture1.Height = UserControl.ScaleHeight
Image1.Move 0, 0
Image2.Move UserControl.ScaleWidth - Image2.Width, 0
MonthL.Move Image1.Width, 0, UserControl.ScaleWidth - Image1.Width - Image2.Width
Dim s As Single, i As Long
Dim T As Single
T = Image1.Height + 2
   s = UserControl.ScaleWidth / 7
   DoW(0).Move 0, T
   For i = 2 To 7
    If DoW.Count < i Then
      Load DoW(i)
    End If
     DoW(i).Caption = sDays(i - 1)
     DoW(i).Visible = True
     DoW(i).Enabled = True
     DoW(i).Move (i - 1) * s, T
   Next i

Picture2.Width = UserControl.ScaleWidth
Picture2.Height = UserControl.ScaleHeight - T - DoW(0).Height

Call FillMonth
End Sub

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

     Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
