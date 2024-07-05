VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Helpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit


Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function GaussianRandom(Mean As Double, Range As Double) As Double


    On Error GoTo Err_Proc
Dim X1 As Double, X2 As Double, W As Double, Y1 As Double, Y2 As Double
W = 3
While W >= 1
  X1 = 2# * Rnd() - 1#
  X2 = 2# * Rnd() - 1#
  W = X1 * X1 + X2 * X2
Wend

W = ((-2# * Log(W)) / W) ^ 0.5
Y1 = X1 * W
Y2 = X2 * W

GaussianRandom = (Y1 + Y2) / 4 * Range + Mean
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "Helpers", "GaussianRandom", Err.Description
    Resume Exit_Proc


End Function

Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)

End Function
