VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit

    Dim oObject As Object
    Dim sMethod As String
    Dim bInstantiated As Boolean
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Sub Class_Initialize()


    On Error GoTo Err_Proc

    bInstantiated = False
  
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "WebEvent", "Class_Initialize", Err.Description
    Resume Exit_Proc


End Sub
Public Sub Set_Destination(oInObject As Object, sInMethod As String)


    On Error GoTo Err_Proc

    Set oObject = oInObject
    sMethod = sInMethod
    bInstantiated = True
  
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "WebEvent", "Set_Destination", Err.Description
    Resume Exit_Proc


End Sub
Public Sub My_Default_Method()


    On Error GoTo Err_Proc

    If bInstantiated Then
        CallByName oObject, sMethod, VbMethod
    End If
  
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "WebEvent", "My_Default_Method", Err.Description
    Resume Exit_Proc


End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

        Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
