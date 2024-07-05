VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "exCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Dim vars As Collection
Dim Names As Collection
Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Public Function Count()
On Error Resume Next
Count = vars.Count
End Function

Public Function ItemName(Key As Variant) As Variant
On Error Resume Next
ItemName = Names(Key)
End Function

Public Sub Remove(Key As Variant)
On Error Resume Next
vars.Remove Key
Names.Remove Key
End Sub

Public Sub RemoveAll()
  On Error Resume Next
  Dim i As Long
  For i = 1 To vars.Count
    vars.Remove 1
  Next i
  For i = 1 To Names.Count
    Names.Remove 1
  Next i
End Sub

Private Sub Class_Initialize()
On Error Resume Next
Set vars = New Collection
Set Names = New Collection
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set vars = Nothing
Set Names = Nothing
End Sub

Public Function Add(Item, Optional Key, Optional Before, Optional After)
On Error Resume Next
  If Key <> "" Then
    vars.Remove Key
    Names.Remove Key
  End If
'On Error
vars.Add Item, Key, Before, After
If Key <> "" Then Names.Add Key, Key, Before, After
Err.Clear
End Function

Public Function Item(Key)
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Item = vars.Item(Key)
  If Err.Number <> 0 Then
    Set Item = vars.Item(Key)
  End If
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)
End Function
