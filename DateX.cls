VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable(s) to hold property value(s)
Private mvarDef As String 'local copy
Private mvarRDate As Date 'local copy

Public Property Let RDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RDate = 5
    mvarRDate = vData
    mvarDef = Month(mvarRDate) & "/" & Day(mvarRDate) & "/ " & Year(mvarRDate)
End Property


Public Property Get RDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RDate
    RDate = mvarRDate
End Property



Public Property Let Def(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Def = 5
  '  Dim junks() As String
    
  '  junks = Split(vdate, "/")
    mvarDef = vData
    mvarRDate = vData
End Property


Public Property Get Def() As String
Attribute Def.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Def
    Def = mvarDef 'Month(mvarRDate) & "/" & Day(mvarRDate) & "/ " & Year(mvarRDate)
End Property

Public Function Today() As String
  
    Today = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function


