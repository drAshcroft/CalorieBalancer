VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DB As Database
Private NutDefs As Collection
Private Function Translate(junk As String) As String
  If junk = "Total lipid (fat)" Then junk = "fat"
  If junk = "Carbohydrate, by difference" Then junk = "carbs"
  If junk = "Energy" Then junk = "calories"
  If junk = "Alcohol, ethyl" Then junk = "alcohol"
  If junk = "Sugars, total" Then junk = "sugar"
  If junk = "Fiber, total dietary" Then junk = "fiber"
  If junk = "Vitamin E (alpha-tocopherol)" Then junk = "vitamin E"
  If junk = "Vitamin C, total ascorbic acid" Then junk = "vitamin c"
  If junk = "Vitamin B-6" Then junk = "Vitamin B6"
  If junk = "Vitamin B-12" Then junk = "Vitamin B12"
  If junk = "Vitamin K (phylloquinone)" Then junk = "Vitamin K"
  If junk = "Folate, DFE" Then junk = "wfasdf"
  If junk = "Vitamin E, added" Then junk = "gsdfg"
  If junk = "Vitamin B-12, added" Then junk = "hsdh"
  If junk = "Fatty acids, total trans" Then junk = "trans fat"
  If junk = "Fatty acids, total saturated" Then junk = "saturated fat"
  If junk = "Fatty acids, total monounsaturated" Then junk = "monounsaturated fat"
  If junk = "Fatty acids, total polyunsaturated" Then junk = "polyunsaturated fat"
  If junk = "Carotene, beta" Then junk = "beta-Carotene"
  If junk = "Carotene, alpha" Then junk = "alpha-Carotene"
  Dim junks() As String
  junks = Split(junk, ",")
  junk = junks(0)
  Translate = junk
End Function
Private Sub LoadNutDef()
Dim junk As String, junks() As String, index As String, jName As String
Open App.Path & "\nutr_def.txt" For Input As #1
Set NutDefs = New Collection
While Not EOF(1)
   Line Input #1, junk
   junks = Split(junk, "^")
   index = Trim$(Replace(junks(0), "~", ""))
   jName = Trim$(Replace(junks(3), "~", ""))
   jName = Translate(jName)
   Debug.Print index & "=" & jName
   NutDefs.Add jName, index
Wend
Close #1
End Sub
Private Sub LoadNutrients()
Dim junk As String, junks() As String, index As String, jName As String
Dim value As Double, rs As Recordset

Open App.Path & "\nut_data.txt" For Input As #1
While Not EOF(1)
   Line Input #1, junk
   junks = Split(junk, "^")
   index = Trim$(Replace(junks(0), "~", ""))
   jName = NutDefs(Trim$(Replace(junks(1), "~", "")))
   value = Val(junks(2))
   Set rs = DB.OpenRecordset("select * from abbrev where ndb_no='" & index & "';", dbOpenDynaset)
   Debug.Print rs("foodname") & "," & jName & "=" & value
   On Error Resume Next
   Err.Clear
   rs.Edit
   rs(jName) = value
   If Err.Number = 0 Then
        rs.Update
   End If
   rs.Close
   DoEvents
Wend
Close #1
End Sub
Private Sub Form_Load()

Call LoadNutDef
Set DB = OpenDatabase(App.Path & "\sr16-2.mdb")
Call LoadNutrients
End Sub
