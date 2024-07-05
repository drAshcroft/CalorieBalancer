VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim DB As Database
Dim rs As Recordset
Set DB = OpenDatabase("C:\Documents and Settings\brian\Desktop\free Calories 1.0\New Folder\Resources\sr16-2.mdb")
Set rs = DB.OpenRecordset("Select * from dailylog;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("select * from daysinfo;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from exerciselog;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from meals;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("select * from ideals where user<>'AnyUser';", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from mealdefinition;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from recipes;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from recipesindex;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from mealplanner where planid>=0;", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("Select * from profiles where user<>'average';", dbOpenDynaset)
While Not rs.EOF
  rs.Delete
  rs.MoveNext
Wend
Set rs = DB.OpenRecordset("select index,usage from abbrev where usage>0;", dbOpenDynaset)

While Not rs.EOF
  rs.Edit
  If rs("index") > 0 Then
    rs("usage") = 1
  Else
    rs("usage") = 1000
  End If
  rs.Update
  rs.MoveNext
Wend

Set rs = DB.OpenRecordset("select index,usage from abbrevexercise where usage>0;", dbOpenDynaset)

While Not rs.EOF
  rs.Edit
  If rs("index") > 0 Then
    rs("usage") = 1
  Else
    rs("usage") = 1000
  End If
  rs.Update
  rs.MoveNext
Wend
rs.Close
Set rs = Nothing
DB.Close
Set DB = Nothing
 Name "C:\Documents and Settings\brian\Desktop\free Calories 1.0\New Folder\Resources\sr16-2.mdb" As _
 "C:\Documents and Settings\brian\Desktop\free Calories 1.0\New Folder\Resources\proto.mdb"
 
End
End Sub
