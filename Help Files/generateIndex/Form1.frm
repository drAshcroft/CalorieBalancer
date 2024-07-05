VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   7920
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim htmls() As String
Dim cc As Long


Private Sub Form_Load()
File1.Path = App.Path
File1.Pattern = "*.htm"
ReDim htmls(File1.ListCount - 1)
For I = 0 To File1.ListCount - 1
  htmls(I) = File1.List(I)
Next I
cc = 0
Open App.Path & "\index.txt" For Output As #1

WB.Navigate2 App.Path & "\" & htmls(0)

End Sub


Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  Dim doc As HTMLDocument
  Set doc = WB.Document
  Title = doc.Title
  Set ele = doc.getElementsByTagName("h2")
  For Each e In ele
  junk = e.innerHTML
  If InStr(1, junk, "<a name", vbTextCompare) <> 0 Then
     I = InStr(1, junk, "></a>", vbTextCompare)
     junk = Left$(junk, I)
     junk = Replace(junk, "<a name=", "#", , , vbTextCompare)
     junk = Replace(junk, ">", "")
     junk = htmls(cc) & junk
     Debug.Print junk
  
  Print #1, "<LI> <OBJECT type=""Text/sitemap"">"
  Print #1, "         <param name=""Name"" value=""" & e.innerText & """>"
  Print #1, "         <param name=""Name"" value=""" & Title & """>"
  
  Print #1, "         <param name=""Local"" value=""NewHTML/" & junk & """>"
  Print #1, "     </OBJECT>"
  
  End If
   
  
  
  Next
  cc = cc + 1
  WB.Navigate2 App.Path & "\" & htmls(cc)
End Sub

