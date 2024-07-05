VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Application :   akReportsSample
' Description :   akreports Sample demonstration applications
' Author      :   Andrew Keat
' Email       :   aktools@ihug.co.nz
' Web Site    :   http://aktools.rentmaster.co.nz
' Create Date :   January 2001
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents objReport As Report  'The main akReport object
Attribute objReport.VB_VarHelpID = -1

Private Const Filename As String = "MealPlan.rdn"
'These public parameters are provided from the form
Public PrinterName As String
Public Output As OutputOptions
Public OutputFile As String

Dim mFields As Collection

Public Sub PrintMealReport(Fields As Collection)
'This is the only public method of the class - the class does the rest.
'On Error GoTo PrintError
Set mFields = Fields
If objReport.LoadReport(App.path & "\" & Filename) Then
    objReport.OutputTo = Output
    objReport.OutputFilename = OutputFile
    objReport.PrinterName = PrinterName
    PrintMealPlanReport
End If
Exit Sub

PrintError:
    MsgBox Err.Number & " - " & Err.Description, vbOKOnly, "clsTable.PrintReport"
    objReport.Cancel = True
    objReport.EndReport
    Exit Sub
End Sub


Private Sub Class_Initialize()
Set objReport = New Report
End Sub

Private Sub objReport_NewPage()
'The akReport object fires this event every time it starts a new page (except first page)
'which gives us a chance to print a page header.  In the case of all the sample reports
'in this project, the page header section is always called 'pageheader'

'This report does not have a page header section
End Sub

Private Sub objReport_PrintPageFooter(ByVal SectionName As String, ByVal LastPage As Boolean, Cancel As Boolean)
'This event is fired just before the page footer is printer so we
'have a chance to modify the page footer before it is printed.
'The section name is to make it easier for us to identify the section for modification
'The LastPage property tells us if this is the end of the report.
'The Cancel property allows us to stop the page footer from being printed by setting to true.


'For this report we do nothing - just use the default page footer

End Sub

Private Sub objReport_PrintToPrinterButton()
'If the user presses the print to printer button after first printing to screen, this
'event is fired.  To print to printer, the output is changed to printer, and the report
'is printed all over again.
objReport.OutputTo = cOutputToPrinter
PrintMealPlanReport
End Sub


Private Sub PrintMealPlanReport()
On Error GoTo errhandl
'Prints the TABLE report
Dim iRow As Integer
Dim iCol As Integer
With objReport.Section("header")
   .Field("username").Value = mFields("username")
   .Field("weekdates").Value = mFields("daterange")

End With

objReport.PrintSection "Header"

With objReport.Section("info").Field("meals")
    
    '.TableNumberCols = 7        'The table has 7 columns...
    '.TableNumberRows = 10       '...and 10 rows
    '.TableFirstColLabel = True  'The first column and row are determined to be...
    '.TableFirstRowLabel = True  '...row/column headings (optional)
    'For iCol = 2 To .TableNumberCols
    '    .CellValue(iCol, 1) = "Col " & iCol   'Fill in the column headings (optional)
    'Next iCol
    'For iRow = 2 To .TableNumberRows
    '    .CellValue(1, iRow) = "Row " & iRow   'Fill in the row headings (optional)
    'Next iRow
    
    'Here I am generating some random data to fill the table with.
    Dim junk As String
    For iCol = 1 To .TableNumberCols
        For iRow = 1 To .TableNumberRows
            junk = mFields(iRow - 1 & "-" & iCol - 1) 'Fill each cell of the table with a value
            .CellValue(iCol, iRow) = junk
            
        Next iRow
    Next iCol
End With
'Print the section which conatins the now populated table
objReport.PrintSection "info"

objReport.EndReport
Exit Sub
errhandl:
If DoDebug Then
  Debug.Print Err.Description
 Stop
 Resume
End If
End Sub



