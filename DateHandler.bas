Attribute VB_Name = "DateHandler"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.



Public Function IsoDate(dteDate As Date)
'Version 1.0
   If IsDate(dteDate) = True Then
      Dim dteDay, dteMonth, dteYear
      dteDay = Day(dteDate)
      dteMonth = Month(dteDate)
      dteYear = Year(dteDate)
      IsoDate = dteYear & _
         "-" & Right(CStr(dteMonth + 100), 2) & _
         "-" & Right(CStr(dteDay + 100), 2)
   Else
      IsoDate = Null
   End If
End Function
Public Function IsoDateString(dteMonth As Integer, dteDay As Integer, dteYear As Integer) As String
 IsoDateString = dteYear & _
         "-" & Right(CStr(dteMonth + 100), 2) & _
         "-" & Right(CStr(dteDay + 100), 2)
End Function
Public Function IsoDateStringString(dMonth As String, dDay As String, dYear As String) As String
Dim dteMonth As Integer, dteDay As Integer, dteYear As Integer
dteMonth = Val(dMonth)
dteDay = Val(dDay)
dteYear = Val(dYear)
 IsoDateStringString = dteYear & _
         "-" & Right(CStr(dteMonth + 100), 2) & _
         "-" & Right(CStr(dteDay + 100), 2)
End Function
Public Function MonthName(dteDate As Date) As String
   Dim Months(12) As String
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
   MonthName = Months(Month(dteDate))
End Function

Public Function MonthAbbrev(dteDate As Date) As String
   Dim Months(12) As String
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
   MonthAbbrev = Left$(Months(Month(dteDate)), 3)
End Function


