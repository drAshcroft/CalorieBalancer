Attribute VB_Name = "MathMod"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private Const e = 2.71828182845905
Private Const Pi = 3.14159265358979

Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Function det(A)


    On Error GoTo Err_Proc
    '{
    'var Length = A.length-1;
    Dim length As Long
    length = UBound(A) ' - 1
        '// formal length of a matrix is one bigger
    If (length = 1) Then
        det = A(1, 1)
        Exit Function
    Else
        '{
        Dim i As Long
        Dim sum As Double
        Dim factor As Double
        sum = 0
        factor = 1
        'var i;
        'var sum = 0;
        'var factor = 1;
        For i = 1 To length
            '{
            If A(1, i) <> 0 Then
                '{
                '// create the minor
                Dim Minor() As Double
                ReDim Minor(length - 1, length - 1)
                Dim m As Long 'var m;
                Dim n As Long 'var n;
                Dim theColumn As Long 'var theColumn;
                For m = 1 To length - 1 '; m++) // columns
                    '{
                    If (m < i) Then
                        theColumn = m ';
                    Else
                        theColumn = m + 1
                    End If
                        
                    For n = 1 To length - 1
                        
                        Minor(n, m) = A(n + 1, theColumn)
'// alert(minor[n][m]);
                    Next n '    } // n
                Next m '    } // m
                '// compute its determinant
                sum = sum + A(1, i) * factor * det(Minor) ';
            End If
            factor = -1 * factor ';   // alternating sum
        Next i '    } // end i
    End If '    } // recursion
    det = sum 'return(sum);
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "MathMod", "det", Err.Description
    Resume Exit_Proc


End Function '    } // end determinant

Private Function inverse(A) ' {


    On Error GoTo Err_Proc

    
    'var Length = A.length - 1;
    'var B = new makeArray2(Length, Length);  // inverse
    'var d = det(A);
    Dim length As Long
    Dim B() As Double
    Dim d
    length = UBound(A) ' - 1
    ReDim B(length, length)
    d = det(A)
    If (d = 0) Then
       Debug.Print "Singular matrix"
    Else
        '{
        Dim i As Long, j As Long
        For i = 1 To length
            '{
            For j = 1 To length
                '{
                '// create the minor
                Dim Minor() As Double
                ReDim Minor(length - 1, length - 1)
                'Minor = makeArray2(length - 1, length - 1) ';
                Dim m As Long, n As Long, theColumn As Long, theRow As Long
                'var m;
                'var n;
                'var theColumn;
                'var theRow;
                For m = 1 To length - 1 '; m++) // columns
                    '{
                    If (m < j) Then
                        theColumn = m
                    Else
                        theColumn = m + 1
                    End If
                    For n = 1 To length - 1 '; n++)
                        '{
                        If (n < i) Then
                           theRow = n
                        Else
                           theRow = n + 1
                        End If
                        Minor(n, m) = A(theRow, theColumn)
'// alert(minor[n][m]);
                    Next n
                Next m
                    '    } // n
                   ' } // m
                '// inverse entry
                Dim temp As Double
                temp = (i + j) / 2 '                var temp = (i+j)/2;
                Dim factor As Long
                If (temp = Round(temp)) Then
                   factor = 1
                Else
                   factor = -1
                End If
                
                B(j, i) = det(Minor) * factor / d ';

                
            Next j '    } // j
            
        Next i '    } // end i
    End If '    } // recursion
    inverse = B 'return(B);
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "MathMod", "inverse", Err.Description
    Resume Exit_Proc


End Function '    } // end inverse

Public Function LinRegr(numX As Long, nPoints As Long, X() As Single, Y() As Single) As Single()
  On Error GoTo errhandl
    Dim k As Long, i As Long, j As Long, sum As Double
    Dim B() As Double, p() As Double
    Dim invP() As Double
    Dim mtemp As Long
        ReDim B(numX + 1)
        ReDim p(numX + 1, numX + 1)
        ReDim invP(numX + 1, numX + 1)
'        B = new makeArray(M+1);
'        P = new makeArray2(M+1, M+1);
'        invP = new makeArray2(M+1, M+1);
        mtemp = numX + 1
        'var mtemp = M+1;
'//      if (N < M+1) alert("your need at least "+ mtemp +" points");
 '           // First define the matrices B and P
          For i = 1 To nPoints
             X(0, i) = 1 ';
          Next i
          For i = 1 To numX + 1
                
                sum = 0
                For k = 1 To nPoints
                  sum = sum + X(i - 1, k) * Y(k)
                Next k
                B(i) = sum ';
                
                For j = 1 To numX + 1 '; j++)
                    
                    sum = 0
                    For k = 1 To nPoints
                      sum = sum + X(i - 1, k) * X(j - 1, k)
                    Next k
                    p(i, j) = sum ';
                Next j
          Next i '} // i
          invP = inverse(p)
          Dim Coeff() As Single
          ReDim Coeff(numX)
          For k = 0 To numX
                sum = 0
    
                For j = 1 To numX + 1
                    sum = sum + invP(k + 1, j) * B(j)
                Next j

                Coeff(k) = sum
          Next k
          LinRegr = Coeff
errhandl:
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

     Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
