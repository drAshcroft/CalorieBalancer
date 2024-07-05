Attribute VB_Name = "REadScriptMod"
'This software is distributed under the GPL v3 or above. It was written by Brian Ashcroft
'accounts@caloriebalancediet.com.   Enjoy.
Option Explicit
Private USER As String
Dim Planscript As String
Public MaxDay As Date
Public firstday As Date
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CopyFile Lib "kernel32" _
  Alias "CopyFileA" (ByVal lpExistingFileName As String, _
  ByVal lpNewFileName As String, ByVal bFailIfExists As Long) _
  As Long
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000

Const KEY_QUERY_VALUE = &H1
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Dim vbSep As String
Dim PlanFile As String

Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or _
    KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0
Private Const REG_EXPAND_SZ = 2

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long _
     ) As Long




Private m_iErr_Handle_Mode As Long 'Init this variable to the desired error handling manage

Function regGetBuyURL(ByVal publisher As String, ByVal appName As String, ByVal appVer As String) _
As String

    On Error GoTo Err_Proc
     'As String
        Dim hKey As Long    ' receives a handle opened registry key
        Dim stringbuffer As String  ' receives data read from the registry
        Dim datatype As Long  ' receives data type of read value
        Dim slength As Long  ' receives length of returned data
        Dim retVal As Long  ' return value

        ' form the registry key path
        Dim keyPath
        keyPath = "SOFTWARE\Digital River\SoftwarePassport\" & publisher & "\" & appName & "\" & appVer
                
        ' open the registry key
        ' try to get from HKEY_LOCAL_MACHINE first
        retVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, hKey)
        ' if fail to get from HKEY_LOCAL_MACHINE branch, try HKEY_CURRENT_USER
        If retVal <> 0 Then
            retVal = RegOpenKeyEx(HKEY_CURRENT_USER, keyPath, 0, KEY_READ, hKey)
        End If
        
        If retVal = 0 Then
            ' Make room in the buffer to receive the incoming data.
            stringbuffer = Space(1024)
            slength = 1024
            
            ' Read the "BuyURL" value from the registry key.
            retVal = RegQueryValueEx(hKey, "BuyURL", 0, datatype, ByVal stringbuffer, slength)
            If retVal = 0 Then
                stringbuffer = Left(stringbuffer, slength - 1)
            Else
                ' "BuyURL" does not exists
                stringbuffer = ""
            End If
            
            ' Close the registry key.
            retVal = RegCloseKey(hKey)
        End If
        
        regGetBuyURL = stringbuffer
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "regGetBuyURL", Err.Description
    Resume Exit_Proc


     End Function



Function CheckRegistryKey(ByVal hKey As Long, ByVal keyname As String) As Boolean


    On Error GoTo Err_Proc
    Dim handle As Long
    'try to open the key
    If RegOpenKeyEx(hKey, keyname, 0, KEY_READ, handle) = 0 Then    ' success
        'the key exists
        CheckRegistryKey = True
        'close it before exiting
        RegCloseKey handle
    End If
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "CheckRegistryKey", Err.Description
    Resume Exit_Proc


End Function

Public Function GetRegString(hKey As Long, strSubKey As String, strValueName As String) As String


    On Error GoTo Err_Proc
Dim strSetting As String
Dim lngDataLen As Long
Dim lngRes As Long
    
    If RegOpenKey(hKey, strSubKey, lngRes) = ERROR_SUCCESS Then
       strSetting = Space(255)
       lngDataLen = Len(strSetting)
       If RegQueryValueEx(lngRes, strValueName, ByVal 0, REG_EXPAND_SZ, ByVal strSetting, lngDataLen) = ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                GetRegString = Left(strSetting, lngDataLen - 1)
            End If
        End If
    
        If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
           MsgBox "RegCloseKey Failed: " & _
           strSubKey, vbCritical
        End If
    End If
    
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "GetRegString", Err.Description
    Resume Exit_Proc


End Function

Function GetRegistryValue(ByVal hKey As Long, ByVal keyname As String, _
    ByVal ValueName As String, ByVal KeyType As Integer, _
    Optional DefaultValue As Variant = Empty) As Variant


    On Error GoTo Err_Proc

    Dim handle As Long, resLong As Long
    Dim resString As String, length As Long
    Dim resBinary() As Byte
    
    'prepare the default result
    GetRegistryValue = DefaultValue
    'open the key, exit if not found
    'If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    
    Select Case KeyType
        Case REG_DWORD
            'read the value, use the default if no found
            If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, _
                resLong, 4) = 0 Then
                GetRegistryValue = resLong
            End If
        Case REG_SZ
            length = 1024: resString = Space$(length)
            If RegQueryValueEx(handle, ValueName, 0, REG_SZ, _
                ByVal resString, length) = 0 Then
                'if value is found, trim characters in excess.
                GetRegistryValue = Left$(resString, length - 1)
            End If
        Case REG_BINARY
            length = 4096
            ReDim resBinary(length - 1) As Byte
            If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, _
                resBinary(0), length) = 0 Then
                ReDim Preserve resBinary(length - 1) As Byte
                GetRegistryValue = resBinary()
            End If
        Case Else
            Err.Raise 1001, , "Unsupported value type"
    End Select
    RegCloseKey handle
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "GetRegistryValue", Err.Description
    Resume Exit_Proc


End Function
    
Function EnumRegistryValues(ByVal hKey As Long, ByVal keyname As String) As Variant()


    On Error GoTo Err_Proc
    Dim handle As Long, Index As Long, valueType As Long
    Dim Name As String, nameLen As Long
    Dim lngValue As Long, strValue As String, dataLen As Long
    
    ReDim result(0 To 1, 0 To 100) As Variant
    
    If Len(keyname) Then
        If RegOpenKeyEx(hKey, keyname, 0, KEY_READ, handle) Then Exit Function
        hKey = handle
    End If
    
    For Index = 0 To 999999
        If Index > UBound(result, 2) Then
            ReDim Preserve result(0 To 1, Index + 99) As Variant
        End If
        nameLen = 260
        Name = Space$(nameLen)
        dataLen = 4096
        ReDim binValue(0 To dataLen - 1) As Byte
        If RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, binValue(0), dataLen) Then Exit For
        result(0, Index) = Left$(Name, nameLen)
        
        Select Case valueType
            Case REG_DWORD
                CopyMemory lngValue, binValue(0), 4
                result(1, Index) = lngValue
            Case REG_SZ
                result(1, Index) = Left$(StrConv(binValue(), vbUnicode), dataLen - 1)
            Case Else
                ReDim Preserve binValue(0 To dataLen - 1) As Byte
                result(1, Index) = binValue()
        End Select
    Next
    
    If handle Then RegCloseKey handle
    
    ReDim Preserve result(0 To 1, Index - 1) As Variant
    EnumRegistryValues = result()
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "EnumRegistryValues", Err.Description
    Resume Exit_Proc


End Function

Function EnumRegistryKeys(ByVal hKey As Long, ByVal keyname As String) As String()


    On Error GoTo Err_Proc
    Dim handle As Long, Index As Long, length As Long
    ReDim result(0 To 100) As String
    Dim FileTimeBuffer(100) As Byte
    
    If Len(keyname) Then
        If RegOpenKeyEx(hKey, keyname, 0, KEY_READ, handle) Then Exit Function
        hKey = handle
    End If
    
    For Index = 0 To 999999
        If Index > UBound(result) Then
            ReDim Preserve result(Index + 99) As String
        End If
        length = 260
        result(Index) = Space$(length)
        If RegEnumKey(hKey, Index, result(Index), length) Then Exit For
        result(Index) = Left$(result(Index), InStr(result(Index), vbNullChar) - 1)
    Next
    
    If handle Then RegCloseKey handle
    ReDim Preserve result(Index - 1) As String
    EnumRegistryKeys = result()
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "EnumRegistryKeys", Err.Description
    Resume Exit_Proc


End Function

Public Function IEVersion() As String
   Dim Version() As Variant
   Dim ver() As String
   Dim i As Byte

   On Error Resume Next
     
    ver() = EnumRegistryKeys(HKEY_LOCAL_MACHINE, _
            "Software\Microsoft\Internet Explorer")
         
    IEVersion = "Invalid"
      
    If UBound(ver) > 0 Then
        Version() = EnumRegistryValues(HKEY_LOCAL_MACHINE, _
                "Software\Microsoft\Internet Explorer")
                
        'MsgBox "Ubound(version,2) = " & UBound(Version, 2)
        For i = 0 To UBound(Version, 2)
            'MsgBox i & ". " & Version(0, i)
            If Version(0, i) = "Version" Then
               Exit For
             '  MsgBox "keluar"
            End If
        Next
        'MsgBox "I terakhir = " & i
      
        If Version(0, i) = "Version" Then IEVersion = Version(1, i)

    End If
    Err.Clear
   'MsgBox IEVersion
End Function

Public Function IEPath() As String


    On Error GoTo Err_Proc
    Dim path() As Variant
    Dim ver As String
    
    ver = IEVersion
    
    If ver = "Invalid" Then
        IEPath = "Invalid"
    Else
        path() = EnumRegistryValues(HKEY_LOCAL_MACHINE, _
            "Software\Microsoft\IE Setup\Setup")
        IEPath = path(1, 7)
    End If
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "IEPath", Err.Description
    Resume Exit_Proc


End Function

Public Function IsAppPresent(strSubKey$, strValueName$) As Boolean


    On Error GoTo Err_Proc
    IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)))
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "IsAppPresent", Err.Description
    Resume Exit_Proc


End Function

Public Function AcrobatVersion() As String
    Dim Version() As Variant
    Dim acrobat() As String
    Dim ver() As String
    
    On Error Resume Next
    
    ver() = EnumRegistryKeys(HKEY_CLASSES_ROOT, _
        "AcroExch.FDFDoc")
        
    AcrobatVersion = "Invalid"
    
    If UBound(ver) > 0 Then
        Version() = EnumRegistryValues(HKEY_CLASSES_ROOT _
            , "AcroExch.FDFDoc\AcrobatVersion")
        'Text1.Text = Version(1, 0)
        acrobat() = EnumRegistryKeys(HKEY_LOCAL_MACHINE, _
            "Software\Adobe\Acrobat")
        'Text1.Text = acrobat(0)

        If Version(1, 0) Or acrobat(0) Then
            AcrobatVersion = Version(1, 0)
        ElseIf acrobat(0) Then
            AcrobatVersion = acrobat(0)
        Else
            AcrobatVersion = "Invalid Version"
        End If
    End If
    Err.Clear
End Function

Public Function AcrobatPath() As String
    Dim Paths() As Variant
    Dim path As String
    Dim strPath As String
    Dim start As Integer
    
    Paths() = EnumRegistryValues(HKEY_CLASSES_ROOT, _
        "AcroExch.FDFDoc\shell\open\command")
    
    If AcrobatVersion = "4.0" Then
        start = 2
    ElseIf AcrobatVersion = "3.0" Then
        start = 1
    End If
    
    On Error Resume Next
    strPath = Paths(1, 0)
    path = Mid(strPath, start, InStr(2, strPath, """") - 2)
    AcrobatPath = path
    Err.Clear
End Function

'Get default browser path and exe file
Private Function BrowserPath() As String
    Dim Paths() As Variant
    Dim path As String
    Dim strPath As String
    
    path = "Invalid"
    
    Paths() = EnumRegistryValues(HKEY_CLASSES_ROOT, _
        "http\shell\open\command")
        
    On Error Resume Next
    strPath = Paths(1, 0)
    path = Mid(strPath, 1, InStr(2, strPath, """"))
    If path = "" Then
       path = "C:\Program Files\Internet Explorer\iexplore.exe"
    End If
    BrowserPath = path
    Err.Clear
End Function

Public Sub OpenURL(strURL As String, Optional Style As VbAppWinStyle = vbNormalFocus)

  '/Call MsgBox(strURL)

    On Error GoTo Err_Proc
Dim bPath As String

    bPath = BrowserPath
    If bPath <> "Invalid" Then
        Shell bPath & Chr(32) & """" & strURL & """", Style
        DoEvents

    Else
        MsgBox "Unable to open URL. Install browser first!"
    End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "REadScriptMod", "OpenURL", Err.Description
    Resume Exit_Proc


End Sub



Public Sub UpdateInternetAbbrev(Matrix, Today As Date, LoadIntoDay As Boolean, Optional RecipeID As Long)


    On Error GoTo Err_Proc
 Dim IDs() As Long
 Dim i As Long, Parts() As String, temp As Recordset, temp2 As Recordset
 Dim MaxIt As Long
          ReDim IDs(UBound(Matrix, 2))
          For i = 0 To UBound(Matrix, 2)
           
           If Matrix(3, i) <> "" Then
              Set temp = DB.OpenRecordset("Select * from abbrev where foodname = '" & _
              Replace(Matrix(3, i), "'", "''") & "';", dbOpenDynaset)
              
           If temp.EOF Then
              Set temp2 = DB.OpenRecordset("Select max(index) as MaxIT from abbrev;", dbOpenDynaset)
              MaxIt = temp2("Maxit") + 1
              temp2.Close
              temp.AddNew
              Set temp2 = DB.OpenRecordset("Select * from weight where index = " & _
                          MaxIt & ";", dbOpenDynaset)
              temp2.AddNew
              temp("Index") = MaxIt
              IDs(i) = MaxIt
              temp2("msre_desc") = Matrix(2, i)
              temp2("Amount") = Matrix(1, i)
              temp2("gm_wgt") = Matrix(10, i)
              temp2("Index") = MaxIt
              
              temp2.Update
              temp("Foodname") = Matrix(3, i)
              temp("Calories") = Matrix(4, i)
              temp("Sugar") = Matrix(5, i)
              temp("Fiber") = Matrix(6, i)
              temp("Carbs") = Matrix(7, i)
              temp("Fat") = Matrix(8, i)
              temp("Protein") = Matrix(9, i)
              temp.Update
              temp.Close
              temp2.Close
              Set temp2 = Nothing
             Else
               IDs(i) = temp("Index")
             End If
             End If
          Next i
          Set temp = Nothing
 
 
       If LoadIntoDay Then
          Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.date)=#" & FixDate(Today) & "#) AND (DaysInfo.user='" & CurrentUser.Username & "')) ORDER BY daysinfo.order;", dbOpenDynaset)
          While Not temp.EOF
             temp.Delete
             temp.MoveNext
          Wend
          For i = 0 To UBound(Matrix, 2)
           If Matrix(3, i) <> "" Then
              temp.AddNew
              temp("date") = Today 'FixDate(Today)
              temp("User") = CurrentUser.Username
              temp("ItemId") = IDs(i)
              If IDs(i) >= 0 Then
                  temp("unit") = Matrix(2, i)
                  temp("servings") = Matrix(1, i)
              End If
              temp("Order") = i
              temp.Update
           End If
          Next i
          temp.Close
          Set temp = Nothing
      Else
          Set temp = DB.OpenRecordset("SELECT * FROM recipes WHERE recipeID =" & RecipeID & " ;", dbOpenDynaset)
          While Not temp.EOF
             temp.Delete
             temp.MoveNext
          Wend
          For i = 0 To UBound(Matrix, 2)
           If Matrix(3, i) <> "" Then
              temp.AddNew
              temp("recipeid") = RecipeID
              temp("ItemId") = IDs(i)
              If IDs(i) >= 0 Then
                  temp("unit") = Matrix(2, i)
                  temp("servings") = Matrix(1, i)
              End If
              temp.Update
           End If
          Next i
          temp.Close
          Set temp = Nothing
           
      End If
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "REadScriptMod", "UpdateInternetAbbrev", Err.Description
    Resume Exit_Proc


End Sub


Public Function APIFileCopy(Src As String, dest As String, _
  Optional FailIfDestExists As Boolean) As Boolean


    On Error GoTo Err_Proc

Dim lRet As Long
lRet = CopyFile(Src, dest, FailIfDestExists)
APIFileCopy = (lRet > 0)

Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler "REadScriptMod", "APIFileCopy", Err.Description
    Resume Exit_Proc


End Function


Public Function UpdateScript(Filename As String) As Date
  Dim OriginalFilename As String
  OriginalFilename = Filename
  If InStr(1, Filename, "resources\plans", vbTextCompare) = 0 Then
     Dim junks() As String
     junks = Split(Filename, "\")
     Call CopyFile(Filename, App.path & "\resources\plans\" & junks(UBound(junks)), False)
     Filename = App.path & "\resources\plans\" & junks(UBound(junks))
     OriginalFilename = Filename
  End If
  If InStr(1, Filename, "cbm", vbTextCompare) <> 0 Then
     Call CopyFile(Filename, App.path & "\resources\temp\temp.mdb", False)
     Filename = App.path & "\resources\temp\temp.mdb"
  End If

  On Error GoTo errhandl
  'on error Resume Next
   Dim DBOut As Database, i As Long, j As Long
   Dim DBin As Database
   Dim AbbrevTrans As Collection, AbbrevItems As Collection
   Dim RecipeTrans As Collection, RecipeNames As Collection
   Dim Extrans As Collection, ExItems As Collection
   Dim MealPlans As Collection, MealPlanNames As Collection
   Dim MealNameTrans As Collection, MealNames As Collection
   
   Set DBin = DB
   Set DBOut = OpenDatabase(Filename)
   'On Error Resume Next
   Call CopyTable(DBin, DBOut, "foodgroups", "category", False)
'   On Error GoTo ErrHandl
   
   Call CopyTable(DBin, DBOut, "profiles", "user", False)
   Call CopyDailyLog(DBin, DBOut, "dailylog", "user", False)
   Call CopyTable(DBin, DBOut, "ideals", "user", False)
   
   
   'need to move dayslog, dailyinfo,exerciselog, and ideals for update
   Call CopyTableWithTrans(DBin, DBOut, "abbrev", "index", "foodname", AbbrevTrans, AbbrevItems)
   Call CopyTableWithTrans(DBin, DBOut, "AbbrevExercise", "index", "exercisename", Extrans, ExItems)
   Call CopyToTrans(DBin, DBOut, "weight", "index", "msre_desc", AbbrevTrans, AbbrevItems)
   'ok gave up on the general  stuff and just used specific routines
   Call CopyRecipes(DBin, DBOut, AbbrevTrans, RecipeTrans, RecipeNames)
   Call CopyRecipesItems(DBin, DBOut, RecipeNames, RecipeTrans, AbbrevTrans)
   
   Call CopyMealPlans(DBin, DBOut, MealPlans, MealPlanNames, "")
   Call CopyMeals(DBin, DBOut, MealPlans, MealPlanNames, MealNameTrans, MealNames, "")
   Call CopyMealIngred(DBin, DBOut, MealNames, MealNameTrans, AbbrevTrans)
   Call LoadMealDays(DBin, DBOut, Today, Today, "", MealNameTrans)
   Call ExerciseDays(DBin, DBOut, Extrans, Today, "")
  
  
  
errhandl:
If DoDebug And Err.Number <> 0 Then
  Debug.Print Err.Description
  Stop
  Resume
End If
  Call frmMain.RefreshDay
  Call frmMain.MakeMealList
  'ReadScript = date
  

End Function

Private Sub CopyDaysInfo(DBin As Database, DBOut As Database, MealTrans As Collection, AbbrevTrans As Collection)
On Error GoTo errhandl
 Dim MDIn As Recordset, MDout As Recordset
 Dim junk As String
 Dim MaxIt As Long, i As Long, j As Long
 Set MDIn = DBin.OpenRecordset("select max(index) as maxit from daysinfo;", dbOpenDynaset)
 MaxIt = DoMax(MDIn)
 Set MDout = DBOut.OpenRecordset("select * from daysinfo;", dbOpenDynaset)
 Set MDIn = DBin.OpenRecordset("select * from daysinfo;", dbOpenDynaset)
 While Not MDout.EOF
      MDIn.AddNew
      On Error Resume Next
      For j = 0 To MDout.Fields.Count - 1
        MDIn.Fields(j) = MDout(MDIn.Fields(j).Name)
      Next j
      On Error GoTo errhandl
      MDIn("index") = MaxIt
      MDIn("itemid") = Val(Replace(AbbrevTrans("_" & MDout("itemid")), "_", ""))
      MDIn("mealid") = Val(Replace(MealTrans("_" & MDout("mealid")), "_", ""))
      MDIn.Update
      MaxIt = MaxIt + 1
      MDout.MoveNext
  Wend
  Set MDIn = Nothing: Set MDout = Nothing
 Exit Sub
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
   

End Sub


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function


Public Function ReadScript(Filename As String, TUser As String, NewPlan As Boolean, Optional UserDate As Date = "1900-01-01") As Date
   Dim DBOut As Database, i As Long, j As Long
   Dim DBin As Database
   Dim AbbrevTrans As Collection, AbbrevItems As Collection
   Dim RecipeTrans As Collection, RecipeNames As Collection
   Dim Extrans As Collection, ExItems As Collection
   Dim MealPlans As Collection, MealPlanNames As Collection
   Dim MealNameTrans As Collection, MealNames As Collection
   Dim RS As Recordset, OtherProfiles As Boolean
  

  Dim OriginalFilename As String
  OriginalFilename = Filename
  If InStr(1, Filename, "resources\plans", vbTextCompare) = 0 Then
     Dim junks() As String
     junks = Split(Filename, "\")
     Call CopyFile(Filename, App.path & "\resources\plans\" & junks(UBound(junks)), False)
     Filename = App.path & "\resources\plans\" & junks(UBound(junks))
     OriginalFilename = Filename
  End If
  If InStr(1, Filename, "cbm", vbTextCompare) <> 0 Then
     Call CopyFile(Filename, App.path & "\resources\temp\temp.mdb", False)
     Filename = App.path & "\resources\temp\temp.mdb"
  End If

  On Error GoTo errhandl
  'on error Resume Next
   
   Set DBin = DB
   Set DBOut = OpenDatabase(Filename)
   On Error Resume Next
   Call CopyTable(DBin, DBOut, "foodgroups", "category", False)
   On Error GoTo errhandl
   Call CopyTable(DBin, DBOut, "profiles", "user", False)
   
   'need to move dayslog, dailyinfo,exerciselog, and ideals for update
   Call CopyTableWithTrans(DBin, DBOut, "abbrev", "index", "foodname", AbbrevTrans, AbbrevItems)
   Call CopyTableWithTrans(DBin, DBOut, "AbbrevExercise", "index", "exercisename", Extrans, ExItems)
   Call CopyToTrans(DBin, DBOut, "weight", "index", "msre_desc", AbbrevTrans, AbbrevItems)
   'ok gave up on the general  stuff and just used specific routines
   Call CopyRecipes(DBin, DBOut, AbbrevTrans, RecipeTrans, RecipeNames)
   Call CopyRecipesItems(DBin, DBOut, RecipeNames, RecipeTrans, AbbrevTrans)
   
   Call CopyMealPlans(DBin, DBOut, MealPlans, MealPlanNames, TUser)
   Call CopyMeals(DBin, DBOut, MealPlans, MealPlanNames, MealNameTrans, MealNames, TUser)
   Call CopyMealIngred(DBin, DBOut, MealNames, MealNameTrans, AbbrevTrans)
   Call LoadMealDays(DBin, DBOut, UserDate, UserDate, TUser, MealNameTrans)
   Call ExerciseDays(DBin, DBOut, Extrans, UserDate, TUser)
  
  
  
errhandl:
If DoDebug And Err.Number <> 0 Then
  Debug.Print Err.Description
  Stop
  Resume
End If
  Call frmMain.RefreshDay
  Call frmMain.MakeMealList
  ReadScript = UserDate
  

End Function

Private Function CopyTable(DBin As Database, DBOut As Database, Table As String, SearchCol As String, OverWrite As Boolean)
   On Error GoTo eout
   Dim Rs_In As Recordset
   Dim RS_Out As Recordset
   Dim i As Long
   
   Set RS_Out = DBOut.OpenRecordset("select * from " & Table & ";", dbOpenDynaset)
   While Not RS_Out.EOF
     Err.Clear
     Set Rs_In = DB.OpenRecordset("select * from " & Table & _
     " where " & SearchCol & "='" & Replace(RS_Out(SearchCol), "'", "''") & "';", dbOpenDynaset)
     If Err.Number = 91 Then GoTo eout
     If Rs_In.EOF Then
       On Error Resume Next
       Rs_In.AddNew
     Else
       Rs_In.Edit
     End If
       For i = 1 To Rs_In.Fields.Count - 1
       Debug.Print Rs_In(i).DataUpdatable
          Rs_In(i) = RS_Out(Rs_In(i).Name)
       Next i
       Rs_In.Update
     'End If
     RS_Out.MoveNext
   Wend
   
eout:

   Set RS_Out = Nothing
   Set Rs_In = Nothing
End Function
Private Function CopyDailyLog(DBin As Database, DBOut As Database, Table As String, SearchCol As String, OverWrite As Boolean)
   On Error GoTo eout
   Dim Rs_In As Recordset
   Dim RS_Out As Recordset
   Dim i As Long
   
   Set RS_Out = DBOut.OpenRecordset("select * from " & Table & ";", dbOpenDynaset)
   While Not RS_Out.EOF
     Err.Clear
     Set Rs_In = DB.OpenRecordset("select * from " & Table & _
     " where user='" & Replace(RS_Out(SearchCol), "'", "''") & "' and date=#" & FixDate(RS_Out("date")) & "#;", dbOpenDynaset)
     If Err.Number = 91 Then GoTo eout
     If Rs_In.EOF Then
       On Error Resume Next
       Rs_In.AddNew
     Else
       Rs_In.Edit
     End If
     For i = 1 To Rs_In.Fields.Count - 1
       
          Rs_In(i) = RS_Out(Rs_In(i).Name)
     Next i
     Rs_In.Update
     
     RS_Out.MoveNext
   Wend
   
eout:

   Set RS_Out = Nothing
   Set Rs_In = Nothing
End Function
   
Private Function CopyTableWithTrans(DBin As Database, DBOut As Database, TableName As String, IndexName As String, SearchCol As String, trans As Collection, transNames As Collection)
   On Error GoTo errhandl

   Dim i As Long, MaxIt As Long
   'Dim Trans As Collection
   Set trans = New Collection
   Set transNames = New Collection
   
   Dim Rs_In As Recordset, RS_Out As Recordset
   
   
   'get all the info from the incoming table
   Set RS_Out = DBOut.OpenRecordset("select * from " & TableName & ";", dbOpenDynaset)
   'now find the room in the existing table
   Set Rs_In = DBin.OpenRecordset("select max(" & IndexName & ") as MAXit from " & TableName & ";", dbOpenDynaset)
   MaxIt = DoMax(Rs_In)
   
   While Not RS_Out.EOF
      Set Rs_In = DB.OpenRecordset("select * from " & TableName & _
      " where " & SearchCol & "='" & Replace(RS_Out(SearchCol), "'", "''") & "';", dbOpenDynaset)
      
      If Rs_In.EOF And Rs_In.BOF Then
         Rs_In.AddNew
         For i = 0 To Rs_In.Fields.Count - 1
             Rs_In(i) = RS_Out(Rs_In.Fields(i).Name)
         Next i
         Rs_In(IndexName) = MaxIt
         trans.Add "_" & MaxIt, "_" & RS_Out("index")
         transNames.Add "_" & RS_Out("index")
         Rs_In.Update
         MaxIt = MaxIt + 1
         DoEvents
      Else
         trans.Add "_" & Rs_In("index"), "_" & RS_Out("index")
         transNames.Add "_" & RS_Out("index")
      End If
      RS_Out.MoveNext
   Wend
   Set RS_Out = Nothing: Set Rs_In = Nothing
   Exit Function
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
End Function
   
Private Function CopyToTrans(DBin As Database, DBOut As Database, TableName As String, TransCol As String, SearchCol As String, trans As Collection, transNames As Collection)
  On Error GoTo errhandl
   'on error Resume Next
   Dim junk As String, i As Long, j As Long
   Dim RS_Out As Recordset, Rs_In As Recordset
   Dim junk2 As String, junk3 As String
   
   For i = 1 To transNames.Count
      junk = transNames(i)
      Set RS_Out = DBOut.OpenRecordset("Select * from " & TableName & " where " & TransCol & "=" & Replace(junk, "_", "") & ";", dbOpenDynaset)
      junk2 = "Select * from " & TableName & " where " & TransCol & "=" & Replace(trans(junk), "_", "")
      'Set WeightIn = DBin.OpenRecordset("select * from weight where index=" & AIndexs(i) & ";", dbOpenDynaset)
      If Not RS_Out Is Nothing Then
      While Not RS_Out.EOF
         junk3 = junk2 & " and " & SearchCol & "='" & Replace(RS_Out(SearchCol), "'", "''") & "';"
         Set Rs_In = DBin.OpenRecordset(junk3, dbOpenDynaset)
         If Rs_In.EOF Then
            Rs_In.AddNew
         Else
            Rs_In.Delete
            Rs_In.AddNew
         End If
         On Error Resume Next
         For j = 0 To Rs_In.Fields.Count - 1
            Rs_In(j) = RS_Out(Rs_In.Fields(j).Name)
         Next j
         On Error GoTo errhandl
         Rs_In(TransCol) = Val(Replace(trans(junk), "_", ""))
         Rs_In.Update
         
         RS_Out.MoveNext
         DoEvents
      Wend
      End If
   Next i
   Set Rs_In = Nothing
   Set RS_Out = Nothing
   Exit Function
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
   
End Function
   
Private Function CopyRecipes(DBin As Database, DBOut As Database, AbbrevTrans As Collection, trans As Collection, transNames As Collection)
  On Error GoTo errhandl
   Dim RecipeOut As Recordset, RecipeIn As Recordset
   Dim MaxIt As Long, i As Long
   
   Set trans = New Collection
   Set transNames = New Collection
   Set RecipeOut = DBOut.OpenRecordset("select * from recipesindex;", dbOpenDynaset)
   
   Set RecipeIn = DBin.OpenRecordset("select max(recipeid) as maxit from recipesindex;", dbOpenDynaset)
   MaxIt = DoMax(RecipeIn)
   
   While Not RecipeOut.EOF
      Set RecipeIn = DB.OpenRecordset("select * from recipesindex where recipename ='" & Replace(RecipeOut("recipename"), "'", "''") & "';", dbOpenDynaset)
     
      If RecipeIn.EOF And RecipeIn.BOF Then
         
         RecipeIn.AddNew
         For i = 0 To RecipeIn.Fields.Count - 1
            RecipeIn(i) = RecipeOut(RecipeIn.Fields(i).Name)
         Next i
         RecipeIn("recipeid") = MaxIt
         RecipeIn("abbrevid") = AbbrevTrans("_" & RecipeOut("abbrevid"))
         trans.Add MaxIt, "_" & MaxIt, "_" & RecipeOut("recipeid")
         transNames.Add "_" & RecipeOut("recipeid")
         RecipeIn.Update
         MaxIt = MaxIt + 1
      Else
         trans.Add "_" & RecipeIn("recipeid"), "_" & RecipeOut("recipeid")
         transNames.Add "_" & RecipeOut("recipeid")
      End If
      RecipeOut.MoveNext
   Wend
   Set RecipeOut = Nothing: Set RecipeIn = Nothing
   
   Exit Function
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
   
End Function
   
Private Function DoMax(RS As Recordset) As Long
   If Not IsNull(RS("Maxit")) Then
     DoMax = RS("maxit") + 1
   Else
     DoMax = 0
   End If
End Function

Private Function CopyRecipesItems(DBin As Database, DBOut As Database, RecipeItems As Collection, trans As Collection, AbbrevTrans As Collection)
On Error GoTo errhandl
   Dim RecipesIn As Recordset
   Dim RecipesOut As Recordset
   Dim MaxIt As Long, i As Long, j As Long
   Dim junk As String
   Set RecipesIn = DBin.OpenRecordset("select max(id) as maxit from recipes;", dbOpenDynaset)
   MaxIt = DoMax(RecipesIn)
   
   For i = 1 To RecipeItems.Count
      junk = RecipeItems(i)
      Set RecipesOut = DBOut.OpenRecordset("select * from recipes where recipeid=" & Replace(junk, "_", "") & ";", dbOpenDynaset)
      While Not RecipesOut.EOF
         Set RecipesIn = DBin.OpenRecordset("select * from recipes where recipeid = " & Replace(trans(junk), "_", "") & ";", dbOpenDynaset)
         While RecipesIn.EOF
           RecipesIn.Delete
           RecipesIn.MoveNext
         Wend
         RecipesIn.AddNew
         For j = 0 To RecipesIn.Fields.Count - 1
           RecipesIn(j) = RecipesOut(RecipesIn.Fields(j).Name)
         Next j
         RecipesIn("id") = MaxIt
         MaxIt = MaxIt + 1
         RecipesIn("recipeid") = Val(Replace(trans(junk), "_", ""))
         RecipesIn("itemid") = Val(Replace(AbbrevTrans("_" & RecipesOut("itemid")), "_", ""))
         RecipesIn.Update
         RecipesOut.MoveNext
      Wend
   Next i
  Set RecipesIn = Nothing: Set RecipesOut = Nothing
     Exit Function
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If

End Function

Private Sub CopyMealPlans(DBin As Database, DBOut As Database, trans As Collection, transNames As Collection, USER As String)
On Error GoTo errhandl
  Dim MealplanIn As Recordset, MealPlanOut As Recordset
  Dim MaxIt As Long, PlanMax As Long, i As Long
  Set trans = New Collection
  Set transNames = New Collection
  
  Dim PlanID As Long, MaxIt2 As Long
  Set MealPlanOut = DBOut.OpenRecordset("select * from mealplanner where mealid=-1;", dbOpenDynaset)
  If MealPlanOut("planid") = 0 Then
    Set MealPlanOut = DBOut.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
    While Not MealPlanOut.EOF
       MealPlanOut.Edit
       MealPlanOut("planid") = 0
       MealPlanOut.Update
       MealPlanOut.MoveNext
    Wend
    Set MealPlanOut = DBOut.OpenRecordset("select * from mealplanner where mealid=-1;", dbOpenDynaset)
  End If
  Set MealplanIn = DBin.OpenRecordset("select max(index) as maxit from mealplanner;", dbOpenDynaset)
  MaxIt = DoMax(MealplanIn)
  
  Set MealplanIn = DBin.OpenRecordset("select max(planid) as maxit from mealplanner;", dbOpenDynaset)
  PlanMax = DoMax(MealplanIn)
  If PlanMax < 2 Then PlanMax = 2
  
  While Not MealPlanOut.EOF
     
     If USER = "" Then
        Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealname='" & MealPlanOut("mealname") & "' and user='" & MealPlanOut("user") & "';", dbOpenDynaset)
     Else
        Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealname='" & MealPlanOut("mealname") & "' and user='" & USER & "';", dbOpenDynaset)
     End If
     If MealplanIn.EOF Then
        MealplanIn.AddNew
        On Error Resume Next
        For i = 0 To MealplanIn.Fields.Count - 1
           MealplanIn(i) = MealPlanOut(MealplanIn.Fields(i).Name)
        Next i
        On Error GoTo errhandl
        MealplanIn("index") = MaxIt
        MealplanIn("planid") = PlanMax
        If USER <> "" Then
          MealplanIn("user") = USER
        End If
        MealplanIn.Update
        trans.Add "_" & PlanMax, "_" & MealPlanOut("planid")
        transNames.Add "_" & MealPlanOut("planid")
        MaxIt = MaxIt + 1
        PlanMax = PlanMax + 1
     Else
        trans.Add "_" & MealplanIn("planid"), "_" & MealPlanOut("planid")
        transNames.Add "_" & MealPlanOut("planid")
     End If
     MealPlanOut.MoveNext
  Wend

   Exit Sub
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
   

End Sub
Private Sub CopyMeals(DBin As Database, DBOut As Database, PlanTrans As Collection, PlanItems As Collection, trans As Collection, transNames As Collection, USER As String)
  On Error GoTo errhandl
   Dim MPIn As Recordset, MPout As Recordset
   Dim junk As String
   Dim MaxIt As Long, MealMax As Long, i As Long, j As Long
   
   Set trans = New Collection
   Set transNames = New Collection
   
   Set MPIn = DBin.OpenRecordset("select max(index) as maxit from mealplanner;", dbOpenDynaset)
   MaxIt = DoMax(MPIn)
   Set MPIn = DBin.OpenRecordset("select max(mealid) as maxit from mealplanner;", dbOpenDynaset)
   MealMax = DoMax(MPIn)
   
   
   For i = 1 To PlanItems.Count
      junk = PlanItems(i)
      Set MPout = DBOut.OpenRecordset("select * from mealplanner where planid =" & Replace(junk, "_", "") & ";", dbOpenDynaset)
      While Not MPout.EOF
        If MPout("mealid") <> -1 Then
        
         If USER = "" Then
             Set MPIn = DBin.OpenRecordset("select * from mealplanner where planid = " & Replace(PlanTrans(junk), "_", "") _
                 & " and mealname='" & Replace(Trim$(MPout("Mealname")), "'", "''") & "' and user='" & MPout("USER") & "';", dbOpenDynaset)
         Else
             Set MPIn = DBin.OpenRecordset("select * from mealplanner where planid = " & Replace(PlanTrans(junk), "_", "") _
                 & " and mealname='" & Replace(Trim$(MPout("Mealname")), "'", "''") & "' and user='" & USER & "';", dbOpenDynaset)
         End If
          
         If MPIn.EOF Then
            MPIn.AddNew
            On Error Resume Next
            For j = 0 To MPIn.Fields.Count - 1
              MPIn.Fields(j) = MPout(MPIn.Fields(j).Name)
            Next j
            On Error GoTo errhandl
            MPIn("index") = MaxIt
            MPIn("Mealid") = MealMax
            MPIn("planid") = Val(Replace(PlanTrans(junk), "_", ""))
            If USER <> "" Then MPIn("user") = USER
            MPIn.Update
            trans.Add "_" & MealMax, "_" & MPout("mealid")
            transNames.Add "_" & MPout("mealid")
            MealMax = MealMax + 1
            MaxIt = MaxIt + 1
         Else
            trans.Add "_" & MPIn("MEalid"), "_" & MPout("mealid")
            transNames.Add "_" & MPout("mealid")
         End If
        End If
        MPout.MoveNext
      Wend
      
   Next i
   Set MPout = Nothing
   Set MPIn = Nothing
   Exit Sub
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
End Sub

Private Sub CopyMealIngred(DBin As Database, DBOut As Database, MealNames As Collection, MealTrans As Collection, AbbrevTrans As Collection)
On Error GoTo errhandl
 Dim MDIn As Recordset, MDout As Recordset
 Dim junk As String
 Dim MaxIt As Long, i As Long, j As Long
 Set MDIn = DBin.OpenRecordset("select max(index) as maxit from mealdefinition;", dbOpenDynaset)
 MaxIt = DoMax(MDIn)
 
  For i = 1 To MealNames.Count
    junk = MealNames(i)
    Set MDIn = DBin.OpenRecordset("select * from mealdefinition where mealid=" & Replace(MealTrans(junk), "_", "") & ";", dbOpenDynaset)
    While Not MDIn.EOF
       MDIn.Delete
       MDIn.MoveNext
    Wend
    Set MDout = DBOut.OpenRecordset("select * from mealdefinition where mealid=" & Replace(junk, "_", "") & ";", dbOpenDynaset)
    While Not MDout.EOF
      MDIn.AddNew
      On Error Resume Next
      For j = 0 To MDout.Fields.Count - 1
        MDIn.Fields(j) = MDout(MDIn.Fields(j).Name)
      Next j
      On Error GoTo errhandl
      MDIn("index") = MaxIt
      MDIn("abbrevid") = Val(Replace(AbbrevTrans("_" & MDout("abbrevid")), "_", ""))
      MDIn("mealid") = Val(Replace(MealTrans(junk), "_", ""))
      MDIn.Update
      MaxIt = MaxIt + 1
      MDout.MoveNext
    Wend
  Next i
  Set MDIn = Nothing: Set MDout = Nothing
 Exit Sub
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If
   
  
  
End Sub

Private Sub LoadMealDays(DBin As Database, DBOut As Database, UserDate As Date, StartD As Date, Username As String, MealTrans As Collection)
   Dim MDIn As Recordset, MDout As Recordset
   Set MDout = DBOut.OpenRecordset("select * from meals;", dbOpenDynaset)
   Dim sDate As String, sParts() As String
   If Not MDout.EOF Then
      If Year(UserDate) = 1900 Then
         sDate = InputBox("Please enter day you wish to start this plan. (mm/dd/yyyy)", "Loading Plan", Today)
         sParts = Split(sDate, "/")
         StartD = IsoDateStringString(sParts(0), sParts(1), sParts(2))
      Else
         StartD = UserDate
      End If
      
   Else
      Exit Sub
   End If
   If Username <> "" Then
      Set MDIn = DBin.OpenRecordset("select * from meals where entrydate>=#" & FixDate(StartD) & "# and user='" & Username & "';", dbOpenDynaset)
      While Not MDIn.EOF
         MDIn.Delete
         MDIn.MoveNext
      Wend
   Else
      
      While Not MDout.EOF
         Set MDIn = DBin.OpenRecordset("select * from meals where entrydate>=#" & FixDate(StartD) & "# and user='" & MDout("user") & "';", dbOpenDynaset)
         While Not MDIn.EOF
           MDIn.Delete
           MDIn.MoveNext
         Wend
         MDout.MoveNext
      Wend
      Set MDIn = DBin.OpenRecordset("select * from meals where entrydate>=#" & FixDate(StartD) & "#;", dbOpenDynaset)
      MDout.MoveFirst
   End If

       
   
   Dim MinD As Date
   MinD = Today
   While Not MDout.EOF
      If MDout("entrydate") < MinD Then MinD = MDout("entrydate")
      MDout.MoveNext
   Wend
   MDout.MoveFirst
   On Error GoTo errhandl
   Dim j As Long, junk As String, T As String
   T = CurrentUser.Username
   
   While Not MDout.EOF
      j = Abs(DateDiff("d", MDout("entrydate"), MinD))
      'this will probably not work for the update version.  need to just enter the data for the update version
      On Error Resume Next
      Err.Clear
      junk = ""
      junk = Val(Replace(MealTrans("_" & MDout("Mealid")), "_", ""))
      If Err.Number <> 0 Then
         junk = MDout("mealid")
      End If
      If Username = "" Then
         CurrentUser.Username = MDout("user")
      Else
         CurrentUser.Username = Username
      End If
      
      Call frmMain.FlexDiet.DropMeal("~~~" & " " & "~~~" & junk, _
      True, DateAdd("d", j, StartD), False, MDout("mealnumber"))
      MDout.MoveNext
   Wend
   MDout.Close
   MDIn.Close
   
errhandl:
  Set MDout = Nothing
  Set MDIn = Nothing
   CurrentUser.Username = T
End Sub

Public Function firstSunday(InDate As Date) As Date
'on error Resume Next
  Dim i As Long
  i = Weekday(InDate, vbSunday) - 1
  firstSunday = DateAdd("d", -1 * i, InDate)
End Function
  
Private Sub ExerciseDays(DBin As Database, DBOut As Database, Extrans As Collection, StartD As Date, USER As String)
On Error GoTo errhandl
  Dim ELout As Recordset, ELIn As Recordset, msD As Date
  Dim sDate As String, sParts() As String
      If Year(StartD) = 1900 Then
         sDate = InputBox("Please enter day you wish to start this plan. (mm/dd/yyyy)", "Loading Plan", Today)
         sParts = Split(sDate, "/")
         StartD = DateHandler.IsoDateStringString(sParts(0), sParts(1), sParts(2))
         
      End If
  StartD = firstSunday(StartD)
  
  Set ELout = DBOut.OpenRecordset("select min(week) as minit from exerciselog;", dbOpenDynaset)
  'on error Resume Next
  Err.Clear
  If IsNull(ELout("minit")) Then
     Exit Sub
  End If
  msD = ELout("minit")
  
  'on error Resume Next
  Set ELIn = DBin.OpenRecordset("select max(index) as maxit from exerciselog;", dbOpenDynaset)
  Dim MaxIt As Long
  MaxIt = DoMax(ELIn)
  
  If USER = "" Then
     Set ELIn = DBin.OpenRecordset("select * from exerciselog where week>=#" & FixDate(StartD) & "#;", dbOpenDynaset)
  Else
     Set ELIn = DBin.OpenRecordset("select * from exerciselog where user='" & USER & "' and week>=#" & FixDate(StartD) & "#;", dbOpenDynaset)
  End If
  If Not (ELIn.EOF And ELIn.BOF) Then
     Dim ret2 As VbMsgBoxResult
     ret2 = MsgBox("Do you wish to overwrite existing exercise entries?" & vbCrLf & "(Choose 'No' to append this plan to your current exercise schedule)", vbYesNo)
     If ret2 = vbYes Then
       If USER = "" Then
           While Not ELout.EOF
              Set ELIn = DBin.OpenRecordset("select * from exerciselog where user='" & ELout("user") & "' and week>=#" & FixDate(StartD) & "#;", dbOpenDynaset)
              While Not ELIn.EOF
                 ELIn.Delete
                 ELIn.MoveNext
              Wend
              ELout.MoveNext
           Wend
           ELout.MoveFirst
           Set ELIn = DBin.OpenRecordset("select * from exerciselog where week>=#" & FixDate(StartD) & "#;", dbOpenDynaset)
       Else
            While Not ELIn.EOF
              ELIn.Delete
              ELIn.MoveNext
            Wend
       End If
     End If
  End If
  
  Set ELout = DBOut.OpenRecordset("select * from exerciselog;", dbOpenDynaset)
  Dim i As Long, j As Long
  While Not ELout.EOF
     ELIn.AddNew
     For i = 0 To ELIn.Fields.Count - 1
        ELIn(i) = ELout(ELIn.Fields(i).Name)
     Next i
     ELIn("index") = MaxIt
     j = Abs(DateDiff("d", msD, ELout("week")))
     ELIn("week") = DateAdd("d", j, StartD) 'FixDate(DateAdd("d", j, StartD))
     ELIn("exerciseid") = Val(Replace(Extrans("_" & ELout("exerciseid")), "_", ""))
     If USER = "" Then
        ELIn("user") = ELout("user")
     Else
        ELIn("user") = USER
     End If
     MaxIt = MaxIt + 1
     ELIn.Update
     ELout.MoveNext
  Wend
  Set ELIn = Nothing: Set ELout = Nothing
  
   Exit Sub
errhandl:
   If DoDebug Then
      Debug.Print Err.Description
      Stop
      Resume
   Else
      Resume Next
   End If

End Sub

Public Sub SaveWeek(OutFilename As String, DB As Database, USER As String)
  Dim OriginalFilename As String
  Dim DBOut As Database
  Call CopyFile(App.path & "\resources\tempscript.mdb", OutFilename, False)
  Set DBOut = OpenDatabase(OutFilename)
  Dim Abbrevs As Collection, EXs As Collection
  Set EXs = New Collection
  Set Abbrevs = New Collection
  SaveTable DB, DBOut, "dailylog", "id", USER
  SaveTable DB, DBOut, "daysinfo", "id", USER, "itemid", Abbrevs
  SaveTable DB, DBOut, "units", "index", ""
  SaveTable DB, DBOut, "exerciselog", "index", USER, , , "exerciseid", EXs
  SaveTable DB, DBOut, "foodgroups", "index"
  SaveTable DB, DBOut, "ideals", "index", USER
  SaveTable DB, DBOut, "mealdefinition", "index", , "abbrevid", Abbrevs
  SaveTable DB, DBOut, "mealplanner", "index", USER
  SaveTable DB, DBOut, "meals", "id", USER
  SaveTable DB, DBOut, "profiles", "", USER
  SaveTable DB, DBOut, "recipes", "id", "", "itemid", Abbrevs
  SaveTable DB, DBOut, "recipesindex", "", "", "abbrevid", Abbrevs
  
  Dim rsSRC As Recordset, RS As Recordset, j As Long, i As Long
  'get all te custom entries
  Set RS = DB.OpenRecordset("select * from abbrev where ndb_no is null;", dbOpenDynaset)
  While Not RS.EOF
  '  Abbrevs.Add rs("index") & ""
    RS.MoveNext
  Wend
  Set RS = Nothing
  For j = 1 To Abbrevs.Count
     Set rsSRC = DB.OpenRecordset("select * from abbrev where index=" & Abbrevs(j) & ";", dbOpenDynaset)
     Set RS = DBOut.OpenRecordset("select * from abbrev where index=" & Abbrevs(j) & ";", dbOpenDynaset)
     If RS.EOF Then
         RS.AddNew
         On Error Resume Next
         For i = 0 To rsSRC.Fields.Count - 1
            RS(rsSRC.Fields(i).Name) = rsSRC(i)
         Next i
         RS.Update
      End If
      RS.Close
      Set RS = Nothing
  Next j
  RS.Close
  rsSRC.Close
  Set RS = Nothing
  Set rsSRC = Nothing
  
  
  For j = 1 To Abbrevs.Count
     Set rsSRC = DB.OpenRecordset("select * from weight where index=" & Abbrevs(j) & ";", dbOpenDynaset)
     Set RS = DBOut.OpenRecordset("select * from weight where index=" & Abbrevs(j) & ";", dbOpenDynaset)
     If RS.EOF Then
         While Not rsSRC.EOF
             RS.AddNew
             On Error Resume Next
             For i = 0 To rsSRC.Fields.Count - 1
                RS(rsSRC.Fields(i).Name) = rsSRC(i)
             Next i
             RS.Update
             rsSRC.MoveNext
         Wend
     End If
     RS.Close
     Set RS = Nothing
  Next j
  RS.Close
  rsSRC.Close
  Set RS = Nothing
  Set rsSRC = Nothing
  
  
  For j = 1 To EXs.Count
     Set rsSRC = DB.OpenRecordset("select * from abbrevexercise where index=" & EXs(j) & ";", dbOpenDynaset)
     Set RS = DBOut.OpenRecordset("select * from abbrevexercise where index=" & EXs(j) & ";", dbOpenDynaset)
     If RS.EOF Then
         RS.AddNew
         On Error Resume Next
         For i = 0 To rsSRC.Fields.Count - 1
            RS(rsSRC.Fields(i).Name) = rsSRC(i)
         Next i
         RS.Update
         rsSRC.MoveNext
     End If
  Next j
  RS.Close
  rsSRC.Close
  Set RS = Nothing
  Set rsSRC = Nothing
End Sub
  
Private Sub SaveTable(DBSrc As Database, DBOut As Database, Table As String, indexCol As String, Optional USER As String = "", Optional AbbrevCol As String = "", Optional Abbrevs As Collection = Nothing, Optional ExCol As String = "", Optional EXs As Collection = Nothing)
  Dim rsSRC As Recordset, RS As Recordset
  Dim i As Long, cc As Long
  If USER <> "" Then
     Set rsSRC = DBSrc.OpenRecordset("select * from " & Table & " where user='" & USER & "';", dbOpenDynaset)
  Else
     Set rsSRC = DBSrc.OpenRecordset("select * from " & Table & ";", dbOpenDynaset)
  End If
  Set RS = DBOut.OpenRecordset("select * from " & Table & ";", dbOpenDynaset)
  cc = 1
  While Not rsSRC.EOF
    RS.AddNew
    On Error Resume Next
    For i = 0 To rsSRC.Fields.Count - 1
       RS(rsSRC.Fields(i).Name) = rsSRC(i)
    Next i
    If indexCol <> "" Then RS(indexCol) = cc
    If AbbrevCol <> "" Then Abbrevs.Add RS(AbbrevCol) & ""
    If ExCol <> "" Then EXs.Add RS(ExCol) & ""
    RS.Update
    rsSRC.MoveNext
    cc = cc + 1
  Wend
  RS.Close
  rsSRC.Close
  Set RS = Nothing
  Set rsSRC = Nothing
End Sub
