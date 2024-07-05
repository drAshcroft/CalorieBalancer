Attribute VB_Name = "REadScriptMod"
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
Private Declare Function CopyFile Lib "kernel32" _
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
    Dim name As String, nameLen As Long
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
        name = Space$(nameLen)
        dataLen = 4096
        ReDim binValue(0 To dataLen - 1) As Byte
        If RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, valueType, binValue(0), dataLen) Then Exit For
        result(0, Index) = Left$(name, nameLen)
        
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


    On Error GoTo Err_Proc
Dim bPath As String

    bPath = BrowserPath
    If bPath <> "Invalid" Then
        Shell bPath & Chr(32) & strURL, Style
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



Public Sub UpdateInternetAbbrev(Matrix, today As Date, LoadIntoDay As Boolean, Optional RecipeID As Long)


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
          Set temp = DB.OpenRecordset("SELECT * FROM DaysInfo WHERE (((DaysInfo.Date)=#" & today & "#) AND (DaysInfo.user='" & CurrentUser.Username & "')) ORDER BY daysinfo.order;", dbOpenDynaset)
          While Not temp.EOF
             temp.Delete
             temp.MoveNext
          Wend
          For i = 0 To UBound(Matrix, 2)
           If Matrix(3, i) <> "" Then
              temp.AddNew
              temp("date") = today
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
Public Function ReadScript(Filename As String, tUser As String, NewPlan As Boolean, Optional UserDate As Date = "1/1/1900") As Date
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

  'On Error GoTo errhandl
  On Error Resume Next
   Dim DBin As Database, i As Long, j As Long
   Set DBin = OpenDatabase(Filename)
   
   
   Dim FGIN As Recordset
   Dim FGOUT As Recordset
   
   Set FGIN = DBin.OpenRecordset("select * from FoodGroups;", dbOpenDynaset)
   While Not FGIN.EOF
     Err.Clear
     Set FGOUT = DB.OpenRecordset("select * from FoodGroups where category='" & Replace(FGIN("category"), "'", "''") & "';", dbOpenDynaset)
     If Err.Number = 91 Then GoTo eout
     If FGOUT.EOF Then
       FGOUT.AddNew
       For i = 1 To FGOUT.Fields.Count - 1
          FGOUT(i) = FGIN(FGOUT(i).name)
       Next i
       FGOUT.Update
     End If
     FGIN.MoveNext
   Wend
eout:
   Set FGOUT = Nothing
   Set FGIN = Nothing
   On Error GoTo errhandl
   Dim Profiles As Recordset
   Set Profiles = DB.OpenRecordset("Select * from profiles where user ='" & tUser & "';", dbOpenDynaset)
   
   
   Dim AbbrevTrans As Collection
   Set AbbrevTrans = New Collection
   
   Dim AbbrevIn As Recordset, Abbrev As Recordset
   Dim AIndexs(), cc As Long, MaxIt As Long
   Dim NewIndexs() As Long
   Set AbbrevIn = DBin.OpenRecordset("select * from abbrev;", dbOpenDynaset)
   Set Abbrev = DB.OpenRecordset("select max(index)  as MAXit from abbrev;", dbOpenDynaset)
   MaxIt = Abbrev("maxit") + 1
   While Not AbbrevIn.EOF
      Set Abbrev = DB.OpenRecordset("select * from abbrev where foodname='" & Replace(AbbrevIn("foodname"), "'", "''") & "';", dbOpenDynaset)
      If Abbrev.EOF And Abbrev.BOF Then
         ReDim Preserve AIndexs(cc)
         ReDim Preserve NewIndexs(cc)
         AIndexs(cc) = AbbrevIn("index")
         cc = cc + 1
         Abbrev.AddNew
         On Error Resume Next
         For i = 0 To Abbrev.Fields.Count - 1
             Abbrev(i) = AbbrevIn(Abbrev.Fields(i).name)
         Next i
         Abbrev("index") = MaxIt
         NewIndexs(cc - 1) = MaxIt
         AbbrevTrans.Add MaxIt, "a" & AbbrevIn("index")
         Abbrev.Update
         MaxIt = MaxIt + 1
         DoEvents
      Else
         AbbrevTrans.Add Val(Abbrev("index")), "a" & AbbrevIn("index")
      End If
      AbbrevIn.MoveNext
   Wend
   On Error GoTo errhandl
   Set Abbrev = Nothing: Set AbbrevIn = Nothing
   Dim Weight As Recordset, WeightIn As Recordset
   Set Weight = DB.OpenRecordset("Select * from weight;", dbOpenDynaset)
   cc = cc - 1
   For i = 0 To cc
      Set WeightIn = DBin.OpenRecordset("select * from weight where index=" & AIndexs(i) & ";", dbOpenDynaset)
      While Not WeightIn.EOF
         Weight.AddNew
         For j = 0 To Weight.Fields.Count - 1
            Weight(j) = WeightIn(Weight.Fields(j).name)
         Next j
         Weight("index") = NewIndexs(i)
         Weight.Update
         WeightIn.MoveNext
         DoEvents
      Wend
   Next i
   Set Weight = Nothing
   Set WeightIn = Nothing
   Dim RecipesIndex As Recordset, RecipeIn As Recordset
   Dim recipeTrans As Collection
   Set recipeTrans = New Collection
   ReDim AIndexs(0): ReDim NewIndexs(0): cc = 0
   
   Set RecipeIn = DBin.OpenRecordset("select * from recipesindex;", dbOpenDynaset)
   Set RecipesIndex = DB.OpenRecordset("select max(recipeid) as maxit from recipesindex;", dbOpenDynaset)
   MaxIt = Val(RecipesIndex("maxit") & "") + 1
   While Not RecipeIn.EOF
      Set RecipesIndex = DB.OpenRecordset("select * from recipesindex where recipename ='" & Replace(RecipeIn("recipename"), "'", "''") & "';", dbOpenDynaset)
     
      If RecipesIndex.EOF And RecipesIndex.BOF Then
         ReDim Preserve NewIndexs(cc)
         ReDim Preserve AIndexs(cc)
         AIndexs(cc) = RecipeIn("recipeid")
         NewIndexs(cc) = MaxIt
         recipeTrans.Add MaxIt, "r" & AIndexs(cc)
         RecipesIndex.AddNew
         For i = 0 To RecipesIndex.Fields.Count - 1
            RecipesIndex(i) = RecipeIn(RecipesIndex.Fields(i).name)
         Next i
         RecipesIndex("recipeid") = MaxIt
         RecipesIndex("abbrevid") = AbbrevTrans("a" & RecipeIn("abbrevid"))
         RecipesIndex.Update
         cc = cc + 1
         MaxIt = MaxIt + 1
      Else
         recipeTrans.Add Val(RecipesIndex("recipeid")), "r" & RecipeIn("recipeid")
      End If
      RecipeIn.MoveNext
   Wend
   cc = cc - 1
   Set RecipesIndex = Nothing: Set RecipeIn = Nothing
   Dim Recipes As Recordset
   Set Recipes = DB.OpenRecordset("select max(id) as maxit from recipes;", dbOpenDynaset)
   MaxIt = Val(Recipes("maxit") & "") + 1
   Set Recipes = DB.OpenRecordset("select * from recipes;", dbOpenDynaset)
   
   For i = 0 To cc
      Set RecipeIn = DBin.OpenRecordset("select * from recipes where recipeid = " & AIndexs(i) & ";", dbOpenDynaset)
      While Not RecipeIn.EOF
         Recipes.AddNew
         For j = 0 To Recipes.Fields.Count - 1
           Recipes(j) = RecipeIn(Recipes.Fields(j).name)
         Next j
         Recipes("id") = MaxIt
         MaxIt = MaxIt + 1
         Recipes("recipeid") = NewIndexs(i)
         Recipes("itemid") = AbbrevTrans("a" & Recipes("itemid"))
         Recipes.Update
         RecipeIn.MoveNext
      Wend
   Next i
  Set RecipeIn = Nothing: Set Recipes = Nothing
  
  Dim MealplanIn As Recordset, MealPlan As Recordset
  Dim MealTrans As Collection
  Set MealTrans = New Collection
  Dim PlanID As Long, MaxIt2 As Long
  Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealid=-1;", dbOpenDynaset)
  
  Dim WriteNewPlan As Boolean
  

  'check if the mealplan is already loaded and then
  'make room for it if it is not
  If Not MealplanIn.EOF Then
      Dim junk As String, junk1 As String
       junk1 = Trim$(MealplanIn("url"))
       junk = Replace(junk1, "~FilePath~", App.path, , , vbTextCompare)
       Dim FSO As New FileSystemObject
       If junk <> "" Then
         If Mid$(junk, 2, 1) = ":" Then
            If FSO.FileExists(junk) Then OpenURL junk
         Else
            OpenURL junk
         End If
       End If
       Set MealPlan = DB.OpenRecordset("select * from mealplanner where user='" _
       & CurrentUser.Username & "' and " _
       & "mealname ='" & Replace(Trim$(MealplanIn("Mealname")), "'", "''") & "';" _
       , dbOpenDynaset)
       
       If Not MealPlan.EOF Then
          Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealid<>-1;", dbOpenDynaset)
          PlanID = MealPlan("planid")
          If MealPlan("calories") = 0 Then
                    Profiles.Edit
                    Profiles("dietplans") = PlanID
                    Profiles.Update
           Else
                    Profiles.Edit
                    Profiles("exerciseplans") = PlanID
                    Profiles.Update
           End If
          WriteNewPlan = False
       Else
          Set MealPlan = DB.OpenRecordset("select max(planid) as maxit from mealplanner;", dbOpenDynaset)
          PlanID = MealPlan("maxit") + 1
          If PlanID < 2 Then PlanID = 2
          Set MealplanIn = DBin.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
          WriteNewPlan = True
       End If
   End If
 'now make room for any meals that may be added
   
   Set MealPlan = DB.OpenRecordset("select max(index) as maxit from mealplanner;", dbOpenDynaset)
   MaxIt2 = MealPlan("maxit") + 1
   Set MealPlan = DB.OpenRecordset("select max(mealid) as maxit from mealplanner;", dbOpenDynaset)
   MaxIt = MealPlan("maxit") + 1
      
  'do final clean up to avoid any preset meals
  If MaxIt = 0 Then MaxIt = 1
  If PlanID < 0 Then PlanID = 1

  While Not MealplanIn.EOF
     Set MealPlan = DB.OpenRecordset("Select * from mealplanner where user='" & CurrentUser.Username _
     & "' and planid = " & PlanID _
     & " and mealname ='" & Replace(Trim$(MealplanIn("mealname")), "'", "''") _
     & "';", dbOpenDynaset)
     If MealPlan.EOF Then
         MealPlan.AddNew
         On Error Resume Next
         For i = 0 To MealPlan.Fields.Count - 1
            MealPlan(i) = MealplanIn(MealPlan.Fields(i).name)
         Next i
         On Error GoTo errhandl
         MealPlan("index") = MaxIt2
         MaxIt2 = MaxIt2 + 1
         MealPlan("planid") = PlanID
         MealPlan("user") = CurrentUser.Username
        
         If MealplanIn("mealid") <> -1 Then
             MealTrans.Add MaxIt, "m" & MealplanIn("mealid")
             MealPlan("Mealid") = MaxIt
             MaxIt = MaxIt + 1
         Else
             MealPlan("planfile") = Replace(OriginalFilename, App.path, "~FilePath~", , , vbTextCompare)
             If MealPlan("calories") = 0 Then
                    Profiles.Edit
                    Profiles("dietplans") = PlanID
                    Profiles.Update
             Else
                    Profiles.Edit
                    Profiles("exerciseplans") = PlanID
                    Profiles.Update
             End If
         End If
         MealPlan("MealName") = Trim$(MealPlan("Mealname"))
         MealPlan.Update
     Else
         'otherwise just make a reference to it
         'If Not WriteNewPlan Then
           On Error Resume Next
             MealTrans.Remove "m" & MealplanIn("mealid")
             MealTrans.Add MealPlan("mealid") & "_s", "m" & MealplanIn("mealid")
         'End If
     End If
     MealplanIn.MoveNext
  Wend
  Set MealPlan = Nothing: Set MealplanIn = Nothing
  
  'now we need to make the meal definitions
  Dim MealDef As Recordset, MealDefIn As Recordset

  Set MealDef = DB.OpenRecordset("select max(index) as maxit from mealdefinition;", dbOpenDynaset)
  MaxIt2 = Val(MealDef("maxit") & "") + 1
  Set MealplanIn = DBin.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
  Set MealPlan = DB.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
  On Error Resume Next
  While Not MealplanIn.EOF
     junk = MealTrans("m" & MealplanIn("mealid"))
     If Right$(junk, 1) <> "s" Then
        MealPlan.AddNew
        For i = 0 To MealPlan.Fields.Count - 1
           MealPlan(i) = MealplanIn(MealPlan.Fields(i).name)
        Next i
        MealPlan("index") = MaxIt2
        MealPlan("mealid") = MealTrans("m" & MealplanIn("mealid"))
        MealPlan("abbrevid") = AbbrevTrans("a" & MealplanIn("abbrevid"))
        MaxIt2 = MaxIt2 + 1
        MealPlan.Update
        
     End If
     MealplanIn.MoveNext
  Wend

  Set MealPlan = Nothing: Set MealplanIn = Nothing
  
JustDays:
  
  Dim Meals As Recordset, Mealsin As Recordset
  
  'MsgBox startD
  Dim startD As Date, msD As Date
  On Error Resume Next
  Set Mealsin = DBin.OpenRecordset("Select min(entrydate) as minit from meals;", dbOpenDynaset)
  'MsgBox "mealsing"
  Err.Clear
  'MsgBox "bex"
  If IsNull(Mealsin("minit")) Then
    GoTo ExerciseSec
  End If
  msD = Mealsin("minit")
  On Error GoTo errhandl
  Set Mealsin = DBin.OpenRecordset("select * from meals;", dbOpenDynaset)
  If Not Mealsin.EOF Then
    If Year(UserDate) = 1900 Then
       startD = InputBox("Please enter day you wish to start this plan. (mm/dd/yyyy)", "Loading Plan", Date)
    Else
       startD = UserDate
    End If
    firstday = startD
    MaxDay = firstday
  End If
  
  
  Set Meals = DB.OpenRecordset("select max(id) as maxit from meals;", dbOpenDynaset)
  MaxIt = Val(Meals("maxit") & "") + 1
  'MsgBox "1"
  Set Meals = DB.OpenRecordset("select * from meals where user='" & tUser & "' and  entrydate >=#" & startD & "#;", dbOpenDynaset)
  If NewPlan And Not Mealsin.EOF Then
      While Not Meals.EOF
         Meals.Delete
         Meals.MoveNext
      Wend
  End If
  'msgbox "2"
  On Error Resume Next
  While Not Mealsin.EOF
'     Meals.AddNew
'     For i = 0 To Meals.Fields.Count - 1
'        Meals(i) = Mealsin(Meals.Fields(i).name)
'     Next i
      j = Abs(DateDiff("d", msD, Mealsin("entrydate")))
'     Meals("entrydate") =
'     If Meals("entrydate") > MaxDay Then MaxDay = Meals("entrydate")
'     Meals("mealid") = Val(MealTrans("m" & Mealsin("Mealid")))
'     Meals("user") = tUser
'     Meals("id") = MaxIt
'     MaxIt = MaxIt + 1
'     Meals.Update
     Call frmMain.FlexDiet.DropMeal("~~~" & " " & "~~~" & Val(MealTrans("m" & Mealsin("Mealid"))), True, DateAdd("d", j, startD), False, Mealsin("mealnumber"))
     Mealsin.MoveNext
  Wend
  Set AbbrevTrans = Nothing
  Set MealTrans = Nothing
  Set recipeTrans = Nothing
  
  'msgbox "3"

   'msgbox "4"
  
  On Error GoTo errhandl
ExerciseSec:


  'load any new exercises
  Dim AbbrevExIn As Recordset, abbrevX As Recordset
  Dim AXtrans As Collection
  Set AXtrans = New Collection
  Set AbbrevExIn = DBin.OpenRecordset("select * from abbrevexercise;", dbOpenDynaset)
  Set abbrevX = DB.OpenRecordset("select max(index) as maxit from abbrevexercise;", dbOpenDynaset)
  MaxIt = Val(abbrevX("maxit") & "") + 1
  While Not AbbrevExIn.EOF
     Set abbrevX = DB.OpenRecordset("select * from abbrevexercise where exercisename ='" & Replace(AbbrevExIn("exercisename"), "'", "''") & "';", dbOpenDynaset)
     If abbrevX.EOF And abbrevX.BOF Then
         abbrevX.AddNew
         For i = 0 To abbrevX.Fields.Count - 1
            abbrevX(i) = AbbrevExIn(abbrevX.Fields(i).name)
         Next i
         abbrevX("index") = MaxIt
         AXtrans.Add MaxIt, "a" & AbbrevExIn("index")
         MaxIt = MaxIt + 1
         abbrevX.Update
     Else
         AXtrans.Add Val(abbrevX("index") & ""), "a" & AbbrevExIn("index")
     End If
     AbbrevExIn.MoveNext
  Wend
  
  Set abbrevX = Nothing: Set AbbrevExIn = Nothing
  Dim ExLog As Recordset, ExlogIn As Recordset
  startD = firstSunday(startD)
  Set ExlogIn = DBin.OpenRecordset("select min(week) as minit from exerciselog;", dbOpenDynaset)
  On Error Resume Next
  Err.Clear
  If IsNull(ExlogIn("minit")) Then
    GoTo ExitSec
  End If
  msD = ExlogIn("minit")
  On Error GoTo errhandl
  Set ExLog = DB.OpenRecordset("select max(index) as maxit from exerciselog;", dbOpenDynaset)
  MaxIt = Val(ExLog("maxit") & "") + 1
  Set ExLog = DB.OpenRecordset("select * from exerciselog where user='" & tUser & "' and week>=#" & startD & "#;", dbOpenDynaset)
  If Not (ExLog.EOF And ExLog.BOF) Then
     Dim ret2 As VbMsgBoxResult
     ret2 = MsgBox("Do you wish to overwrite existing exercise entries?" & vbCrLf & "(Choose 'No' to append this plan to your current exercise schedule)", vbYesNo)
     If ret2 = vbYes Then
        While Not ExLog.EOF
          ExLog.Delete
          ExLog.MoveNext
        Wend
     End If
  End If
  
  Set ExlogIn = DBin.OpenRecordset("select * from exerciselog;", dbOpenDynaset)
  
  While Not ExlogIn.EOF
     ExLog.AddNew
     For i = 0 To ExLog.Fields.Count - 1
        ExLog(i) = ExlogIn(ExLog.Fields(i).name)
     Next i
     ExLog("index") = MaxIt
     j = Abs(DateDiff("d", msD, ExLog("week")))
     ExLog("week") = DateAdd("d", j, startD)
     ExLog("exerciseid") = AXtrans("a" & ExlogIn("exerciseid"))
     ExLog("user") = tUser
     MaxIt = MaxIt + 1
     ExLog.Update
     ExlogIn.MoveNext
  Wend
ExitSec:
  On Error GoTo errhandl
  Set ExLog = Nothing: Set ExlogIn = Nothing
  DBin.Close
  Set DBin = Nothing
  
  
 
  'Call frmMenuPlanner.RefreshMenu
  Call frmMain.RefreshDay
  Call frmMain.MakeMealList
  'OpenURL App.path & "\Resources\plans\pregnant.htm"
  ReadScript = startD
  Exit Function
errhandl:

   'Resume 'todo: need better errhandling here
  Call frmMain.RefreshDay
  'Call frmMenuPlanner.RefreshMenu
  Call frmMain.MakeMealList
End Function

Private Function SimpleCopy(TableName As String, DBin As Database, DB As Database)
  On Error Resume Next
   Dim Profiles As Recordset, ProfIn As Recordset, i As Long
   Set ProfIn = DBin.OpenRecordset("select * from " & TableName & ";", dbOpenDynaset)
   Set Profiles = DB.OpenRecordset("Select * from " & TableName & ";", dbOpenDynaset)
   While Not ProfIn.EOF
     Profiles.AddNew
     For i = 0 To Profiles.Fields.Count - 1
       Profiles.Fields(i) = ProfIn.Fields(Profiles.Fields(i).name)
     Next i
     Profiles.Update
     ProfIn.MoveNext
   Wend
   ProfIn.Close
   Profiles.Close
   Set ProfIn = Nothing
   Set Profiles = Nothing
   

  
End Function
Public Function UpdateScript(Filename As String) As Date
  'On Error GoTo errhandl
   Dim DBin As Database, i As Long, j As Long
   Set DBin = OpenDatabase(Filename)
   
   Dim FGIN As Recordset
   Dim FGOUT As Recordset
   'On Error Resume Next
   Set FGIN = DBin.OpenRecordset("select * from FoodGroups;", dbOpenDynaset)
   While Not FGIN.EOF
     Err.Clear
     Set FGOUT = DB.OpenRecordset("select * from FoodGroups where category='" & Replace(FGIN("category"), "'", "''") & "';", dbOpenDynaset)
     If Err.Number = 91 Then GoTo eout
     If FGOUT.EOF Then
       FGOUT.AddNew
       For i = 1 To FGOUT.Fields.Count - 1
          FGOUT(i) = FGIN(FGOUT(i).name)
       Next i
       FGOUT.Update
     End If
     FGIN.MoveNext
   Wend
eout:
   Set FGOUT = Nothing
   Set FGIN = Nothing
'   On Error GoTo errhandl
   
   Call SimpleCopy("profiles", DBin, DB)
   Call SimpleCopy("ideals", DBin, DB)
   Call SimpleCopy("dailylog", DBin, DB)
   
   
   
   Dim AbbrevTrans As Collection
   Set AbbrevTrans = New Collection
   
   Dim AbbrevIn As Recordset, Abbrev As Recordset
   Dim AIndexs(), cc As Long, MaxIt As Long
   Dim NewIndexs() As Long
   Set AbbrevIn = DBin.OpenRecordset("select * from abbrev;", dbOpenDynaset)
   Set Abbrev = DB.OpenRecordset("select max(index) as MAXit from abbrev;", dbOpenDynaset)
   If IsNull(Abbrev("maxit")) Then
      MaxIt = 1
   Else
      MaxIt = Abbrev("maxit") + 1
   End If
   While Not AbbrevIn.EOF
      Set Abbrev = DB.OpenRecordset("select * from abbrev where foodname='" & Replace(AbbrevIn("foodname"), "'", "''") & "';", dbOpenDynaset)
      If Abbrev.EOF And Abbrev.BOF Then
         ReDim Preserve AIndexs(cc)
         ReDim Preserve NewIndexs(cc)
         AIndexs(cc) = AbbrevIn("index")
         cc = cc + 1
         Abbrev.AddNew
         'on error Resume Next
         For i = 0 To Abbrev.Fields.Count - 1
             Abbrev(i) = AbbrevIn(Abbrev.Fields(i).name)
         Next i
         If AbbrevIn("index") <= -200 Then
            Abbrev("index") = AbbrevIn("index")
            NewIndexs(cc - 1) = AbbrevIn("index")
            AbbrevTrans.Add AbbrevIn("index") + 0, "a" & AbbrevIn("index")
         Else
            Abbrev("index") = MaxIt
            NewIndexs(cc - 1) = MaxIt
            AbbrevTrans.Add MaxIt, "a" & AbbrevIn("index")
            MaxIt = MaxIt + 1
         End If
         Abbrev.Update
         
      Else
         AbbrevTrans.Add Val(Abbrev("index")), "a" & AbbrevIn("index")
      End If
      AbbrevIn.MoveNext
   Wend
   'on error GoTo errhandl
   Set Abbrev = Nothing: Set AbbrevIn = Nothing
   
   Dim Weight As Recordset, WeightIn As Recordset
   Set Weight = DB.OpenRecordset("Select * from weight;", dbOpenDynaset)
   cc = cc - 1
   For i = 0 To cc
      Set WeightIn = DBin.OpenRecordset("select * from weight where index=" & AIndexs(i) & ";", dbOpenDynaset)
      While Not WeightIn.EOF
         Weight.AddNew
         For j = 0 To Weight.Fields.Count - 1
            Weight(j) = WeightIn(Weight.Fields(j).name)
         Next j
         Weight("index") = NewIndexs(i)
         Weight.Update
         WeightIn.MoveNext
      Wend
   Next i
   Set Weight = Nothing
   Set WeightIn = Nothing
   Dim RecipesIndex As Recordset, RecipeIn As Recordset
   Dim recipeTrans As Collection
   Set recipeTrans = New Collection
   ReDim AIndexs(0): ReDim NewIndexs(0): cc = 0
   
   Set RecipeIn = DBin.OpenRecordset("select * from recipesindex;", dbOpenDynaset)
   Set RecipesIndex = DB.OpenRecordset("select max(recipeid) as maxit from recipesindex;", dbOpenDynaset)
   MaxIt = Val(RecipesIndex("maxit") & "") + 1
   While Not RecipeIn.EOF
      Set RecipesIndex = DB.OpenRecordset("select * from recipesindex where recipename ='" & Replace(RecipeIn("recipename"), "'", "''") & "';", dbOpenDynaset)
     
      If RecipesIndex.EOF And RecipesIndex.BOF Then
         ReDim Preserve NewIndexs(cc)
         ReDim Preserve AIndexs(cc)
         AIndexs(cc) = RecipeIn("recipeid")
         NewIndexs(cc) = MaxIt
         recipeTrans.Add MaxIt, "r" & AIndexs(cc)
         RecipesIndex.AddNew
         For i = 0 To RecipesIndex.Fields.Count - 1
            RecipesIndex(i) = RecipeIn(RecipesIndex.Fields(i).name)
         Next i
         RecipesIndex("recipeid") = MaxIt
         RecipesIndex("abbrevid") = AbbrevTrans("a" & RecipeIn("abbrevid"))
         RecipesIndex.Update
         cc = cc + 1
         MaxIt = MaxIt + 1
      Else
         recipeTrans.Add Val(RecipesIndex("recipeid")), "r" & RecipeIn("recipeid")
      End If
      RecipeIn.MoveNext
   Wend
   cc = cc - 1
   Set RecipesIndex = Nothing: Set RecipeIn = Nothing
   Dim Recipes As Recordset
   Set Recipes = DB.OpenRecordset("select max(id) as maxit from recipes;", dbOpenDynaset)
   MaxIt = Val(Recipes("maxit") & "") + 1
   Set Recipes = DB.OpenRecordset("select * from recipes;", dbOpenDynaset)
   
   For i = 0 To cc
      Set RecipeIn = DBin.OpenRecordset("select * from recipes where recipeid = " & AIndexs(i) & ";", dbOpenDynaset)
      While Not RecipeIn.EOF
         Recipes.AddNew
         For j = 0 To Recipes.Fields.Count - 1
           Recipes(j) = RecipeIn(Recipes.Fields(j).name)
         Next j
         Recipes("id") = MaxIt
         MaxIt = MaxIt + 1
         Recipes("recipeid") = NewIndexs(i)
         Recipes("itemid") = AbbrevTrans("a" & Recipes("itemid"))
         Recipes.Update
         RecipeIn.MoveNext
      Wend
   Next i
  Set RecipeIn = Nothing: Set Recipes = Nothing
  
  
  Dim MealplanIn As Recordset, MealPlan As Recordset
  Dim MealTrans As Collection, MealPlanTrans As New Collection
  Set MealTrans = New Collection
  Dim PlanID As Long, MaxIt2 As Long
  
  Dim WriteNewPlan As Boolean
  

  'check if the mealplan is already loaded and then
  'make room for it if it is not
  
  Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealid=-1;", dbOpenDynaset)
  While Not MealplanIn.EOF
       Set MealPlan = DB.OpenRecordset("select * from mealplanner where " _
       & "mealname ='" & Replace(Trim$(MealplanIn("Mealname")), "'", "''") & "';" _
       , dbOpenDynaset)
       
       If Not MealPlan.EOF Then
          'Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealid<>-1;", dbOpenDynaset)
          PlanID = MealPlan("planid")
       Else
          Set MealPlan = DB.OpenRecordset("select max(planid) as maxit from mealplanner;", dbOpenDynaset)
          If IsNull(MealPlan("maxit")) Then
            PlanID = 1
          Else
            PlanID = MealPlan("maxit") + 1
          End If
          If PlanID < 1 Then PlanID = 1
          If MealplanIn("planid") < 0 Then PlanID = MealplanIn("planid")
          
          Set MealPlan = DB.OpenRecordset("select * from mealplanner;", dbOpenDynaset)
          MealPlan.AddNew
          For i = 0 To MealPlan.Fields.Count - 1
            MealPlan.Fields(i) = MealplanIn.Fields(MealPlan.Fields(i).name)
          Next i
          MealPlan.Update
       End If
       MealPlanTrans.Add PlanID, MealplanIn("planid") & ""
       MealplanIn.MoveNext
   Wend
 'now make room for any meals that may be added
   MealPlanTrans.Remove "1"
   MealPlanTrans.Add 1, "1"
   Set MealPlan = DB.OpenRecordset("select max(index) as maxit from mealplanner;", dbOpenDynaset)
   MaxIt2 = MealPlan("maxit") + 1
   Set MealPlan = DB.OpenRecordset("select max(mealid) as maxit from mealplanner;", dbOpenDynaset)
   MaxIt = MealPlan("maxit") + 1
      
  'do final clean up to avoid any preset meals
  If MaxIt = 0 Then MaxIt = 1
  If PlanID < 0 Then PlanID = 1
  
  Set MealplanIn = DBin.OpenRecordset("select * from mealplanner where mealid>-1;", dbOpenDynaset)

  While Not MealplanIn.EOF
     Set MealPlan = DB.OpenRecordset("Select * from mealplanner where planid = " & PlanID _
     & " and mealname ='" & Replace(Trim$(MealplanIn("mealname")), "'", "''") _
     & "';", dbOpenDynaset)
     If MealPlan.EOF Then
         MealPlan.AddNew
         For i = 0 To MealPlan.Fields.Count - 1
            MealPlan(i) = MealplanIn(MealPlan.Fields(i).name)
         Next i
         'on error GoTo errhandl
         MealPlan("index") = MaxIt2
         MaxIt2 = MaxIt2 + 1
         MealPlan("planid") = MealPlanTrans(MealplanIn("planid") & "")
         
         MealPlan("user") = MealplanIn("user")
         
         If MealplanIn("mealid") <> -1 Then
             MealTrans.Add MaxIt, "m" & MealplanIn("mealid")
             MealPlan("Mealid") = MaxIt
             MaxIt = MaxIt + 1
         Else
             MealPlan("planfile") = Replace(Filename, App.path, "~FilePath~", , , vbTextCompare)
         End If
         MealPlan("MealName") = Trim$(MealPlan("Mealname"))
         MealPlan.Update
     Else
         'otherwise just make a reference to it
         'If Not WriteNewPlan Then
             MealTrans.Remove "m" & MealplanIn("mealid")
             MealTrans.Add MealPlan("mealid") & "_s", "m" & MealplanIn("mealid")
     End If
     MealplanIn.MoveNext
  Wend
  Set MealPlan = Nothing: Set MealplanIn = Nothing
      

  'now we need to make the meal definitions
  Dim MealDef As Recordset, MealDefIn As Recordset
  Dim junk As String
  Set MealDef = DB.OpenRecordset("select max(index) as maxit from mealdefinition;", dbOpenDynaset)
  If IsNull(MealDef("maxit")) Then
    MaxIt2 = 1
  Else
    MaxIt2 = Val(MealDef("maxit") & "") + 1
  End If
  Set MealplanIn = DBin.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
  Set MealPlan = DB.OpenRecordset("select * from mealdefinition;", dbOpenDynaset)
  On Error Resume Next
  While Not MealplanIn.EOF
     junk = ""
     junk = MealTrans("m" & MealplanIn("mealid"))
     If Right$(junk, 1) <> "s" And junk <> "" Then
        MealPlan.AddNew
        For i = 0 To MealPlan.Fields.Count - 1
           MealPlan(i) = MealplanIn(MealPlan.Fields(i).name)
        Next i
        MealPlan("index") = MaxIt2
        MealPlan("mealid") = MealTrans("m" & MealplanIn("mealid"))
        MealPlan("abbrevid") = AbbrevTrans("a" & MealplanIn("abbrevid"))
        MaxIt2 = MaxIt2 + 1
        MealPlan.Update
        
     End If
     MealplanIn.MoveNext
  Wend
   'On Error GoTo 0
  Set MealPlan = Nothing: Set MealplanIn = Nothing
  
JustDays:
  
  Dim Meals As Recordset, Mealsin As Recordset
  
  'MsgBox startD
  Dim startD As Date, msD As Date
  'on error Resume Next
  Set Mealsin = DBin.OpenRecordset("Select min(entrydate) as minit from meals;", dbOpenDynaset)
  If IsNull(Mealsin("minit")) Then
    GoTo ExerciseSec
  End If
  msD = Mealsin("minit")
  'on error GoTo errhandl
  Set Mealsin = DBin.OpenRecordset("select * from meals;", dbOpenDynaset)
  Set Meals = DB.OpenRecordset("select max(id) as maxit from meals;", dbOpenDynaset)
  
  MaxIt = Val(Meals("maxit") & "") + 1
  Set Meals = DB.OpenRecordset("select * from meals;", dbOpenDynaset)
  Dim mealIDTrans As New Collection
  'on error Resume Next
  While Not Mealsin.EOF
      j = Abs(DateDiff("d", msD, Mealsin("entrydate")))
      Meals.AddNew
      For i = 0 To Meals.Fields.Count - 1
        Meals.Fields(i) = Mealsin.Fields(Meals.Fields(i).name)
      Next i
      Meals("mealid") = Val(MealTrans("m" & Mealsin("Mealid")))
      Meals("id") = MaxIt
      mealIDTrans.Add MaxIt, Mealsin("id") & ""
      MaxIt = MaxIt + 1
      Meals.Update
      
      Mealsin.MoveNext
  Wend
  
  
  Dim DaysInfo As Recordset
  Dim DYINFOin As Recordset
  Set DYINFOin = DBin.OpenRecordset("select * from daysinfo;", dbOpenDynaset)
  Set DaysInfo = DB.OpenRecordset("select * from daysinfo;", dbOpenDynaset)
  While Not DYINFOin.EOF
     DaysInfo.AddNew
     For i = 0 To DaysInfo.Fields.Count - 1
       DaysInfo.Fields(i) = DYINFOin.Fields(DaysInfo.Fields(i).name)
     Next i
     DaysInfo("itemid") = AbbrevTrans("a" & DaysInfo("itemid"))
     DaysInfo("Mealid") = mealIDTrans(DaysInfo("mealid") & "")
     DaysInfo.Update
     DYINFOin.MoveNext
  Wend
  
  
  Set AbbrevTrans = Nothing
  Set MealTrans = Nothing
  Set recipeTrans = Nothing
  
  
ExerciseSec:


  'load any new exercises
  Dim AbbrevExIn As Recordset, abbrevX As Recordset
  Dim AXtrans As Collection
  Set AXtrans = New Collection
  Set AbbrevExIn = DBin.OpenRecordset("select * from abbrevexercise;", dbOpenDynaset)
  Set abbrevX = DB.OpenRecordset("select max(index) as maxit from abbrevexercise;", dbOpenDynaset)
  If IsNull(abbrevX("maxit")) Then
     MaxIt = 1
  Else
     MaxIt = Val(abbrevX("maxit") & "") + 1
  End If
  While Not AbbrevExIn.EOF
     Set abbrevX = DB.OpenRecordset("select * from abbrevexercise where exercisename ='" & Replace(AbbrevExIn("exercisename"), "'", "''") & "';", dbOpenDynaset)
     If abbrevX.EOF And abbrevX.BOF Then
         abbrevX.AddNew
         For i = 0 To abbrevX.Fields.Count - 1
            abbrevX(i) = AbbrevExIn(abbrevX.Fields(i).name)
         Next i
         abbrevX("index") = MaxIt
         AXtrans.Add MaxIt, "a" & AbbrevExIn("index")
         MaxIt = MaxIt + 1
         abbrevX.Update
     Else
         AXtrans.Add Val(abbrevX("index") & ""), "a" & AbbrevExIn("index")
     End If
     AbbrevExIn.MoveNext
  Wend
  
  Set abbrevX = Nothing: Set AbbrevExIn = Nothing
  Dim ExLog As Recordset, ExlogIn As Recordset
  startD = firstSunday(startD)
  Set ExlogIn = DBin.OpenRecordset("select min(week) as minit from exerciselog;", dbOpenDynaset)
  'on error Resume Next
  Err.Clear
  If IsNull(ExlogIn("minit")) Then
    GoTo ExitSec
  End If
  msD = ExlogIn("minit")
  'on error GoTo errhandl
  Set ExLog = DB.OpenRecordset("select max(index) as maxit from exerciselog;", dbOpenDynaset)
  MaxIt = Val(ExLog("maxit") & "") + 1
  Set ExlogIn = DBin.OpenRecordset("select * from exerciselog;", dbOpenDynaset)
  Set ExLog = DB.OpenRecordset("SELECT * from exerciselog;", dbOpenDynaset)
  While Not ExlogIn.EOF
     ExLog.AddNew
     For i = 0 To ExLog.Fields.Count - 1
        ExLog(i) = ExlogIn(ExLog.Fields(i).name)
     Next i
     ExLog("index") = MaxIt
     j = Abs(DateDiff("d", msD, ExLog("week")))
     ExLog("week") = DateAdd("d", j, startD)
     ExLog("exerciseid") = AXtrans("a" & ExlogIn("exerciseid"))
     ExLog("user") = ExlogIn("user")
     MaxIt = MaxIt + 1
     ExLog.Update
     ExlogIn.MoveNext
  Wend
ExitSec:
  'on error GoTo errhandl
  Set ExLog = Nothing: Set ExlogIn = Nothing
  DBin.Close
  Set DBin = Nothing
  
  
 
  Exit Function
errhandl:
End Function

Public Function firstSunday(InDate As Date) As Date
On Error Resume Next
  Dim i As Long
  i = Weekday(InDate, vbSunday) - 1
  firstSunday = DateAdd("d", -1 * i, InDate)
End Function


Private Function Err_Handler(ByVal ModuleName As String, ByVal ProcName As String, ByVal ErrorDesc As String) As Boolean

    Err_Handler = G_Err_Handler(ModuleName, ProcName, ErrorDesc)


End Function
