Attribute VB_Name = "mdlTools"
Option Explicit
Option Compare Text
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Username
    ' Declare for call to mpr.dll.
   Declare PtrSafe Function WNetGetUser Lib "mpr.dll" _
() '      Alias "WNetGetUserA" (ByVal lpName As String, _
      ByVal lpUserName As String, lpnLength As Long) As Long
     '// 32-bit Function version.
    Const NoError = 0       'The Function call was successful

    '// Note: Get unc-path
    Declare PtrSafe Function WNetGetConnection32 _
    Lib "mpr.dll" _
() '    Alias "WNetGetConnectionA" ( _
    ByVal lpszLocalName As String, _
    ByVal lpszRemoteName As String, _
    lSize As Long) _
    As Long

     '// 32-bit declarations:
    Dim lpszRemoteName As String
    Dim lSize As Long

     '// Use for the return value of WNetGetConnection() API.
    Const NO_ERROR As Long = 0

     '// The size used for the string buffer. Adjust this if you
     '// need a larger buffer.
    Const lBUFFER_SIZE As Long = 1052
    Const cDeactivateString As String = "Deactivated"
Enum fsoSpecFolder
    SpecFolderWindowsInstallation = 1
    SpecFolderSystemFolder = 2
    SpecFolderWindowsTemp = 3
End Enum

Private Declare PtrSafe Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" _
() '(ByVal lpBuffer As String, _
nSize As Long)
'|The rules for NAMING OF VARIABLES:
'|"x_" (x&undescore) meands that variable is produced in  the modell (modell can export it)
'|"i_" (i&undescore) meands that variable is imported
'| After "x_"/"i_" there is a plce for dimensions putted in order of appearance eg. "x_OSAY_"
'| The letters stands for:
'|         O - Origin
'|         A - Age
'|         S - Sex
'|         T - sTatus (Tillstånd in swedish)
'|         Y - Year
'| In cases when number of dimensions coresponds to number of letters the information has the trivial interpration
'| In other cases, when number of dimensions is one fewer the interpretation is that the time dimension (year) is missing in data and
'| the values of the vector reprents valueas for the current simulated year
'| Print order for exports
'| vertical: O,S,A,T
'| horizontal: Y
'| or all dimensions placed verticaly (incl Y - year)
'| Some new ideas (not implemented yet)
'| An alternative output is on textFiles, pivot-formated:
'| dimensions: O,S,A,T,Y,Aggr
'| {Aggr = aggrationtypes as Sum, Count, Min, Max, Mean, Median, Px (eg. P10,P25,P30,P40,P60,P70,P75,P80,P90)}
'|

Function CheckSolver() As Boolean

  Dim bSolverInstalled As Boolean

  '' Assume true unless otherwise
  CheckSolver = True

  On Error Resume Next
  ' check whether Solver is installed

  bSolverInstalled = Application.AddIns("Problemlösaren").Installed
  err.Clear
  If Not bSolverInstalled Then _
    bSolverInstalled = Application.AddIns("Solver Add-In").Installed

  If Not bSolverInstalled Then
    CheckSolver = False
  End If
  On Error GoTo 0
End Function

Public Function fnOpenFromZip() As Boolean
    'Returns true if main is open from zip-archive
    fnOpenFromZip = False 'default
    If ThisWorkbook.ReadOnly Then
        'if zip, allways readonly
        If InStr(1, LCase(ThisWorkbook.FullName), ".zip", vbTextCompare) <> 0 Then
            fnOpenFromZip = True
            MsgBox fngetMsgBoxtext(14), vbCritical
        End If
    End If
End Function



Public Sub formatBoldTitles(ByRef r As Range, S As Integer)
    With r
        If Not .Font.Bold Then .Font.Bold = True
        If .Font.Name <> "Arial" Then .Font.Name = "Arial"
        If .Font.Size <> S Then .Font.Size = S
        If .Interior.Color <> vbYellow Then .Interior.Color = vbYellow
    End With
End Sub



Function FileExists(FileName As String) As Boolean
' returns TRUE if the file exists
    FileExists = Len(Dir(ThisWorkbook.Path & "\" & FileName)) > 0
End Function

Function FileExists_Path(strPath As String) As Boolean
' returns TRUE if the file exists
    FileExists_Path = Len(Dir(strPath)) > 0
End Function


Public Function getOpenedWorkbook(wb_name As String) As Workbook
 Dim wb As Workbook, wb_i As Workbook
 Dim found As Boolean
'| Function returning the open Workbook
'> wb_name as string - name of the workbook to be found including the file extension ("xls", "xlsm")
'< reference to the the found workbook or Nothing if not found (test proposal "result is Nothing")

    For Each wb_i In Workbooks
      If LCase(wb_i.Name) = LCase(wb_name) Then
         Set wb = wb_i
         found = True
         Exit For
      End If
    Next wb_i

    If found Then
        Set getOpenedWorkbook = wb
    Else
        Set getOpenedWorkbook = Nothing
    End If
End Function

Public Sub doBookVisable(wb As Workbook)
Dim w As window
    For Each w In wb.Windows
        If Not w.Visible Then w.Visible = True
    Next w
End Sub


 Public Function GetBook(theName As String) As Workbook
    Dim wb As Workbook
    Dim k As Integer
    Dim fName As String
    Dim Ws As Worksheet
'| LoadFile checks if the excelbook is already open in the application then return True
'| if the book is not open yet the routine locates the book on the disk (in the model folder)
'| and tries to open it. Once opened all windows are turned to visable status
'> theName as string - name of the workbook with the macro including the file extension ("xlsx", "xlsm")
'< reference to the workbook if already succes, otherwise nothing

    Set wb = getOpenedWorkbook(theName)

    If Not (wb Is Nothing) Then
        Set GetBook = wb 'Case: already there
        Exit Function
    End If

    fName = ThisWorkbook.Path & "\" & theName
    If Not TheFileExists(theName) Then
       k = MsgBox("'" & theName & "'" & " not found in the model directory. Do you want to create it?", Buttons:=vbYesNo)
       If k = vbYes Then
        ChDir ThisWorkbook.Path
        Workbooks.Add
        Set wb = Workbooks(Workbooks.Count)
        Application.DisplayAlerts = False
        wb.Worksheets(2).Delete
        wb.Worksheets(2).Delete
        wb.Worksheets(1).Name = "Empty"
        Application.DisplayAlerts = False

        wb.SaveAs FileName:=fName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Set GetBook = wb
       Else ' k = vbNO
        MsgBox fngetMsgBoxtext(76) & theName & " " & fngetMsgBoxtext(78)
        Set GetBook = Nothing
        Exit Function
       End If
       Exit Function
    End If

    Workbooks.Open FileName:=fName, UpdateLinks:=False 'Case: LOAD
    Set wb = Workbooks(theName)
    wb.Windows(1).WindowState = xlMinimized
    Set GetBook = wb
 End Function

 Public Function getSheet(bookName As String, Sheetname As String) As Worksheet
 Dim S As String
 Dim Ws As Worksheet
 Dim wb As Workbook

    Set wb = GetBook(bookName)

    If wb Is Nothing Then
        Set getSheet = Nothing
        Exit Function
    End If

    S = LCase(Sheetname)
    For Each Ws In wb.Sheets
      If LCase(Ws.Name) = S Then
         Set getSheet = Ws
         Exit Function
      End If
    Next Ws

    Set Ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

    Ws.Name = Sheetname

    Set getSheet = Ws
 End Function

Function fngetWindowsState(ByVal strWorkbookName As String) As Long
    'Some times windows name is different from workbook name, e.g. if user have more views open for same book.
    'First try name
    Dim lngWinState As Long
    Dim lngLenName As Long
    Dim awinDow 'As Application.Windows
    On Error Resume Next
    lngWinState = Application.Windows(strWorkbookName).WindowState
    If err.Number = 0 Then
        On Error GoTo 0
        fngetWindowsState = lngWinState
        Exit Function
    End If
    err.Clear
    'Search by looping

    For Each awinDow In Application.Windows
        If UCase(Left(awinDow.Caption, Len(strWorkbookName))) = UCase(strWorkbookName) Then
            On Error GoTo 0
            fngetWindowsState = awinDow.WindowState
            Exit Function
        End If
    Next

    'here nothing was found. Should never happen.
    MsgBox "Unable to locate workbook " & strWorkbookName, vbCritical

End Function
Sub exempelHowToUse()
    Debug.Print fnReplaceUnvalidChars("_KalleåäöKalle/&%¤#¨!=99/*-€'*~¨¨¨")
End Sub
Function fnReplaceUnvalidChars(ByVal strToBeReplaced) As String
    Dim i As Long
    For i = 1 To Len(strToBeReplaced)
        If Not fnValidCharForNamedRange(Mid(strToBeReplaced, i, 1)) Then
            'Replace with "_"
            strToBeReplaced = Left(strToBeReplaced, i - 1) & "_" & Mid(strToBeReplaced, i + 1)
        End If
    Next
    fnReplaceUnvalidChars = strToBeReplaced
End Function

Function fnValidCharForNamedRange(ByVal strChar As String) As Boolean
    'Valid Chr between 48 and 122 and not in range 58-64 or 91-96
    Dim lngAscii As Long
    lngAscii = Asc(strChar)
    fnValidCharForNamedRange = True 'Dft
    If lngAscii < 48 Or lngAscii > 122 Then
        fnValidCharForNamedRange = False
        Exit Function
    End If
    If (lngAscii >= 58 And lngAscii <= 64) Or (lngAscii >= 91 And lngAscii <= 96) Then
        fnValidCharForNamedRange = False
        Exit Function
    End If
End Function
Public Function getLastRowInColumn(wb As Workbook, wkName As String, intColumn As Integer) As Long
    getLastRowInColumn = wb.Worksheets(wkName).Cells(1000000, intColumn).End(xlUp).Row
End Function
Sub CreateNamedRange(strName As String, rngReference As Range)
    ActiveWorkbook.Names.Add Name:=strName, RefersTo:=rngReference
End Sub
Sub CreateNamedRangeInWB(strName As String, rngReference As Range, strWB As String)
    Workbooks(strWB).Names.Add Name:=strName, RefersTo:=rngReference
End Sub
Sub DeleteNames(wbName As String)
    Dim nName As Name
    For Each nName In Workbooks(wbName).Names
        If Left(nName.Name, 5) = "_scn_" Then
            nName.Delete
        End If
    Next nName
End Sub
Sub DeleteSpecName(wbName As String, rngName As String)
    Dim nName As Name
    For Each nName In Workbooks(wbName).Names
        If nName.Name = "_scn_" & rngName Then
            nName.Delete
        End If
    Next nName
End Sub
Sub DeleteNameRange(wbName As String, rngName As String)
    Dim nName As Name
    For Each nName In Workbooks(wbName).Names
        'If LCase(Left(nName.Name, Len(rngName))) = LCase(rngName) Then
        If InStr(nName.Name, rngName) > 0 Then
            nName.Delete
        End If
    Next nName
End Sub
Sub CleanNameRangeScenario(wb As Workbook)
    Dim strN As String
    Dim N As Name
    For Each N In wb.Names
        If InStr(N.Name, "!") > 0 Then
            strN = Right(N.Name, Len(N.Name) - WorksheetFunction.Search("!", N.Name))
            If LCase(Left(strN, 4)) = "top_" Or LCase(Left(strN, 5)) = "last_" Then
                N.Delete
            End If
        End If
    Next N
End Sub
Sub CleanOneNameRangeScenario(wb As Workbook, strNamerange As String)
    Dim strN As String
    Dim N As Name
    On Error Resume Next
    For Each N In wb.Names
        strN = Right(N.Name, Len(N.Name) - WorksheetFunction.Search("!", N.Name))
        If LCase(strN) = LCase(strNamerange) Then
            N.Delete
        End If
    Next N
End Sub
Sub AllNameRange()
    Dim nName As Name
    For Each nName In ThisWorkbook.Names
        Debug.Print nName.Name '& "--" & nName.RefersToRange
    Next nName
End Sub

Public Function RangeNameExists(argRangeName As String) As Boolean
    Dim N As Name
    RangeNameExists = False
    For Each N In ActiveWorkbook.Names
        If UCase(N.Name) = UCase(argRangeName) Then
            RangeNameExists = True
            Exit Function
        End If
    Next N
End Function

Public Function RangeNameExistsWB(argRangeName As String, wb As Workbook, wk As Worksheet) As Boolean
    Dim N As Name
    Dim strN As String
    RangeNameExistsWB = False
    For Each N In wb.Sheets(wk.Name).Names
        strN = Right(N.Name, Len(N.Name) - WorksheetFunction.Search("!", N.Name))
        If UCase(strN) = UCase(argRangeName) Then
            RangeNameExistsWB = True
            Exit Function
        End If
    Next N
End Function
Public Function RangeNameExistsInWorkbook(argRangeName As String, wb As Workbook) As Boolean
    Dim N As Name
    RangeNameExistsInWorkbook = False
    For Each N In wb.Names
        If UCase(N.Name) = UCase(argRangeName) Then
            RangeNameExistsInWorkbook = True
            Exit Function
        End If
    Next N
End Function

Public Function RangeNameDead(argRangeName As String, wb As Workbook, wk As Worksheet) As Boolean
    Dim N As Name
    RangeNameDead = False
    On Error Resume Next
    For Each N In wb.Sheets(wk.Name).Names
        If InStr(1, N.RefersTo, "#REF!") > 0 Then
            RangeNameDead = True
            On Error GoTo 0
            Exit Function
        End If
    Next N
    On Error GoTo 0
End Function
Public Function RangeNameDeadInWb(argRangeName As String, wb As Workbook) As Boolean
    Dim N As Name
    RangeNameDeadInWb = False
    For Each N In wb.Names
        If InStr(1, N.RefersTo, "#REF!") > 0 Then
            RangeNameDeadInWb = True
            Exit Function
        End If
    Next N
End Function

Public Function RangeNameExistsInWorkbookString(argRangeName As String, wbName As String) As Boolean
    Dim N As Name
    RangeNameExistsInWorkbookString = False
    For Each N In Workbooks(wbName).Names
        If UCase(N.Name) = UCase(argRangeName) Then
            RangeNameExistsInWorkbookString = True
            Exit Function
        End If
    Next N
End Function
Sub CheckNameRangeValidName(ByVal strNamerange As String, ByVal wbName As String, ByVal wkName As String)

    Dim r As Range
    Dim c As Range
    Dim strRngAddress As String
    Dim strInput As String
    Dim rng As Range
    Dim wb As Workbook
    'On Error Resume Next
    If Not RangeNameExistsInWorkbookString(strNamerange, wbName) Then
        Workbooks(wbName).Activate
        Worksheets(wkName).Activate
        Set rng = Application.InputBox(prompt:=fngetMsgBoxtext(149) & "...", _
        Title:=fngetMsgBoxtext(149) & ":", Type:=8)
        If rng Is Nothing Then
            Exit Sub
        Else
            If rng.Columns.Count > 2 Then
                MsgBox "Please check your selected cells.", vbOKOnly
            Else
                If rng.Columns.Count = 1 Then
                     strRngAddress = rng.Address
                     Workbooks(rng.Parent.Parent.Name).Names.Add Name:=strNamerange, RefersTo:="=[" & rng.Parent.Parent.Name & "]" & rng.Parent.Name & "!" & rng.Address
                     ThisWorkbook.Activate
                Else
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Sub
Sub msgError(ByVal msg As String)
    Dim Style, Title, Response
    Style = vbOKOnly + vbCritical + vbDefaultButton2
    Title = "Uponor - Control"
    Response = MsgBox(msg, Style, Title)
End Sub
Function GetFilenameFromPath(ByVal strPath As String) As String

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Public Function TheFileExists(ByVal fName As String) As Boolean
    Dim x As String
    x = Dir(ThisWorkbook.Path & "\" & fName)
    If x <> "" Then TheFileExists = True _
        Else TheFileExists = False
End Function

Public Function FileNameOnly(pname) As String
    Dim Length As Long
    Dim temp As Variant
    Length = Len(pname)
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function

Public Function PathExists(pname) As Boolean

  If Dir(pname, vbDirectory) = "" Then
    PathExists = False
 Else
    PathExists = True
    PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
 End If
End Function
Function CreateFolder(strFolder)
    If Dir(strFolder, vbDirectory) = vbNullString Then
        Debug.Print strFolder
        MkDir (strFolder)
    Else
    End If

End Function
Function SheetExists(Sheetname As String) As Boolean
    SheetExists = False
    On Error GoTo NoSuchSheet
    If Len(Sheets(Sheetname).Name) > 0 Then
        SheetExists = True
        Exit Function
    End If
NoSuchSheet:
End Function

Function SheetExistsWb(Sheetname As String, wbName As String) As Boolean
    SheetExistsWb = False
    On Error GoTo NoSuchSheet
    If Len(Workbooks(wbName).Worksheets(Sheetname).Name) > 0 Then
        SheetExistsWb = True
        Exit Function
    End If
NoSuchSheet:
End Function
Function WorkbookOpen(WorkBookName As String) As Boolean
    WorkbookOpen = False
    On Error GoTo WorkBookNotOpen
    If Len(Application.Workbooks(WorkBookName).Name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkBookNotOpen:
End Function
Public Function RealUsedRange(ByRef Ws As Worksheet) As Range

    Dim FirstRow        As Long
    Dim LastRow         As Long
    Dim FirstColumn     As Integer
    Dim LastColumn      As Integer


   On Error Resume Next
    FirstRow = Ws.Cells.Find(What:="*", After:=Range("IV65536"), LookIn:=xlValues, Lookat:= _
    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    If FirstRow = 0 Then FirstRow = 3

    FirstColumn = Ws.Cells.Find(What:="*", After:=Range("IV65536"), LookIn:=xlValues, Lookat:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
    If FirstColumn = 0 Then FirstColumn = 3

    LastRow = Ws.Cells.Find(What:="*", After:=Range("D3"), LookIn:=xlValues, Lookat:= _
    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If LastRow = 0 Then LastRow = 3

    LastColumn = Ws.Cells.Find(What:="*", After:=Range("C1"), LookIn:=xlValues, Lookat:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If LastColumn = 0 Then LastColumn = 3

    Set RealUsedRange = Range(Ws.Cells(FirstRow, FirstColumn), Ws.Cells(LastRow, LastColumn))

    On Error GoTo 0

End Function


Function GetTheNameNETWORK() As String
    Dim objNet As Object
    On Error Resume Next
    Set objNet = CreateObject("WScript.NetWork")
    'MsgBox "Network username is: " & objNet.UserName
    GetTheNameNETWORK = objNet.UserName
    Set objNet = Nothing
End Function
Function fnGetFolderName(ByVal srtPathname As String) As String
    Dim lFind As Long
    Dim strFolder As String
    lFind = InStrRev(srtPathname, "\")
    strFolder = Mid(srtPathname, lFind + 1, Len(srtPathname) - lFind)
    fnGetFolderName = strFolder
End Function


Sub DebugPrintNamerange()
    Dim N As Name

    For Each N In ThisWorkbook.Names
        Debug.Print N.Name & "," & N.RefersToLocal
    Next N

End Sub

Sub DebugPrintWorksheets()
    Dim N As Worksheet

    For Each N In ThisWorkbook.Worksheets
        Debug.Print N.Name
    Next N

End Sub
Sub SortListBox(oLb As MSForms.ListBox)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim vTemp As Variant

    vaItems = oLb.List
    For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
        For j = i + 1 To UBound(vaItems, 1)
            If vaItems(i, 0) > vaItems(j, 0) Then
                vTemp = vaItems(i, 0)
                vaItems(i, 0) = vaItems(j, 0)
                vaItems(j, 0) = vTemp
            End If
        Next j
    Next i
    oLb.Clear
    For i = LBound(vaItems, 1) To UBound(vaItems, 1)
        If Len(vaItems(i, 0)) > 1 Then oLb.AddItem vaItems(i, 0)
    Next i
End Sub


Public Function LastRowInColumn(ByVal wb As Workbook, ByVal Ws As String, ByVal strColumn As String) As Long
    Dim wk As Worksheet
    Set wk = wb.Worksheets(Ws)
    With wk
        LastRowInColumn = .Cells(.Rows.Count, strColumn).End(xlUp).Row
    End With
    Set wk = Nothing
End Function

Function fnSheetCodeName(wb As Workbook, CodeName As String) As String
    fnSheetCodeName = wb.VBProject.VBComponents(CodeName).Properties("Name").Value
 End Function


Public Sub subAbortSystemMainFilesIsMissing()
    'This is run if system critical system binary files does not exist.
    MsgBox "System files for Pensionsmodellen is missing. The run will end." & vbCrLf & vbCrLf & "Recommended action: Close excel. Try to copy all the files (including folder Data) from your source to a new folder. ", vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    End
End Sub
Function fnNamedRangeOutputExists(strName As String, ByVal wkbName As String) As Boolean
    Dim rngRangeNameToFind As Range
    Dim i As Long
    On Error Resume Next
    If wkbName = vbNullString Then wkbName = ActiveWorkbook.Name

    With Workbooks(wkbName)
        For i = 1 To .Sheets.Count Step 1
            Set rngRangeNameToFind = .Sheets(i).Range(strName)
            Select Case err.Number
               Case 0
                fnNamedRangeOutputExists = True
                Exit Function
               Case 1004
                err.Clear
               Case Else
                MsgBox err.Number & " " & err.Description
            End Select
        Next
    End With
    On Error GoTo 0
End Function

Sub CheckComputerProperties()

    Debug.Print Environ("NUMBER_OF_PROCESSORS")
End Sub
Sub List_Environment_Variables()
Dim EnvString As String
Dim Indx As Long
Indx = 1
Do
    EnvString = Environ(Indx)
    Debug.Print EnvString
    Indx = Indx + 1
Loop Until EnvString = ""
End Sub


    Function fnUNCPath(strDriveLetter As String) As String
         '// Takes specified Local Drive Letter
         '// eg E,D,H Etc and converts to UNC

        Dim cbRemoteName As Long
        Dim lStatus As Long

         '// Add a colon to the drive letter entered.
        strDriveLetter = Left(strDriveLetter, 1) & ":"

         '// Specifies the size in charaters of the buffer.
        cbRemoteName = lBUFFER_SIZE

         '// Prepare a string variable by padding spaces.
        lpszRemoteName = lpszRemoteName & Space(lBUFFER_SIZE)

         '// Return the UNC path (eg.\\Server\Share).
        lStatus = WNetGetConnection32( _
        strDriveLetter, _
        lpszRemoteName, _
        cbRemoteName)

         '// Has WNetGetConnection() succeeded.
         '// WNetGetConnection()returns 0 (NO_ERROR)
         '// if it succesfully retrieves the UNC path.
        If lStatus = NO_ERROR Then
             '// Get UNC path.
            fnUNCPath = Trim(lpszRemoteName)
            If Asc(Right(fnUNCPath, 1)) = 0 Then fnUNCPath = Mid(fnUNCPath, 1, Len(fnUNCPath) - 1) 'remove trailing char ascii=0
        Else
             '// Unable to obtain the UNC path.
            fnUNCPath = "0"
        End If

    End Function

    Function fnReplacePathWithUncPath(ByVal strPath As String) As String
        '  Returns UNC-path: \\ppm.nu\dfs\Gemensam
        '1.Get drive letter from path
        '2.Get mapped unc and replace
        Dim strDriveLetter As String
        Dim strUncPath As String
        Dim strNewPath As String
        strPath = Trim(strPath)
        strDriveLetter = Left(strPath, 1)
        strNewPath = Mid(strPath, 3) 'remove driveletter and colon (C:)

        strUncPath = fnUNCPath(strDriveLetter)
        If strUncPath = "0" Then
            fnReplacePathWithUncPath = strPath
        Else
            fnReplacePathWithUncPath = strUncPath & strNewPath
        End If

    End Function

   Public Function fnGetUserName() As String

      ' Buffer size for the return string.
      Const lpnLength As Integer = 255

      ' Get return buffer space.
      Dim Status As Integer

      ' For getting user information.
      Dim lpName, lpUserName As String

      ' Assign the buffer size constant to lpUserName.
      lpUserName = Space$(lpnLength + 1)

      ' Get the log-on name of the person using product.
      Status = WNetGetUser(lpName, lpUserName, lpnLength)

      ' See whether error occurred.
      If Status = NoError Then
         ' This line removes the null character. Strings in C are null-
         ' terminated. Strings in Visual Basic are not null-terminated.
         ' The null character must be removed from the C strings to be used
         ' cleanly in Visual Basic.
         lpUserName = Left$(lpUserName, InStr(lpUserName, chr(0)) - 1)
      Else

         ' An error occurred.
         fnGetUserName = ""
         End
      End If

      ' Display the name of the person logged on to the machine.
      fnGetUserName = lpUserName

   End Function


Function fnFileIsOpen(ByVal strFilename As String) As Boolean
    Dim iFilenum As Long
    Dim iErr As Long

    On Error Resume Next
    iFilenum = FreeFile()
    Open strFilename For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = err
    On Error GoTo 0

    Select Case iErr
        Case 0:    fnFileIsOpen = False
        Case 70:   fnFileIsOpen = True
        Case Else
            fnFileIsOpen = False
            'MsgBox fngetMsgBoxtext(66), vbInformation
    End Select

End Function
Function fnPathAddBackslashIfNeeded(ByVal strPath As String) As String
    'Adding "\" if needed, replace "/" with "\"
     If strPath = "" Then
        'If empty string, do nothing
        fnPathAddBackslashIfNeeded = ""
        Exit Function
    End If
    strPath = Replace(strPath, "/", "\", , , vbTextCompare) ' Slash is valid, but replace / with \
    If Right(Trim(strPath), 1) <> "\" Then strPath = Trim(strPath) & "\"
    fnPathAddBackslashIfNeeded = strPath
End Function

Function fnIsOpenwb(ByVal strwbName As String, Optional ByVal blnReadOnly As Boolean) As Boolean
Dim strFullpath As String
Dim wb As Workbook
On Error Resume Next
Set wb = Workbooks(strwbName)
'strfullPath = ThisWorkbook.Path
'strfullPath = strfullPath & "\" & strWbName
strFullpath = strwbName
If wb Is Nothing Then
    fnIsOpenwb = False
    Application.ScreenUpdating = False
    Application.Workbooks.Open FileName:=strFullpath, AddToMru:=False, ReadOnly:=blnReadOnly, UpdateLinks:=True
    Application.ScreenUpdating = True
Else
    fnIsOpenwb = True
End If
Set wb = Nothing
End Function

'Function fnPathAddBackslashIfNeeded(ByVal strPath As String) As String
'
'
'
'    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
'    fnPathAddBackslashIfNeeded = strPath
'End Function
Function fnOpenWb(ByRef wb As Workbook, ByVal strwbName As String, Optional ByVal blnReadOnly As Boolean, Optional blnUpdateLinks As Boolean = True, Optional blnVisible As Boolean = True, Optional appXL As Excel.Application) As Boolean
    'Sets wb as workbook
    'returns true of was open
    Dim strFullpath As String
    Dim StatusScreenUpdating As Long

    If appXL Is Nothing Then Set appXL = Application 'Default

    strFullpath = ThisWorkbook.Path
    strFullpath = fnPathAddBackslashIfNeeded(strFullpath)
    strFullpath = strFullpath & strwbName

    If Not FileExists_Path(strFullpath) Then
        fnOpenWb = False
        Exit Function
    End If

    'first test if wb open
    On Error Resume Next
    Set wb = appXL.Workbooks(strwbName)
    If err.Number = 0 Then
        fnOpenWb = True
        On Error GoTo 0
        Exit Function
    Else
        fnOpenWb = False
    End If
    On Error GoTo 0

    'if open, then open
    If fnOpenWb Then
        If Not blnReadOnly Then
            'Check if open
            If fnFileIsOpen(strFullpath) Then
                Set wb = Nothing
                Exit Function
            End If
        End If
    End If
    If Not blnVisible Then
        'Only when hidden
        StatusScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
    End If
    'Dim lngAppWinState As Long

    'lngAppWinState = appXL.WindowState 'Get windowsstate for application
    Set wb = appXL.Workbooks.Open(FileName:=strFullpath, AddToMru:=False, ReadOnly:=blnReadOnly, UpdateLinks:=blnUpdateLinks)
    'appXL.Windows(fngetWb_WindowName(wb.Name)).WindowState = lngAppWinState 'Set windowsstate same as for application
    appXL.Windows(fngetWb_WindowName(wb.Name)).WindowState = xlMaximized
    appXL.Windows(fngetWb_WindowName(wb.Name)).Visible = blnVisible

    If Not blnVisible Then Application.ScreenUpdating = StatusScreenUpdating

End Function


Function fnOpenWb_withPath(ByRef wb As Workbook, ByVal strwbName As String, ByVal strPath As String, Optional ByVal blnReadOnly As Boolean, Optional blnUpdateLinks As Boolean = True, Optional blnVisible As Boolean = True) As Boolean
    'Sets wb as workbook
    'returns true of was open
    Dim strFullpath As String
    Dim StatusScreenUpdating As Long

    strFullpath = strPath

    'first test if wb open
    On Error Resume Next
    Set wb = Workbooks(strwbName)
    If err.Number = 0 Then
        fnOpenWb_withPath = True
        On Error GoTo 0
        Exit Function
    Else
        fnOpenWb_withPath = False
    End If
    On Error GoTo 0

    'application.Windows("dddd").Visible=
    'if open, then open
    If fnOpenWb_withPath Then
        If Not blnReadOnly Then
            'Check if open
            If fnFileIsOpen(strFullpath) Then
                Set wb = Nothing
                Exit Function
            End If
        End If
    End If
    If Not blnVisible Then 'Application.ScreenUpdating = StatusScreenUpdating
        'Only when hidden
        StatusScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
    End If
    Set wb = Application.Workbooks.Open(FileName:=strFullpath, AddToMru:=False, ReadOnly:=blnReadOnly, UpdateLinks:=blnUpdateLinks)
    'subShowHideWorkbook strWbName, False 'problem retrieve info from hidden wb?!
    'Set wb = ActiveWorkbook
    Application.Windows(fngetWb_WindowName(wb.Name)).Visible = blnVisible
    If Not blnVisible Then Application.ScreenUpdating = StatusScreenUpdating

End Function

Function fngetWb_WindowName(ByVal strWorkbookName As String) As String
    'Some times windows name is different from workbook name, e.g. if user have more views open for same book.
    'First try name
    Dim lngLenName As Long
    Dim awinDow 'As Application.Windows
    Dim dummy As Long
    On Error Resume Next
    dummy = Application.Windows(strWorkbookName).Visible
    If err.Number = 0 Then
        On Error GoTo 0
        fngetWb_WindowName = strWorkbookName
        Exit Function
    End If
    err.Clear

    'Search by looping
    For Each awinDow In Application.Windows
        If UCase(Left(awinDow.Caption, Len(strWorkbookName))) = UCase(strWorkbookName) Then
            On Error GoTo 0
            fngetWb_WindowName = awinDow.Caption
            Exit Function
        End If
    Next

    'here nothing was found. Should never happen.
    MsgBox fngetMsgBoxtext(67) & " " & strWorkbookName, vbCritical
    Set awinDow = Nothing
End Function

Function fnDisconnectWb(ByRef wb As Workbook, Optional ByVal blnSave As Boolean) As Boolean
    Dim i As Long
    Application.ScreenUpdating = False
    If blnSave Then
        'subShowHideWorkbook wb.Name, True 'make visible
        wb.Save
        fnWait (50)
    End If
    wb.Close False
    fnWait (50)
    Set wb = Nothing
    DoEvents
    Application.ScreenUpdating = True
End Function
Function fnWait(ByVal lngWaitNbrOfLoops As Long)
    Dim i As Long
    For i = 1 To lngWaitNbrOfLoops
        DoEvents
    Next
End Function

Public Function fnGetParentFolderName() As String
    Dim ParentFolderName As String
    ParentFolderName = CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder.Path
    fnGetParentFolderName = ParentFolderName
End Function

Function fnGetRangeFromNames(ByRef wb As Workbook, ByVal strNameOfRange As String, Optional blnSilent As Boolean = True) As Range
    'Returns a Range object
    'Locating a Range (using Names collection)
    'Only way to grab with wb and range-name
    Dim strRefTo As String
    Dim strAdressA1 As String
    Dim strWs As String
'''    Dim i As Long

    On Error Resume Next
    strRefTo = wb.Names(strNameOfRange) 'Returning e.g.: =main!$F$7
    If err.Number > 0 Then
        Set fnGetRangeFromNames = Nothing 'error
        If Not blnSilent Then MsgBox "Error finding range " & strNameOfRange & " in workbook " & wb.Name & ".", vbCritical
    Else
        'Parse worksheet
        Set fnGetRangeFromNames = wb.Names(strNameOfRange).RefersToRange
'''        i = InStr(1, strRefTo, "!")
'''        strWs = Mid(strRefTo, 2, i - 2)
'''        strWs = Replace(strWs, "'", "")
'''        strAdressA1 = Mid(strRefTo, i + 1)
'''        Set fnGetRangeFromNames = wb.Worksheets(strWs).Range(strAdressA1)
    End If
    On Error GoTo 0

End Function
Function fnGetRangeFromNamesTw(ByVal strNameOfRange As String, Optional blnSilent As Boolean = True) As Range
    Set fnGetRangeFromNamesTw = fnGetRangeFromNames(ThisWorkbook, strNameOfRange, blnSilent)
End Function

Public Function fnRangeNameExists(argRangeName As String) As Boolean
    Dim N As Name
    fnRangeNameExists = False
    For Each N In ThisWorkbook.Names
        If UCase(N.Name) = UCase(argRangeName) Then
            fnRangeNameExists = True
            Exit Function
        End If
    Next N
End Function

'Sub DeleteNameRange(wbName As String, rngName As String)
'    Dim nName As Name
'    For Each nName In Workbooks(wbName).Names
'        If InStr(nName.Name, rngName) > 0 Then
'            nName.Delete
'        End If
'    Next nName
'End Sub

Public Function fnCheckAbout() As Boolean
    Dim rng As Range
    Dim r As Range

    fnCheckAbout = False
    Set rng = Application.Range("rng_TopLeftAbout").Resize(Application.Range("rng_TopLeftAbout").CurrentRegion.Rows.Count, 1)
    If Not rng Is Nothing Then
        For Each r In rng
            If LCase(r.Value) = LCase(GetTheNameNETWORK) And LCase(r.Offset(, 1).Value) = LCase(ThisWorkbook.Path) Then
                fnCheckAbout = True
                Set rng = Nothing
                Exit Function
            End If
        Next r
    End If
    fnCheckAbout = False
    Set rng = Nothing
End Function

Sub cmdActivateSheet()
    subActivateOrDeactivateSheets True
End Sub
Sub cmdDeActivateSheet()
    subActivateOrDeactivateSheets False
End Sub
Sub subActivateOrDeactivateSheets(ByVal blnActivate As Boolean)
    Dim rng As Range
    Dim cll As Range
    Set rng = ThisWorkbook.Names("rngDeActivatedSheets").RefersToRange.CurrentRegion.Offset(1, 0)
    Set rng = rng.Resize(rng.Rows.Count - 1, 1)

    Application.EnableEvents = False
    For Each cll In rng.Cells

        'Check for sheet existence
        On Error Resume Next
        If ThisWorkbook.Worksheets(cll.Value).Name <> "dummy" Then
        End If
        If err.Number > 0 Then
            On Error GoTo 0
            MsgBox "Sheet " & cll.Value & " is missing.", vbInformation
        Else
            On Error GoTo 0
            If blnActivate Then
                subActivateSheetFormulas cll.Value, CStr(cll.Offset(0, 1).Value)
            Else
                subDeActivateSheetFormulas cll.Value, CStr(cll.Offset(0, 1).Value)
            End If
        End If
    Next
    Application.EnableEvents = True
    Application.CalculateFullRebuild
End Sub


Public Sub subActivateSheetFormulas(ByVal strSheetName As String, Optional strRangename As String = "")

    'Replace all formulas in a sheet with 'Deactivated

    'To speed up - but not need to delete sheets that might be used later in development process.

    Dim cll As Range
    Dim rng As Range
    Dim dblRowHeight As Double

    On Error Resume Next
    If strRangename = "" Then
        'entire worksheet
        Set rng = ThisWorkbook.Worksheets(strSheetName).Cells.SpecialCells(xlCellTypeConstants)
    Else
        'Only part of worksheet
        On Error Resume Next
        Set rng = ThisWorkbook.Worksheets(strSheetName).Range(strRangename)
        On Error GoTo 0
        If rng Is Nothing Then
            MsgBox "Du har angivit fel range för flik " & strSheetName & ": " & strRangename, vbInformation
        Else
            Set rng = rng.SpecialCells(xlCellTypeConstants)
        End If

    End If

    On Error GoTo 0
    If Not rng Is Nothing Then
        For Each cll In rng
            If Left(cll.Value, Len(cDeactivateString)) = cDeactivateString Then
                dblRowHeight = cll.RowHeight
                'Only modify cells with deactivate code
                If Mid(cll.Value, Len(cDeactivateString) + 1, 1) = "A" Then
                    'Array formula
                    cll.FormulaArray = Replace(cll.Value, cDeactivateString & "A", "", , , vbTextCompare)
                Else
                    'Normal formula
                    cll.Formula = Replace(cll.Value, cDeactivateString, "", , , vbTextCompare)
                End If
                If dblRowHeight <> cll.RowHeight Then cll.RowHeight = dblRowHeight
            End If
        Next
    End If
End Sub

Public Sub subDeActivateSheetFormulas(ByVal strSheetName As String, Optional strRangename As String = "")

    'Replace all formulas in a sheet with 'Deactivated

    'To speed up - but not need to delete sheets that might be used later in development process.

    Dim cll As Range
    Dim rng As Range
    Dim dblRowHeight As Double

    On Error Resume Next
    If strRangename = "" Then
        'entire worksheet
        Set rng = ThisWorkbook.Worksheets(strSheetName).Cells.SpecialCells(xlCellTypeFormulas)
    Else
        'Only part of worksheet
        On Error Resume Next
        Set rng = ThisWorkbook.Worksheets(strSheetName).Range(strRangename)
        On Error GoTo 0
        If rng Is Nothing Then
            MsgBox "Du har angivit fel range för flik " & strSheetName & ": " & strRangename, vbInformation
        Else
            Set rng = rng.SpecialCells(xlCellTypeFormulas)
        End If

    End If

    On Error GoTo 0

    If Not rng Is Nothing Then

        For Each cll In rng
            dblRowHeight = cll.RowHeight
            If cll.HasArray Then
                cll.Value = cDeactivateString & "A" & cll.Formula
            Else
                cll.Value = cDeactivateString & cll.Formula
            End If
            If dblRowHeight <> cll.RowHeight Then cll.RowHeight = dblRowHeight
       Next

    End If

End Sub
Sub subDeactivateCalcStatusForSheets()
    'Deactivate calc status if needed.
    'Store calc status in collection pcCalcStatusForSheets
    Dim Ws As Worksheet

    For Each Ws In ThisWorkbook.Worksheets
        Ws.EnableCalculation = False
    Next

End Sub
Sub subActivateCalcStatusForSheets()
    'Activate calc status if needed.
    'ReStore calc status from collection pcCalcStatusForSheets
    Dim Ws As Worksheet

    For Each Ws In ThisWorkbook.Worksheets
        Ws.EnableCalculation = True
    Next

End Sub

Sub loadsheets()

    Dim rng As Range
    Set rng = ActiveCell
    Dim Ws As Worksheet
    For Each Ws In ThisWorkbook.Worksheets
        rng.Value = Ws.Name
        Set rng = rng.Offset(1, 0)
    Next
End Sub


Sub ShowTaxRows()
    Call Unprotect_sheet(wsStart.CodeName)
    wsStart.Rows(wsStart.Range("rng_Skatt_Row_Top").Row & ":" & wsStart.Range("rng_Skatt_Row_Top").Row + 4).EntireRow.Hidden = False
    Call Protect_sheet(wsStart.CodeName)
End Sub
Sub HideTaxRows()
    Call Unprotect_sheet(wsStart.CodeName)
    wsStart.Rows(wsStart.Range("rng_Skatt_Row_Top").Row & ":" & wsStart.Range("rng_Skatt_Row_Top").Row + 4).EntireRow.Hidden = True
    Call Protect_sheet(wsStart.CodeName)
End Sub
Sub ShowMonthlyChart()
    With wsStart
        .Unprotect
        .Rows(wsStart.Range("rngChartSection_month").Row & ":" & wsStart.Range("rngChartSection_month").Row + 21).EntireRow.Hidden = False
        .Protect
    End With
End Sub
Sub HideMonthlyChart()
    With wsStart
        .Unprotect
        .Rows(wsStart.Range("rngChartSection_month").Row & ":" & wsStart.Range("rngChartSection_month").Row + 21).EntireRow.Hidden = True
        .Protect
    End With
End Sub


Sub Change_colorRed_button_Caculate(Optional blnFromStrat = False)
    If fnGetRangeFromNames(ThisWorkbook, "rng_Run_From_Indata") Then Exit Sub
    If Not fnGetRangeFromNames(ThisWorkbook, "rng_Change_button_color") Then Exit Sub
    On Error Resume Next
    wsStart.Select
    Call Unprotect_sheet(wsStart.CodeName)
    wsStart.Shapes("cmd_Cal").Visible = msoTrue

    wsStart.Shapes.Range(Array("cmd_Cal")).Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Fet"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
    End With
    If Not blnFromStrat Then Range("A2").Select
    Call Protect_sheet(wsStart.CodeName)
    On Error GoTo 0
End Sub
Sub Change_colorBlack_button_Caculate(Optional blnFromStrat = False)

    If fnGetRangeFromNames(ThisWorkbook, "rng_Run_From_Indata") Then Exit Sub
    If Not fnGetRangeFromNames(ThisWorkbook, "rng_Change_button_color") Then Exit Sub
    On Error Resume Next
    wsStart.Select
    Call Unprotect_sheet(wsStart.CodeName)
    wsStart.Shapes("cmd_Cal").Visible = msoTrue
    wsStart.Shapes.Range(Array("cmd_Cal")).Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Fet"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 0
    End With
    If Not blnFromStrat Then Range("A2").Select
    Call Protect_sheet(wsStart.CodeName)
    On Error GoTo 0
End Sub

Sub ScrolllWin()
   'ActiveWindow.ScrollRow = 1
   Dim i As Integer
    For i = 1 To 9
        ActiveWindow.SmallScroll Down:=2
        Sleep 500
    Next i
    ActiveWindow.SmallScroll Down:=1
    Range("D23").Select
End Sub
Sub BackToStart()
    'I fliken indata_lista (för egen inkomstvektor)  finns knappen att gå tillbaka till startfliken
    wsStart.Activate
    fnGetRangeFromNames(ThisWorkbook, "rngModelHeader").Select
End Sub
Sub GoToCases()
    'I fliken indata_lista (för egen inkomstvektor)  finns knappen att gå till färdiga typisar (fliken Typfall)
    wsEgenInkomst.Visible = True
    wsEgenInkomst.Activate
    fnGetRangeFromNames(ThisWorkbook, "rng_EgenLönelista_Index").Select
End Sub

''''Sub FixControllSize()
''''    With wsStart
''''        .lstSysLang.Width = wsStart.Shapes("cmdAbout").Width
''''        .lstSysLang.Height = 28
''''        .lstSysLang.Left = 975.75
''''    End With
''''End Sub

Sub HideControll(ByVal frmName As Object, ByVal cnName As String)
    Dim Uf As UserForm
    Set Uf = frmName
    Uf.Controls(cnName).Visible = False
    Set Uf = Nothing
End Sub

Sub UnHideControll(ByVal frmName As Object, ByVal cnName As String)
    Dim Uf As UserForm
    Set Uf = frmName
    Uf.Controls(cnName).Visible = True
    Set Uf = Nothing
End Sub

Function fnCheckNumeric(ByVal strValue) As Boolean
    fnCheckNumeric = False
    If IsNumeric(strValue) Or Right(strValue, 1) = "%" Then
        fnCheckNumeric = True
    End If
End Function
Function fnCheckBlank(ByVal strValue) As Boolean
    fnCheckBlank = False
    If Len(strValue) = 0 Then
        fnCheckBlank = True
    End If
End Function
Sub Unprotect_Allsheets()
    Dim wk As Worksheet
    For Each wk In ThisWorkbook.Worksheets
        wk.Unprotect
    Next wk
End Sub
Sub Protect_sheet(ByVal strSheetCodeName)
    Dim wk As Worksheet
    For Each wk In ThisWorkbook.Worksheets
        If LCase(wk.CodeName) = LCase(strSheetCodeName) Then
            wk.Protect
        End If
    Next wk
End Sub
Sub Unprotect_sheet(ByVal strSheetCodeName)
    Dim wk As Worksheet
    For Each wk In ThisWorkbook.Worksheets
        If LCase(wk.CodeName) = LCase(strSheetCodeName) Then
            wk.Unprotect
        End If
    Next wk
'''End Sub
'''Sub Calculator()
'''    Dim Program As String
'''    Dim TaskID As Double
'''    Program = "calc.exe"
'''    On Error Resume Next
'''    AppActivate "Calculator"
'''    If Err <> 0 Then
'''    Err = 0
'''    TaskID = shell(Program, 1)
'''    If Err <> 0 Then MsgBox fngetMsgBoxtext(39) & Program
'''    End If
'''    On Error GoTo 0
End Sub
Sub UnhideAllSheets()
    Dim wk As Worksheet

    For Each wk In ThisWorkbook.Worksheets
        wk.Visible = xlSheetVisible
    Next wk
End Sub
Sub Application_Rest()
    If Not Application.Calculation = xlCalculationManual Then Application.Calculation = xlCalculationManual
    If ThisWorkbook.Names("rngTurboMode").RefersToRange.Value Then
        If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
    End If
        

End Sub
Sub Application_Rest_Screen()
    If ThisWorkbook.Names("rngTurboMode").RefersToRange.Value Then
        If Application.ScreenUpdating = True Then
            Application.ScreenUpdating = False
        End If
    End If
End Sub
Sub Application_Wakeup_Screen()
    If Application.ScreenUpdating = False Then
        Application.ScreenUpdating = True
    End If
End Sub
Sub Application_Wakeup()
    If Not Application.Calculation = xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    If Application.ScreenUpdating = False Then Application.ScreenUpdating = True

End Sub
Sub Zoom_In()
    ActiveWindow.Zoom = ActiveWindow.Zoom + 10
End Sub
Sub Zoom_Out()
    ActiveWindow.Zoom = ActiveWindow.Zoom - 10
End Sub
