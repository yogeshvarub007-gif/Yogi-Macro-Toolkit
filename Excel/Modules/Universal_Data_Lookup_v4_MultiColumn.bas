Option Explicit

' ============================================================
'  Universal_Data_Lookup  v4.0  (Multi-Column Return)
'  Compatible : Excel 2016 and later (Windows)
'
'  CHANGE FROM v3.2:
'  - Source RETURN is now a COLUMN RANGE (e.g. B:D) instead of
'    a single column. All columns in that range are pulled back
'    for a matching key.
'  - Active sheet OUTPUT is now a STARTING column (e.g. E). The
'    returned columns are written consecutively starting there
'    (e.g. Source B:D -> Active E:G).
'  - Internal record delimiters switched from "|" to the ASCII
'    Record/Unit Separator characters (Chr(30)/Chr(31)) so that
'    "|" appearing inside real cell data can no longer corrupt
'    the duplicate-tracking records.
' ============================================================

' -------------------------------------------------------
'  ENTRY POINT
' -------------------------------------------------------
Public Sub Universal_Data_Lookup()

    ' ---- Object variables ----
    Dim wsActive        As Worksheet
    Dim wbSource        As Workbook
    Dim wbTarget        As Workbook
    Dim wsSource        As Worksheet

    ' ---- String inputs ----
    Dim sFilePath           As String
    Dim sActiveLookupCol    As String
    Dim sActiveOutputCol    As String
    Dim sActiveStartRow     As String
    Dim sSrcLookupCol       As String
    Dim sSrcReturnRange     As String
    Dim sSrcStartRow        As String
    Dim sSearchChoice       As String
    Dim sSheetChoice        As String
    Dim sSheetList          As String

    ' ---- Numeric inputs ----
    Dim lActiveStartRow     As Long
    Dim lSrcStartRow        As Long
    Dim lActiveLookupColNum As Long
    Dim lActiveOutputColStart As Long
    Dim lActiveOutputColEnd   As Long
    Dim lSrcLookupColNum    As Long
    Dim lSrcReturnColStart  As Long
    Dim lSrcReturnColEnd    As Long
    Dim lReturnColCount     As Long
    Dim lActiveLastRow      As Long

    ' ---- Counters ----
    Dim lMatchCount         As Long
    Dim lNotFoundCount      As Long
    Dim i                   As Long
    Dim j                   As Long

    ' ---- Lookup variables ----
    Dim sLookupKey          As String
    Dim sOutputVal          As String
    Dim arrOut()            As String

    ' ---- Dictionaries ----
    Dim dictLookup          As Object   ' key -> delimited return values
    Dim dictFirst           As Object   ' key -> first-occurrence metadata
    Dim dictDuplicateKeys   As Object   ' keys that have duplicates (for filtering report)

    ' ---- Duplicate report data ----
    Dim colDupRows          As Collection
    Dim bHasDuplicates      As Boolean

    ' ---- Delimiters (kept out of user data's way) ----
    Dim sRecSep As String
    Dim sUnitSep As String
    sRecSep = Chr(30)   ' separates fields within an internal record
    sUnitSep = Chr(31)  ' separates individual returned column values within a record field

    ' ---- Excel state ----
    Dim bCalcWasAuto        As Boolean
    Dim bEventsWereOn       As Boolean
    Dim bScreenWasOn        As Boolean

    ' -------------------------------------------------------
    '  STEP 0 – Performance settings
    ' -------------------------------------------------------
    bCalcWasAuto = (Application.Calculation = xlCalculationAutomatic)
    bEventsWereOn = Application.EnableEvents
    bScreenWasOn = Application.ScreenUpdating

    On Error GoTo SafeCleanup

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsActive = ActiveSheet
    Set wbTarget = wsActive.Parent

    If wsActive Is Nothing Then
        MsgBox "No active worksheet found.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' -------------------------------------------------------
    '  STEP 1 – Active Sheet inputs
    ' -------------------------------------------------------
    sActiveLookupCol = InputBox("Enter the LOOKUP column letter in the ACTIVE sheet." & vbNewLine & "Example: A", _
                                "Active Sheet - Lookup Column", "A")
    If sActiveLookupCol = "" Then GoTo UserCancelled
    sActiveLookupCol = UCase(Trim(sActiveLookupCol))
    lActiveLookupColNum = ColLetterToNumber(sActiveLookupCol)
    If lActiveLookupColNum = 0 Then
        MsgBox "Invalid column letter '" & sActiveLookupCol & "'.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sActiveOutputCol = InputBox("Enter the STARTING OUTPUT column letter in the ACTIVE sheet." & vbNewLine & _
                                "The returned columns will be written consecutively starting here." & vbNewLine & _
                                "Example: E", "Active Sheet - Output Starting Column", "E")
    If sActiveOutputCol = "" Then GoTo UserCancelled
    sActiveOutputCol = UCase(Trim(sActiveOutputCol))
    lActiveOutputColStart = ColLetterToNumber(sActiveOutputCol)
    If lActiveOutputColStart = 0 Then
        MsgBox "Invalid column letter '" & sActiveOutputCol & "'.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sActiveStartRow = InputBox("Enter the STARTING ROW in the ACTIVE sheet." & vbNewLine & "Example: 2", _
                               "Active Sheet - Starting Row", "2")
    If sActiveStartRow = "" Then GoTo UserCancelled
    If Not IsNumeric(sActiveStartRow) Then
        MsgBox "Starting row must be a number.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If
    lActiveStartRow = CLng(sActiveStartRow)
    If lActiveStartRow < 1 Then
        MsgBox "Starting row must be at least 1.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' -------------------------------------------------------
    '  STEP 2 – Source File
    ' -------------------------------------------------------
    Application.ScreenUpdating = True
    sFilePath = GetSourceFilePath()
    Application.ScreenUpdating = False

    If sFilePath = "" Then
        MsgBox "No source file selected. Operation cancelled.", vbInformation, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' -------------------------------------------------------
    '  STEP 3 – Source inputs
    ' -------------------------------------------------------
    sSrcLookupCol = InputBox("Enter the LOOKUP column letter in the SOURCE workbook." & vbNewLine & "Example: B", _
                             "Source Workbook - Lookup Column", "B")
    If sSrcLookupCol = "" Then GoTo UserCancelled
    sSrcLookupCol = UCase(Trim(sSrcLookupCol))
    lSrcLookupColNum = ColLetterToNumber(sSrcLookupCol)
    If lSrcLookupColNum = 0 Then
        MsgBox "Invalid column letter '" & sSrcLookupCol & "'.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sSrcReturnRange = InputBox("Enter the RETURN column(s) in the SOURCE workbook." & vbNewLine & _
                               "Single column example: C" & vbNewLine & _
                               "Multiple consecutive columns example: C:E", _
                               "Source Workbook - Return Column(s)", "C:E")
    If sSrcReturnRange = "" Then GoTo UserCancelled

    If Not ParseColumnRange(sSrcReturnRange, lSrcReturnColStart, lSrcReturnColEnd) Then
        MsgBox "Invalid return column range '" & sSrcReturnRange & "'." & vbNewLine & _
               "Use a single column (e.g. C) or a range (e.g. C:E) with the start column at or before the end column.", _
               vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    lReturnColCount = lSrcReturnColEnd - lSrcReturnColStart + 1
    lActiveOutputColEnd = lActiveOutputColStart + lReturnColCount - 1

    If lActiveOutputColEnd > 16384 Then
        MsgBox "The output range extends past the last worksheet column (XFD).", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' Overlap check - lookup column must not sit inside the output range
    If lActiveLookupColNum >= lActiveOutputColStart And lActiveLookupColNum <= lActiveOutputColEnd Then
        MsgBox "The active sheet lookup column cannot fall inside the output column range (" & _
               ColNumberToLetter(lActiveOutputColStart) & ":" & ColNumberToLetter(lActiveOutputColEnd) & ").", _
               vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sSrcStartRow = InputBox("Enter the STARTING ROW in the SOURCE workbook." & vbNewLine & "Example: 2", _
                            "Source Workbook - Starting Row", "2")
    If sSrcStartRow = "" Then GoTo UserCancelled
    If Not IsNumeric(sSrcStartRow) Then
        MsgBox "Source starting row must be a number.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If
    lSrcStartRow = CLng(sSrcStartRow)
    If lSrcStartRow < 1 Then GoTo SafeCleanup

    ' -------------------------------------------------------
    '  STEP 4 – Search scope
    ' -------------------------------------------------------
    sSearchChoice = UCase(Trim(InputBox("Type A for ALL sheets" & vbNewLine & "Type S for SPECIFIC sheet", _
                                        "Source Search Scope", "A")))
    If sSearchChoice <> "A" And sSearchChoice <> "S" Then
        MsgBox "Invalid choice. Please enter A or S.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' -------------------------------------------------------
    '  STEP 5 – Open Source
    ' -------------------------------------------------------
    On Error Resume Next
    Set wbSource = Workbooks.Open(sFilePath, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo SafeCleanup

    If wbSource Is Nothing Then
        MsgBox "Could not open source workbook.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    ' -------------------------------------------------------
    '  STEP 6 – Specific sheet
    ' -------------------------------------------------------
    If sSearchChoice = "S" Then
        sSheetList = GetSheetList(wbSource)
        sSheetChoice = InputBox("Available sheets:" & vbNewLine & vbNewLine & sSheetList & vbNewLine & _
                                "Enter the exact sheet name:", "Select Source Sheet", "")
        If sSheetChoice = "" Then
            wbSource.Close SaveChanges:=False
            GoTo UserCancelled
        End If

        On Error Resume Next
        Set wsSource = wbSource.Worksheets(sSheetChoice)
        On Error GoTo SafeCleanup

        If wsSource Is Nothing Then
            MsgBox "Sheet '" & sSheetChoice & "' not found.", vbCritical, "Universal Data Lookup"
            wbSource.Close SaveChanges:=False
            GoTo SafeCleanup
        End If
    End If

    ' -------------------------------------------------------
    '  STEP 7 – Build Dictionary + track duplicate keys
    ' -------------------------------------------------------
    Set dictLookup = CreateObject("Scripting.Dictionary")
    dictLookup.CompareMode = 1

    Set dictFirst = CreateObject("Scripting.Dictionary")
    dictFirst.CompareMode = 1

    Set dictDuplicateKeys = CreateObject("Scripting.Dictionary")
    dictDuplicateKeys.CompareMode = 1

    Set colDupRows = New Collection
    bHasDuplicates = False

    If sSearchChoice = "A" Then
        Dim wsLoop As Worksheet
        For Each wsLoop In wbSource.Worksheets
            Call LoadSheetIntoDictionary(wsLoop, dictLookup, dictFirst, dictDuplicateKeys, colDupRows, _
                                         lSrcLookupColNum, lSrcReturnColStart, lReturnColCount, lSrcStartRow, _
                                         bHasDuplicates, sRecSep, sUnitSep)
        Next wsLoop
    Else
        Call LoadSheetIntoDictionary(wsSource, dictLookup, dictFirst, dictDuplicateKeys, colDupRows, _
                                     lSrcLookupColNum, lSrcReturnColStart, lReturnColCount, lSrcStartRow, _
                                     bHasDuplicates, sRecSep, sUnitSep)
    End If

    If dictLookup.Count = 0 Then
        MsgBox "No data found in source workbook.", vbCritical, "Universal Data Lookup"
        wbSource.Close SaveChanges:=False
        GoTo SafeCleanup
    End If

    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Set wsSource = Nothing

    ' -------------------------------------------------------
    '  STEP 8 – Perform lookup on Active Sheet + collect used keys
    ' -------------------------------------------------------
    Dim dictUsedKeys As Object
    Set dictUsedKeys = CreateObject("Scripting.Dictionary")
    dictUsedKeys.CompareMode = 1

    lActiveLastRow = wsActive.Cells(wsActive.Rows.Count, lActiveLookupColNum).End(xlUp).Row

    If lActiveLastRow < lActiveStartRow Then
        MsgBox "No data found in active sheet lookup column.", vbInformation, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    lMatchCount = 0
    lNotFoundCount = 0

    For i = lActiveStartRow To lActiveLastRow
        sLookupKey = NormalizeValue(wsActive.Cells(i, lActiveLookupColNum).Value)
        If sLookupKey = "" Then GoTo NextRow

        dictUsedKeys(sLookupKey) = True   ' Track keys actually used in lookup

        If dictLookup.Exists(sLookupKey) Then
            sOutputVal = CStr(dictLookup(sLookupKey))
            arrOut = Split(sOutputVal, sUnitSep)
            For j = 0 To lReturnColCount - 1
                With wsActive.Cells(i, lActiveOutputColStart + j)
                    If j <= UBound(arrOut) Then
                        .Value = arrOut(j)
                    Else
                        .Value = ""
                    End If
                    .Interior.Color = xlNone
                End With
            Next j
            lMatchCount = lMatchCount + 1
        Else
            For j = 0 To lReturnColCount - 1
                With wsActive.Cells(i, lActiveOutputColStart + j)
                    .Value = "Not Found"
                    .Interior.Color = RGB(255, 255, 0)
                End With
            Next j
            lNotFoundCount = lNotFoundCount + 1
        End If
NextRow:
    Next i

    ' -------------------------------------------------------
    '  STEP 9 – Write Duplicate_Report (only for used keys)
    ' -------------------------------------------------------
    If bHasDuplicates Then
        Call WriteDuplicateReport(wbTarget, colDupRows, dictUsedKeys, lReturnColCount, sRecSep, sUnitSep)
    End If

    ' -------------------------------------------------------
    '  STEP 10 – Summary
    ' -------------------------------------------------------
    Dim sSummary As String
    sSummary = "Universal Data Lookup - Complete!" & vbNewLine & vbNewLine & _
               "Rows Processed : " & (lMatchCount + lNotFoundCount) & vbNewLine & _
               "Matches Found  : " & lMatchCount & vbNewLine & _
               "Not Found      : " & lNotFoundCount & vbNewLine & _
               "Columns Returned : " & lReturnColCount & " (" & ColNumberToLetter(lActiveOutputColStart) & _
               ":" & ColNumberToLetter(lActiveOutputColEnd) & ")"

    If bHasDuplicates Then
        sSummary = sSummary & vbNewLine & vbNewLine & _
                   "See Duplicate_Report sheet for relevant duplicate key details."
    End If

    MsgBox sSummary, vbInformation, "Universal Data Lookup"

    Call RestoreExcelSettings(bCalcWasAuto, bEventsWereOn, bScreenWasOn)

    ' Cleanup
    Set dictLookup = Nothing
    Set dictFirst = Nothing
    Set dictDuplicateKeys = Nothing
    Set dictUsedKeys = Nothing
    Set colDupRows = Nothing
    Exit Sub

UserCancelled:
    MsgBox "Operation cancelled by user.", vbInformation, "Universal Data Lookup"
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Call RestoreExcelSettings(bCalcWasAuto, bEventsWereOn, bScreenWasOn)
    Exit Sub

SafeCleanup:
    Dim lErrNum As Long, sErrDesc As String
    lErrNum = Err.Number: sErrDesc = Err.Description
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Call RestoreExcelSettings(bCalcWasAuto, bEventsWereOn, bScreenWasOn)
    If lErrNum <> 0 Then
        MsgBox "Error " & lErrNum & ": " & sErrDesc, vbCritical, "Universal Data Lookup - Error"
    End If
End Sub

' ============================================================
'  LoadSheetIntoDictionary - Pure key-based, no filtering
'  Now captures a RANGE of return columns per key (joined by
'  sUnitSep) instead of a single value.
' ============================================================
Private Sub LoadSheetIntoDictionary( _
        ByVal ws As Worksheet, _
        ByRef dictLookup As Object, _
        ByRef dictFirst As Object, _
        ByRef dictDuplicateKeys As Object, _
        ByRef colDupRows As Collection, _
        ByVal lLookupCol As Long, _
        ByVal lReturnColStart As Long, _
        ByVal lReturnColCount As Long, _
        ByVal lStartRow As Long, _
        ByRef bHasDuplicates As Boolean, _
        ByVal sRecSep As String, _
        ByVal sUnitSep As String)

    Dim lLastRow As Long, i As Long, j As Long
    Dim sKey As String, sValue As String
    Dim sFirstMeta As String, arrMeta() As String
    Dim sDupRecord As String
    Dim arrVals() As String

    lLastRow = ws.Cells(ws.Rows.Count, lLookupCol).End(xlUp).Row
    If lLastRow < lStartRow Then Exit Sub

    For i = lStartRow To lLastRow
        sKey = NormalizeValue(ws.Cells(i, lLookupCol).Value)
        If sKey = "" Then GoTo NextRow

        ReDim arrVals(0 To lReturnColCount - 1)
        For j = 0 To lReturnColCount - 1
            arrVals(j) = CStr(ws.Cells(i, lReturnColStart + j).Value)
        Next j
        sValue = Join(arrVals, sUnitSep)

        If dictLookup.Exists(sKey) Then
            bHasDuplicates = True
            dictDuplicateKeys(sKey) = True   ' Mark key as having duplicates

            sFirstMeta = CStr(dictFirst(sKey))
            arrMeta = Split(sFirstMeta, sRecSep)

            sDupRecord = sKey & sRecSep & arrMeta(0) & sRecSep & arrMeta(1) & sRecSep & _
                         ws.Name & sRecSep & CStr(i) & sRecSep & sValue
            colDupRows.Add sDupRecord
        Else
            dictLookup.Add sKey, sValue
            dictFirst.Add sKey, ws.Name & sRecSep & CStr(i) & sRecSep & sValue
        End If
NextRow:
    Next i
End Sub

' ============================================================
'  WriteDuplicateReport - FILTERED to only used keys
'  Return value(s) now expand into lReturnColCount columns.
' ============================================================
Private Sub WriteDuplicateReport( _
        ByVal wbTarget As Workbook, _
        ByVal colDupRows As Collection, _
        ByVal dictUsedKeys As Object, _
        ByVal lReturnColCount As Long, _
        ByVal sRecSep As String, _
        ByVal sUnitSep As String)

    Const SHEET_NAME As String = "Duplicate_Report"
    Dim wsDup As Worksheet
    Dim lRow As Long, v As Variant
    Dim arrParts() As String
    Dim arrRetVals() As String
    Dim lFixedCols As Long, lTotalCols As Long
    Dim k As Long

    lFixedCols = 5   ' Lookup Key, First Sheet, First Row, Duplicate Sheet, Duplicate Row
    lTotalCols = lFixedCols + lReturnColCount

    On Error Resume Next
    Set wsDup = wbTarget.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If wsDup Is Nothing Then
        Set wsDup = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
        wsDup.Name = SHEET_NAME
    Else
        wsDup.Cells.Clear
    End If

    With wsDup
        .Cells(1, 1).Value = "Lookup Key"
        .Cells(1, 2).Value = "First Sheet"
        .Cells(1, 3).Value = "First Occurrence Row"
        .Cells(1, 4).Value = "Duplicate Sheet"
        .Cells(1, 5).Value = "Duplicate Row"
        For k = 1 To lReturnColCount
            .Cells(1, lFixedCols + k).Value = "Return Value " & k
        Next k

        With .Range(.Cells(1, 1), .Cells(1, lTotalCols))
            .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        lRow = 2
        For Each v In colDupRows
            arrParts = Split(CStr(v), sRecSep)
            If UBound(arrParts) >= 5 Then
                If dictUsedKeys.Exists(arrParts(0)) Then   ' <-- Only include if key was used in lookup
                    .Cells(lRow, 1).Value = arrParts(0)
                    .Cells(lRow, 2).Value = arrParts(1)
                    .Cells(lRow, 3).Value = CLng(arrParts(2))
                    .Cells(lRow, 4).Value = arrParts(3)
                    .Cells(lRow, 5).Value = CLng(arrParts(4))

                    arrRetVals = Split(arrParts(5), sUnitSep)
                    For k = 0 To lReturnColCount - 1
                        If k <= UBound(arrRetVals) Then
                            .Cells(lRow, lFixedCols + 1 + k).Value = arrRetVals(k)
                        End If
                    Next k

                    If lRow Mod 2 = 0 Then
                        .Range(.Cells(lRow, 1), .Cells(lRow, lTotalCols)).Interior.Color = RGB(255, 242, 204)
                    End If
                    lRow = lRow + 1
                End If
            End If
        Next v

        .Columns("A:" & ColNumberToLetter(lTotalCols)).AutoFit
        If lRow > 2 Then   ' Freeze only if there is data
            .Activate
            .Rows(2).Select
            ActiveWindow.FreezePanes = True
            .Cells(1, 1).Select
        End If
    End With

    Set wsDup = Nothing
End Sub

' ============================================================
'  Helper Functions
' ============================================================
Private Function NormalizeValue(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NormalizeValue = ""
    Else
        NormalizeValue = Trim(CStr(v))
    End If
End Function

Private Function ColLetterToNumber(ByVal sCol As String) As Long
    Dim n As Long, i As Integer, c As String
    sCol = UCase(Trim(sCol))
    If Len(sCol) < 1 Or Len(sCol) > 3 Then Exit Function
    For i = 1 To Len(sCol)
        c = mid(sCol, i, 1)
        If c < "A" Or c > "Z" Then Exit Function
        n = n * 26 + (Asc(c) - Asc("A") + 1)
    Next i
    If n >= 1 And n <= 16384 Then ColLetterToNumber = n
End Function

Private Function ColNumberToLetter(ByVal lCol As Long) As String
    Dim n As Long, s As String
    n = lCol
    Do While n > 0
        Dim rem_ As Long
        rem_ = (n - 1) Mod 26
        s = Chr(65 + rem_) & s
        n = (n - rem_ - 1) \ 26
    Loop
    ColNumberToLetter = s
End Function

' Parses a single column ("C") or a column range ("C:E") into
' start/end column numbers. Returns False if invalid.
Private Function ParseColumnRange(ByVal sRange As String, ByRef lStart As Long, ByRef lEnd As Long) As Boolean
    Dim arrParts() As String
    Dim sTrimmed As String

    sTrimmed = UCase(Trim(sRange))
    ParseColumnRange = False

    If InStr(sTrimmed, ":") > 0 Then
        arrParts = Split(sTrimmed, ":")
        If UBound(arrParts) <> 1 Then Exit Function
        lStart = ColLetterToNumber(Trim(arrParts(0)))
        lEnd = ColLetterToNumber(Trim(arrParts(1)))
    Else
        lStart = ColLetterToNumber(sTrimmed)
        lEnd = lStart
    End If

    If lStart = 0 Or lEnd = 0 Then Exit Function
    If lStart > lEnd Then Exit Function

    ParseColumnRange = True
End Function

Private Function GetSourceFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Source Workbook"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb", 1
        .AllowMultiSelect = False
        If .Show = True Then GetSourceFilePath = .SelectedItems(1)
    End With
    Set fd = Nothing
End Function

Private Function GetSheetList(ByVal wb As Workbook) As String
    Dim ws As Worksheet, sList As String, n As Integer
    For Each ws In wb.Worksheets
        n = n + 1
        sList = sList & "  " & n & ".  " & ws.Name & vbNewLine
    Next ws
    GetSheetList = sList
End Function

Private Sub RestoreExcelSettings(ByVal bCalcWasAuto As Boolean, ByVal bEventsWereOn As Boolean, ByVal bScreenWasOn As Boolean)
    On Error Resume Next
    If bCalcWasAuto Then Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = bEventsWereOn
    Application.ScreenUpdating = bScreenWasOn
    On Error GoTo 0
End Sub
