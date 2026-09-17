Attribute VB_Name = "Module_datalookup"
Option Explicit

' ============================================================
'  Universal_Data_Lookup  v3.2
'  Compatible : Excel 2016 and later (Windows)
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
    Dim sSrcReturnCol       As String
    Dim sSrcStartRow        As String
    Dim sSearchChoice       As String
    Dim sSheetChoice        As String
    Dim sSheetList          As String

    ' ---- Numeric inputs ----
    Dim lActiveStartRow     As Long
    Dim lSrcStartRow        As Long
    Dim lActiveLookupColNum As Long
    Dim lActiveOutputColNum As Long
    Dim lSrcLookupColNum    As Long
    Dim lSrcReturnColNum    As Long
    Dim lActiveLastRow      As Long

    ' ---- Counters ----
    Dim lMatchCount         As Long
    Dim lNotFoundCount      As Long
    Dim i                   As Long

    ' ---- Lookup variables ----
    Dim sLookupKey          As String
    Dim sOutputVal          As String

    ' ---- Dictionaries ----
    Dim dictLookup          As Object   ' key -> return value
    Dim dictFirst           As Object   ' key -> first-occurrence metadata
    Dim dictDuplicateKeys   As Object   ' keys that have duplicates (for filtering report)

    ' ---- Duplicate report data ----
    Dim colDupRows          As Collection
    Dim bHasDuplicates      As Boolean

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
    sActiveLookupCol = InputBox("Enter the LOOKUP column letter in the ACTIVE sheet." & vbNewLine & "Example: E", _
                                "Active Sheet – Lookup Column", "E")
    If sActiveLookupCol = "" Then GoTo UserCancelled
    sActiveLookupCol = UCase(Trim(sActiveLookupCol))
    lActiveLookupColNum = ColLetterToNumber(sActiveLookupCol)
    If lActiveLookupColNum = 0 Then
        MsgBox "Invalid column letter '" & sActiveLookupCol & "'.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sActiveOutputCol = InputBox("Enter the OUTPUT column letter in the ACTIVE sheet." & vbNewLine & "Example: F", _
                                "Active Sheet – Output Column", "F")
    If sActiveOutputCol = "" Then GoTo UserCancelled
    sActiveOutputCol = UCase(Trim(sActiveOutputCol))
    lActiveOutputColNum = ColLetterToNumber(sActiveOutputCol)
    If lActiveOutputColNum = 0 Then
        MsgBox "Invalid column letter '" & sActiveOutputCol & "'.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    If lActiveLookupColNum = lActiveOutputColNum Then
        MsgBox "Lookup column and Output column cannot be the same.", vbCritical, "Universal Data Lookup"
        GoTo SafeCleanup
    End If

    sActiveStartRow = InputBox("Enter the STARTING ROW in the ACTIVE sheet." & vbNewLine & "Example: 2", _
                               "Active Sheet – Starting Row", "2")
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
                             "Source Workbook – Lookup Column", "B")
    If sSrcLookupCol = "" Then GoTo UserCancelled
    sSrcLookupCol = UCase(Trim(sSrcLookupCol))
    lSrcLookupColNum = ColLetterToNumber(sSrcLookupCol)
    If lSrcLookupColNum = 0 Then GoTo SafeCleanup

    sSrcReturnCol = InputBox("Enter the RETURN column letter in the SOURCE workbook." & vbNewLine & "Example: C", _
                             "Source Workbook – Return Column", "C")
    If sSrcReturnCol = "" Then GoTo UserCancelled
    sSrcReturnCol = UCase(Trim(sSrcReturnCol))
    lSrcReturnColNum = ColLetterToNumber(sSrcReturnCol)
    If lSrcReturnColNum = 0 Then GoTo SafeCleanup

    sSrcStartRow = InputBox("Enter the STARTING ROW in the SOURCE workbook." & vbNewLine & "Example: 2", _
                            "Source Workbook – Starting Row", "2")
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
                                         lSrcLookupColNum, lSrcReturnColNum, lSrcStartRow, bHasDuplicates)
        Next wsLoop
    Else
        Call LoadSheetIntoDictionary(wsSource, dictLookup, dictFirst, dictDuplicateKeys, colDupRows, _
                                     lSrcLookupColNum, lSrcReturnColNum, lSrcStartRow, bHasDuplicates)
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
            wsActive.Cells(i, lActiveOutputColNum).Value = sOutputVal
            wsActive.Cells(i, lActiveOutputColNum).Interior.Color = xlNone
            lMatchCount = lMatchCount + 1
        Else
            With wsActive.Cells(i, lActiveOutputColNum)
                .Value = "Not Found"
                .Interior.Color = RGB(255, 255, 0)
            End With
            lNotFoundCount = lNotFoundCount + 1
        End If
NextRow:
    Next i

    ' -------------------------------------------------------
    '  STEP 9 – Write Duplicate_Report (only for used keys)
    ' -------------------------------------------------------
    If bHasDuplicates Then
        Call WriteDuplicateReport(wbTarget, colDupRows, dictUsedKeys)
    End If

    ' -------------------------------------------------------
    '  STEP 10 – Summary
    ' -------------------------------------------------------
    Dim sSummary As String
    sSummary = "Universal Data Lookup – Complete!" & vbNewLine & vbNewLine & _
               "Rows Processed : " & (lMatchCount + lNotFoundCount) & vbNewLine & _
               "Matches Found  : " & lMatchCount & vbNewLine & _
               "Not Found      : " & lNotFoundCount

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
        MsgBox "Error " & lErrNum & ": " & sErrDesc, vbCritical, "Universal Data Lookup – Error"
    End If
End Sub

' ============================================================
'  LoadSheetIntoDictionary - Pure key-based, no filtering
' ============================================================
Private Sub LoadSheetIntoDictionary( _
        ByVal ws As Worksheet, _
        ByRef dictLookup As Object, _
        ByRef dictFirst As Object, _
        ByRef dictDuplicateKeys As Object, _
        ByRef colDupRows As Collection, _
        ByVal lLookupCol As Long, _
        ByVal lReturnCol As Long, _
        ByVal lStartRow As Long, _
        ByRef bHasDuplicates As Boolean)

    Dim lLastRow As Long, i As Long
    Dim sKey As String, sValue As String
    Dim sFirstMeta As String, arrMeta() As String
    Dim sDupRecord As String

    lLastRow = ws.Cells(ws.Rows.Count, lLookupCol).End(xlUp).Row
    If lLastRow < lStartRow Then Exit Sub

    For i = lStartRow To lLastRow
        sKey = NormalizeValue(ws.Cells(i, lLookupCol).Value)
        If sKey = "" Then GoTo NextRow

        sValue = CStr(ws.Cells(i, lReturnCol).Value)

        If dictLookup.Exists(sKey) Then
            bHasDuplicates = True
            dictDuplicateKeys(sKey) = True   ' Mark key as having duplicates

            sFirstMeta = CStr(dictFirst(sKey))
            arrMeta = Split(sFirstMeta, "|")

            sDupRecord = sKey & "|" & arrMeta(0) & "|" & arrMeta(1) & "|" & _
                         ws.Name & "|" & CStr(i) & "|" & sValue
            colDupRows.Add sDupRecord
        Else
            dictLookup.Add sKey, sValue
            dictFirst.Add sKey, ws.Name & "|" & CStr(i) & "|" & sValue
        End If
NextRow:
    Next i
End Sub

' ============================================================
'  WriteDuplicateReport - FILTERED to only used keys
' ============================================================
Private Sub WriteDuplicateReport( _
        ByVal wbTarget As Workbook, _
        ByVal colDupRows As Collection, _
        ByVal dictUsedKeys As Object)

    Const SHEET_NAME As String = "Duplicate_Report"
    Dim wsDup As Worksheet
    Dim lRow As Long, v As Variant
    Dim arrParts() As String

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
        .Cells(1, 6).Value = "Return Value"

        With .Range(.Cells(1, 1), .Cells(1, 6))
            .Font.Bold = True
            .Interior.Color = RGB(0, 70, 127)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        lRow = 2
        For Each v In colDupRows
            arrParts = Split(CStr(v), "|")
            If UBound(arrParts) >= 5 Then
                If dictUsedKeys.Exists(arrParts(0)) Then   ' <-- Only include if key was used in lookup
                    .Cells(lRow, 1).Value = arrParts(0)
                    .Cells(lRow, 2).Value = arrParts(1)
                    .Cells(lRow, 3).Value = CLng(arrParts(2))
                    .Cells(lRow, 4).Value = arrParts(3)
                    .Cells(lRow, 5).Value = CLng(arrParts(4))
                    .Cells(lRow, 6).Value = arrParts(5)

                    If lRow Mod 2 = 0 Then
                        .Range(.Cells(lRow, 1), .Cells(lRow, 6)).Interior.Color = RGB(255, 242, 204)
                    End If
                    lRow = lRow + 1
                End If
            End If
        Next v

        .Columns("A:F").AutoFit
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

