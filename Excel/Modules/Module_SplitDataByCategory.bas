Attribute VB_Name = "Module_SplitDataByCategory"
Option Explicit

'-----------------------------------------
' Check if a sheet exists
'-----------------------------------------
Function SheetExists(shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(shtName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'-----------------------------------------
' Main Macro
'-----------------------------------------
Sub SplitDataByCategory_FINAL_STABLE()

    Dim srcWS As Worksheet, tgtWS As Worksheet
    Dim headerRow As Long, catCol As Long
    Dim LastRow As Long, lastCol As Long
    Dim hdrInput As String, colInput As String
    Dim dict As Object
    Dim headerArr As Variant, dataArr As Variant
    Dim i As Long, r As Long
    Dim category As String, SheetName As String
    Dim key As Variant
    Dim NextRow As Long, slCol As Long
    Dim hdrText As String
    Dim lastDataRow As Long

    Set srcWS = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    '-------------------------------
    ' Header row input
    '-------------------------------
    hdrInput = InputBox("Enter the HEADER row number:", "Header Row")
    If hdrInput = "" Or Not IsNumeric(hdrInput) Then Exit Sub
    headerRow = CLng(hdrInput)

    '-------------------------------
    ' Category column input
    '-------------------------------
    colInput = InputBox("Enter CATEGORY column letter:", "Category Column")
    If colInput = "" Then Exit Sub
    catCol = srcWS.Range(UCase(colInput) & headerRow).Column

    '-------------------------------
    ' Detect bounds
    '-------------------------------
    LastRow = srcWS.Cells(srcWS.Rows.Count, catCol).End(xlUp).Row
    lastCol = srcWS.Cells(headerRow, srcWS.Columns.Count).End(xlToLeft).Column

    '-------------------------------
    ' Load data into memory (FAST)
    '-------------------------------
    headerArr = srcWS.Range(srcWS.Cells(headerRow, 1), srcWS.Cells(headerRow, lastCol)).Value
    dataArr = srcWS.Range(srcWS.Cells(headerRow + 1, 1), srcWS.Cells(LastRow, lastCol)).Value

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '-------------------------------
    ' Split data
    '-------------------------------
    For i = 1 To UBound(dataArr, 1)

        category = Trim(dataArr(i, catCol))
        If category = "" Then GoTo NextRow

        ' Clean sheet name
        SheetName = category
        SheetName = Replace(SheetName, "/", "-")
        SheetName = Replace(SheetName, "\", "-")
        SheetName = Replace(SheetName, ":", "-")
        SheetName = Replace(SheetName, "*", "")
        SheetName = Replace(SheetName, "?", "")
        SheetName = Replace(SheetName, "[", "")
        SheetName = Replace(SheetName, "]", "")
        SheetName = Left(SheetName, 31)

        If SheetName = "" Then GoTo NextRow

        ' Create sheet if needed
        If Not dict.Exists(SheetName) Then
            If Not SheetExists(SheetName) Then
                Set tgtWS = Sheets.Add(After:=Sheets(Sheets.Count))
                tgtWS.Name = SheetName

                ' Header values + formats
                tgtWS.Range("A1").Resize(1, lastCol).Value = headerArr
                srcWS.Rows(headerRow).Copy
                tgtWS.Rows(1).PasteSpecial xlPasteFormats
            Else
                Set tgtWS = Sheets(SheetName)
            End If
            dict.Add SheetName, SheetName
        End If

        Set tgtWS = Sheets(dict(SheetName))
        NextRow = tgtWS.Cells(tgtWS.Rows.Count, 1).End(xlUp).Row + 1
        tgtWS.Cells(NextRow, 1).Resize(1, lastCol).Value = Application.Index(dataArr, i, 0)

NextRow:
    Next i

    Application.CutCopyMode = False

    '-------------------------------
    ' Post-processing per sheet
    '-------------------------------
    For Each key In dict.Keys

        Set tgtWS = Sheets(key)

        ' Detect Serial Number column (robust)
        slCol = 0
        For i = 1 To lastCol
            hdrText = UCase(Trim(Replace(tgtWS.Cells(1, i).Value, Chr(160), " ")))
            If hdrText Like "*SL*NO*" Or hdrText Like "*SERIAL*" Then
                slCol = i
                Exit For
            End If
        Next i

        ' Re-number Serial Number
        If slCol > 0 Then
            For r = 2 To tgtWS.Cells(tgtWS.Rows.Count, slCol).End(xlUp).Row
                tgtWS.Cells(r, slCol).Value = r - 1
            Next r
        End If

        ' Apply body formatting (exclude header)
        lastDataRow = tgtWS.Cells(tgtWS.Rows.Count, 1).End(xlUp).Row

        If lastDataRow >= 2 Then
            With tgtWS.Range(tgtWS.Cells(2, 1), tgtWS.Cells(lastDataRow, lastCol))
                .Font.Name = "Times New Roman"
                .Font.Size = 12

                ' Full borders
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            End With
        End If

        ' AutoFit columns
        tgtWS.Columns.AutoFit

    Next key

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Data split completed successfully.", vbInformation

End Sub

