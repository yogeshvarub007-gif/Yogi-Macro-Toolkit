Attribute VB_Name = "Module_Transpose"
Sub FlattenSelectedRangeWithGaps_KeepFormatting_Fast()
    Dim selectedRange As Range
    Dim outputSheet As Worksheet
    Dim c As Long
    Dim outRow As Long
    Dim colRange As Range

    On Error Resume Next
    Set selectedRange = Application.Selection
    On Error GoTo 0

    If selectedRange Is Nothing Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Create or clear the "FlattenedData" sheet
    On Error Resume Next
    Set outputSheet = Worksheets("FlattenedData")
    If outputSheet Is Nothing Then
        Set outputSheet = Worksheets.Add
        outputSheet.Name = "FlattenedData"
    Else
        outputSheet.Cells.Clear
    End If
    On Error GoTo 0

    outRow = 1

    ' Copy one column block at a time instead of each cell
    For c = 1 To selectedRange.Columns.Count
        Set colRange = selectedRange.Columns(c)
        colRange.Copy
        outputSheet.Cells(outRow, 1).PasteSpecial Paste:=xlPasteAll
        outRow = outRow + colRange.Rows.Count + 1
    Next c

    Application.CutCopyMode = False
    outputSheet.Columns("A:A").AutoFit

    ' Re-enable Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Data flattened into 'FlattenedData' sheet with formatting preserved.", vbInformation
End Sub

