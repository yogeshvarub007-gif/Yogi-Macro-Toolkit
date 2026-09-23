Sub TransposeVerticalGroupsToRows()
    '=====================================================================
    ' Source: ONE column of data, e.g. A1:A1000
    ' User enters a group size N (e.g. 10).
    ' Every N consecutive source rows become ONE horizontal row in the output:
    '     A1:A10   -> A1:J1   (row 1)
    '     A11:A20  -> A2:J2   (row 2)
    '     A21:A30  -> A3:J3   (row 3)
    '     ...
    ' Output goes to a NEW worksheet starting at A1.
    ' The original/source sheet is never modified.
    ' Works for any group size and any number of selected rows
    ' (970 rows / group size 10 -> 97 output rows x 10 columns).
    ' Preserves exact text (e.g. leading zeros) and leaves blanks blank.
    '=====================================================================

    Dim srcRange As Range
    Dim srcCell As Range
    Dim groupSize As Long
    Dim totalRows As Long
    Dim newWs As Worksheet
    Dim i As Long, j As Long
    Dim outRow As Long, outCol As Long
    Dim srcIndex As Long
    Dim cellVal As Variant
    Dim isTextFmt As Boolean
    Dim inputStr As String
    Dim destCell As Range

    On Error GoTo ErrHandler

    ' 1. Validate selection --------------------------------------------------
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If

    Set srcRange = Selection

    If srcRange.Columns.Count > 1 Then
        MsgBox "Please select a SINGLE column of data (one column, multiple rows).", vbExclamation
        Exit Sub
    End If

    totalRows = srcRange.Rows.Count

    If totalRows < 1 Then
        MsgBox "The selected range has no rows.", vbExclamation
        Exit Sub
    End If

    ' 2. Ask for group size ---------------------------------------------------
    inputStr = InputBox("You selected " & totalRows & " row(s)." & vbCrLf & _
                         "Enter the group size (rows per individual / rows to transpose per output row):", _
                         "Transpose Vertical Groups To Rows")

    If Trim(inputStr) = "" Then
        Exit Sub ' user cancelled
    End If

    If Not IsNumeric(inputStr) Then
        MsgBox "Please enter a valid whole number.", vbExclamation
        Exit Sub
    End If

    groupSize = CLng(inputStr)

    If groupSize < 1 Then
        MsgBox "Group size must be at least 1.", vbExclamation
        Exit Sub
    End If

    ' 3. Create a NEW worksheet for the output (source sheet untouched) ------
    Application.ScreenUpdating = False

    Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error Resume Next
    newWs.Name = "Transposed_" & Format(Now, "hhmmss")
    On Error GoTo ErrHandler

    ' 4. Loop through source cells, groupSize rows at a time ------------------
    outRow = 0
    For i = 1 To totalRows Step groupSize
        outRow = outRow + 1
        outCol = 0

        For j = 0 To groupSize - 1
            srcIndex = i + j
            outCol = outCol + 1

            If srcIndex <= totalRows Then
                Set srcCell = srcRange.Cells(srcIndex, 1)
                Set destCell = newWs.Cells(outRow, outCol)

                ' Force destination to Text format FIRST so leading zeros /
                ' exact strings are preserved on entry.
                destCell.NumberFormat = "@"

                If IsEmpty(srcCell.Value) Or srcCell.Value = "" Then
                    destCell.Value = ""
                Else
                    isTextFmt = (srcCell.NumberFormat = "@")
                    If isTextFmt Then
                        cellVal = srcCell.Value      ' exact stored string, e.g. "0012345"
                    Else
                        cellVal = srcCell.Text       ' fallback: displayed string
                    End If
                    destCell.Value = cellVal
                End If
            End If
            ' If srcIndex > totalRows (incomplete final group), that output
            ' cell is simply left blank — no error, no data corruption.
        Next j
    Next i

    ' 5. Tidy up ---------------------------------------------------------------
    newWs.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Done! " & totalRows & " row(s) converted into " & outRow & " row(s) x " & _
           groupSize & " column(s) on sheet """ & newWs.Name & """.", vbInformation

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub
