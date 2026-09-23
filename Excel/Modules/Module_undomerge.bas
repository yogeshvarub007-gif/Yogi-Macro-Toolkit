Attribute VB_Name = "Module_undomerge"
Sub UndoMergeRows()

    Dim rng As Range
    Dim cell As Range
    Dim mergeArea As Range
    Dim splitText As Variant
    Dim i As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim rowCount As Long

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set rng = Selection

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each cell In rng.Cells

        If cell.MergeCells Then

            'Store the merged area information BEFORE unmerging
            Set mergeArea = cell.mergeArea

            startRow = mergeArea.Row
            startCol = mergeArea.Column
            rowCount = mergeArea.Rows.Count

            'Store the text before unmerging
            splitText = Split(CStr(cell.Value), vbLf)

            'Unmerge
            mergeArea.UnMerge

            'Put each line into a separate row
            For i = LBound(splitText) To UBound(splitText)

                If i < rowCount Then
                    Cells(startRow + i, startCol).Value = splitText(i)
                End If

            Next i

        End If

    Next cell

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Undo merge completed!", vbInformation

End Sub
