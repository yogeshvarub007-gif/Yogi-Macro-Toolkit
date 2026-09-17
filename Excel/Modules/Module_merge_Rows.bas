Attribute VB_Name = "Module_merge_Rows"
Sub MergeRowsWithOptions()
    Dim ws As Worksheet
    Dim rng As Range
    Dim firstRow As Long, LastRow As Long
    Dim col As Long, i As Long
    Dim mergedText As String
    Dim cell As Range
    Dim rowsToMerge As Variant
    Dim skipGap As VbMsgBoxResult
    
    Set ws = ActiveSheet
    If Selection Is Nothing Then Exit Sub
    Set rng = Selection
    
    ' Ask user how many rows to merge
    rowsToMerge = Application.InputBox("Enter number of rows to merge:", "Merge Rows", Type:=1)
    If rowsToMerge = False Then Exit Sub
    If rowsToMerge < 1 Then Exit Sub
    
    ' Ask whether to skip a gap row
    skipGap = MsgBox("Do you want to leave a blank row gap after each merged block?", vbYesNo + vbQuestion, "Merge Rows")
    
    firstRow = rng.Row
    LastRow = rng.Rows(rng.Rows.Count).Row
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    i = firstRow
    Do While i + rowsToMerge - 1 <= LastRow
        For col = rng.Column To rng.Column + rng.Columns.Count - 1
            mergedText = ""
            
            ' Collect displayed text from rows
            For Each cell In ws.Range(ws.Cells(i, col), ws.Cells(i + rowsToMerge - 1, col))
                If Trim(cell.Text) <> "" Then
                    mergedText = mergedText & cell.Text & Chr(10) ' same as Alt+Enter
                End If
            Next cell
            
            If Len(mergedText) > 0 Then
                mergedText = Left(mergedText, Len(mergedText) - 1) ' remove last LF
            End If
            
            ' Merge and put text safely
            With ws.Range(ws.Cells(i, col), ws.Cells(i + rowsToMerge - 1, col))
                .Merge
                .NumberFormat = "@"            ' force text format
                .Value = mergedText             ' assign combined string
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True                 ' ensure line breaks display correctly
            End With
        Next col
        
        ' Move pointer
        If skipGap = vbYes Then
            i = i + rowsToMerge + 1
        Else
            i = i + rowsToMerge
        End If
    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Merging completed!"
End Sub

