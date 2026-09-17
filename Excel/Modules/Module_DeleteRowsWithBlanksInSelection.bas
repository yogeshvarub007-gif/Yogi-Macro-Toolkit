Attribute VB_Name = "Module_removerow"
Sub DeleteRowsWithBlanksInSelection()

    Dim rng As Range
    Dim cell As Range
    Dim delRange As Range
    
    'Work with the current selection
    Set rng = Selection
    
    'Check each cell in selection
    For Each cell In rng
        If IsEmpty(cell.Value) Or Trim(cell.Value) = "" Then
            'Mark the entire row for deletion
            If delRange Is Nothing Then
                Set delRange = cell.EntireRow
            Else
                Set delRange = Union(delRange, cell.EntireRow)
            End If
        End If
    Next cell
    
    'Delete rows if any found
    If Not delRange Is Nothing Then
        delRange.Delete
    End If
    
End Sub

