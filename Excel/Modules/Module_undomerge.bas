Attribute VB_Name = "Module_undomerge"
Sub UndoMergeRows()
    Dim rng As Range, cell As Range
    Dim splitText As Variant
    Dim i As Long
    
    If Selection Is Nothing Then Exit Sub
    Set rng = Selection
    
    Application.ScreenUpdating = False
    
    For Each cell In rng
        If cell.MergeCells Then
            With cell.MergeArea
                ' Split merged text by line breaks
                splitText = Split(cell.Value, vbNewLine)
                
                ' Unmerge first
                .UnMerge
                
                ' Put split text back into rows
                For i = LBound(splitText) To UBound(splitText)
                    If i + .Row <= .Row + .Rows.Count - 1 Then
                        cell.Offset(i, 0).Value = splitText(i)
                    End If
                Next i
            End With
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox "Undo merge completed!"
End Sub

