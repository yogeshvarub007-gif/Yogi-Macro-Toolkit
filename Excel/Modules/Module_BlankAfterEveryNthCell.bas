Attribute VB_Name = "Module_blankafterEveryNthCell"
Sub InsertBlankCellsAfterEveryNthCell()
    Dim rng As Range
    Dim cellCount As Long
    Dim insertEvery As Long
    Dim insertCount As Long
    Dim i As Long
    Dim isColumn As Boolean
    Dim cellIndex As Long

    ' Get selected range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    Set rng = Selection

    ' Determine if range is vertical or horizontal
    If rng.Columns.Count = 1 Then
        isColumn = True
        cellCount = rng.Rows.Count
    ElseIf rng.Rows.Count = 1 Then
        isColumn = False
        cellCount = rng.Columns.Count
    Else
        MsgBox "Please select a single row or column.", vbExclamation
        Exit Sub
    End If

    ' Get user input
    insertCount = Application.InputBox("Enter number of blank cells to insert:", Type:=1)
    If insertCount < 1 Then Exit Sub

    insertEvery = Application.InputBox("Insert after how many cells?", Type:=1)
    If insertEvery < 1 Then Exit Sub

    ' Loop from bottom to top (or right to left) to prevent shifting problems
    For i = cellCount To 1 Step -1
        If i Mod insertEvery = 0 Then
            If isColumn Then
                rng.Cells(i + 1, 1).Resize(insertCount, 1).Insert Shift:=xlDown
            Else
                rng.Cells(1, i + 1).Resize(1, insertCount).Insert Shift:=xlToRight
            End If
        End If
    Next i

    MsgBox "Blank cells inserted successfully.", vbInformation
End Sub

