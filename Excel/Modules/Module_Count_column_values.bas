Attribute VB_Name = "Module_Count_column_values"
Sub CountColumnValues_ActiveWorkbook()
    Dim ws As Worksheet
    Dim indexWs As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim countVals As Long
    Dim colInput As String
    Dim colNum As Long
    Dim awb As Workbook
    Dim cell As Range
    Dim displayOption As String
    Dim headerValue As String
    
    ' Work on the active workbook
    Set awb = ActiveWorkbook
    
    ' Ask user which column to check (A, B, C etc.)
    colInput = InputBox("Enter the column letter to count values (e.g., A, B, C):", "Select Column")
    If colInput = "" Then Exit Sub ' Cancel pressed
    colInput = UCase(colInput)
    colNum = Range(colInput & "1").Column
    
    ' Ask user what to display: sheet names, column headers, or both
    displayOption = InputBox("What to display in SheetIndex?" & vbCrLf & _
                            "1 = Sheet Names" & vbCrLf & _
                            "2 = Column Headers" & vbCrLf & _
                            "3 = Both Sheet Names and Column Headers", _
                            "Display Option", "1")
    If displayOption = "" Then Exit Sub ' Cancel pressed
    If Not (displayOption = "1" Or displayOption = "2" Or displayOption = "3") Then
        MsgBox "Invalid option. Please enter 1, 2, or 3.", vbExclamation
        Exit Sub
    End If
    
    ' Check if "SheetIndex" exists, if yes clear it, if not create it
    On Error Resume Next
    Set indexWs = awb.Worksheets("SheetIndex")
    On Error GoTo 0
    
    If indexWs Is Nothing Then
        Set indexWs = awb.Worksheets.Add
        indexWs.Name = "SheetIndex"
    Else
        indexWs.Cells.Clear
    End If
    
    ' Always move SheetIndex to the first position
    indexWs.Move Before:=awb.Sheets(1)
    
    ' Set headers based on display option
    If displayOption = "1" Then
        indexWs.Range("A1").Value = "Name"
        indexWs.Range("B1").Value = "Count"
    ElseIf displayOption = "2" Then
        indexWs.Range("A1").Value = "Column Header"
        indexWs.Range("B1").Value = "Count"
    Else ' displayOption = "3"
        indexWs.Range("A1").Value = "Name"
        indexWs.Range("B1").Value = "Column Header"
        indexWs.Range("C1").Value = "Count"
    End If
    
    i = 2 ' start row for summary
    
    ' Loop through all sheets except SheetIndex
    For Each ws In awb.Worksheets
        If ws.Name <> "SheetIndex" Then
            ' Find last row in selected column
            LastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
            
            ' Get header value (row 1 of selected column)
            headerValue = ws.Cells(1, colNum).Value
            If headerValue = "" Then headerValue = "No Header"
            
            ' Reset counter
            countVals = 0
            
            ' Only count if data rows exist (row 2 and below)
            If LastRow > 1 Then
                For Each cell In ws.Range(ws.Cells(2, colNum), ws.Cells(LastRow, colNum))
                    If Trim(cell.Value) <> "" Then
                        countVals = countVals + 1
                    End If
                Next cell
            End If
            
            ' Write results in index sheet based on display option
            If displayOption = "1" Then
                indexWs.Cells(i, 1).Value = ws.Name
                indexWs.Cells(i, 2).Value = countVals
            ElseIf displayOption = "2" Then
                indexWs.Cells(i, 1).Value = headerValue
                indexWs.Cells(i, 2).Value = countVals
            Else ' displayOption = "3"
                indexWs.Cells(i, 1).Value = ws.Name
                indexWs.Cells(i, 2).Value = headerValue
                indexWs.Cells(i, 3).Value = countVals
            End If
            
            i = i + 1
        End If
    Next ws
    
    ' Autofit columns
    If displayOption = "3" Then
        indexWs.Columns("A:C").AutoFit
    Else
        indexWs.Columns("A:B").AutoFit
    End If
End Sub

