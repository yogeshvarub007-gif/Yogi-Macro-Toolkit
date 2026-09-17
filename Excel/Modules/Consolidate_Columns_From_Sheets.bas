Attribute VB_Name = "Module_Consolidatecolumns"
Sub ConsolidateColumnsFromSheets_Fixed()
    Dim ws As Worksheet, destWs As Worksheet
    Dim colInput As String, rawArray() As String, colArray() As String
    Dim LastRow As Long, destRow As Long
    Dim i As Long, k As Long, s As Long
    Dim colLetter As String, headerText As String
    Dim firstHeaderFound As Boolean
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    ' Ask for columns (comma separated letters, allow spaces)
    colInput = InputBox("Enter column letters to consolidate (e.g. D,B,A):", "Select Columns")
    If Trim(colInput) = "" Then Exit Sub
    
    rawArray = Split(UCase(colInput), ",")
    ReDim colArray(LBound(rawArray) To UBound(rawArray))
    ' Trim spaces and validate simple letter format (A..ZZ)
    For i = LBound(rawArray) To UBound(rawArray)
        colArray(i) = Replace(Trim(rawArray(i)), "$", "")
        If colArray(i) = "" Then
            MsgBox "Invalid column input. Please retry.", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Create or clear Consolidated sheet
    On Error Resume Next
    Set destWs = wb.Sheets("Consolidated")
    If destWs Is Nothing Then
        Set destWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        destWs.Name = "Consolidated"
    Else
        destWs.Cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' -----------------------
    ' Build header row
    ' -----------------------
    destWs.Cells(1, 1).Value = "SheetName"
    
    For i = LBound(colArray) To UBound(colArray)
        colLetter = colArray(i)
        firstHeaderFound = False
        headerText = ""
        ' Search through sheets (in order) to find first non-empty header for this column
        For s = 1 To wb.Sheets.Count
            If wb.Sheets(s).Name <> destWs.Name Then
                On Error Resume Next
                headerText = Trim(CStr(wb.Sheets(s).Range(colLetter & "1").Value))
                On Error GoTo 0
                If headerText <> "" Then
                    firstHeaderFound = True
                    Exit For
                End If
            End If
        Next s
        If firstHeaderFound Then
            destWs.Cells(1, i + 2).Value = headerText
        Else
            destWs.Cells(1, i + 2).Value = "Column " & colLetter
        End If
    Next i
    
    ' Set header row height 25
    destWs.Rows(1).RowHeight = 25
    
    ' -----------------------
    ' Consolidate data
    ' -----------------------
    destRow = 2
    
    For Each ws In wb.Worksheets
        If ws.Name <> destWs.Name Then
            ' Skip completely empty sheets
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then GoTo NextSheet2
            
            ' Determine last relevant row for this sheet by checking requested columns
            LastRow = 0
            For k = LBound(colArray) To UBound(colArray)
                colLetter = colArray(k)
                On Error Resume Next
                Dim r As Long
                r = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
                On Error GoTo 0
                If r > LastRow Then LastRow = r
            Next k
            
            ' If only header or no data rows, skip
            If LastRow < 2 Then GoTo NextSheet2
            
            ' Insert blank row between sheets (but not before first sheet block)
            If destRow > 2 Then
                destWs.Rows(destRow).Insert Shift:=xlDown
                destRow = destRow + 1
            End If
            
            ' Copy data row by row for requested columns (preserve user column order)
            For i = 2 To LastRow
                ' Check if entire requested-col row is empty; if so skip that row
                Dim isRowEmpty As Boolean
                isRowEmpty = True
                For k = LBound(colArray) To UBound(colArray)
                    If Trim(CStr(ws.Cells(i, colArray(k)).Value)) <> "" Then
                        isRowEmpty = False
                        Exit For
                    End If
                Next k
                If isRowEmpty Then
                    ' still write SheetName? No - usually skip empty data rows
                    ' continue to next row
                Else
                    destWs.Cells(destRow, 1).Value = ws.Name
                    For k = LBound(colArray) To UBound(colArray)
                        colLetter = colArray(k)
                        On Error Resume Next
                        destWs.Cells(destRow, k + 2).Value = ws.Cells(i, colLetter).Value
                        On Error GoTo 0
                    Next k
                    destWs.Rows(destRow).RowHeight = 25
                    destRow = destRow + 1
                End If
            Next i
        End If
NextSheet2:
    Next ws
    
    ' Final formatting
    destWs.Columns.AutoFit
    ' Ensure row height 25 for any remaining rows (including header already set)
    If destWs.UsedRange.Rows.Count > 0 Then
        Dim rr As Range
        Set rr = destWs.UsedRange
        For Each rr In rr.Rows
            rr.RowHeight = 25
        Next rr
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Consolidation finished.", vbInformation
End Sub

