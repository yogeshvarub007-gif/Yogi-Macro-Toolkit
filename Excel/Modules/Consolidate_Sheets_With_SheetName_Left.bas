Attribute VB_Name = "Module_consol_with_Sheet_Name"
Sub Consolidate_Sheets_With_SheetName_Left()
    Dim wb As Workbook
    Dim ws As Worksheet, wsCons As Worksheet
    Dim hdrSet As Boolean
    Dim LastRow As Long, lastCol As Long
    Dim dataRows As Long
    Dim currentRow As Long
    Dim BlankRow As Long, dataStart As Long
    Dim i As Long
    Dim rowsToCopy As Long
    Dim userInput As String
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Prompt user for number of rows to consolidate
    userInput = InputBox("Enter the number of rows to consolidate from each sheet (excluding header):", "Rows to Consolidate")
    If userInput = "" Then
        MsgBox "Operation cancelled.", vbInformation
        GoTo Cleanup
    End If
    
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation
        GoTo Cleanup
    End If
    
    rowsToCopy = CLng(userInput)
    If rowsToCopy <= 0 Then
        MsgBox "Please enter a positive number.", vbExclamation
        GoTo Cleanup
    End If
    
    Set wb = ActiveWorkbook
    
    ' Delete existing "Consolidated" if present
    On Error Resume Next
    wb.Worksheets("Consolidated").Delete
    On Error GoTo 0
    
    ' Create new Consolidated sheet at the end
    Set wsCons = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsCons.Name = "Consolidated"
    
    ' --- 1) Build header row in Consolidated (SheetName + first sheet's headers) ---
    hdrSet = False
    For Each ws In wb.Worksheets
        If ws.Name <> wsCons.Name Then
            If Application.WorksheetFunction.CountA(ws.Rows(1)) > 0 Then
                ' Find last used column in row 1 (header)
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy Destination:=wsCons.Cells(1, 2)
                wsCons.Cells(1, 1).Value = "SheetName"
                hdrSet = True
                Exit For
            End If
        End If
    Next ws
    If Not hdrSet Then wsCons.Cells(1, 1).Value = "SheetName"
    
    ' --- 2) Loop sheets, add blank row then data (exclude each sheet's header row) ---
    currentRow = 1 ' header sits on row 1
    For Each ws In wb.Worksheets
        If ws.Name <> wsCons.Name Then
            ' Find last used row/col in this sheet (robust)
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
                LastRow = 0
                lastCol = 0
            Else
                LastRow = ws.Cells.Find(What:="*", _
                                        After:=ws.Cells(1, 1), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlPrevious).Row
                lastCol = ws.Cells.Find(What:="*", _
                                        After:=ws.Cells(1, 1), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlPrevious).Column
            End If
            
            ' Only copy if there's at least one data row after header
            If LastRow >= 2 Then
                ' Ensure consolidated header has names for any new columns from this sheet
                For i = 1 To lastCol
                    If Trim(wsCons.Cells(1, i + 1).Value) = "" Then
                        wsCons.Cells(1, i + 1).Value = ws.Cells(1, i).Value
                    End If
                Next i
                
                ' Determine number of rows to copy (user input or available rows, whichever is smaller)
                dataRows = Application.Min(rowsToCopy, LastRow - 1)
                If dataRows > 0 Then
                    BlankRow = currentRow + 1
                    dataStart = BlankRow + 1
                    
                    ' Copy sheet's data rows (row 2 to row 2+dataRows-1) into consolidated starting at column B
                    ws.Range(ws.Cells(2, 1), ws.Cells(2 + dataRows - 1, lastCol)).Copy Destination:=wsCons.Cells(dataStart, 2)
                    
                    ' Fill sheet name in column A for the pasted rows
                    wsCons.Range(wsCons.Cells(dataStart, 1), wsCons.Cells(dataStart + dataRows - 1, 1)).Value = ws.Name
                    
                    ' Debug: Log the number of rows copied from this sheet
                    Debug.Print "Sheet: " & ws.Name & ", Requested: " & rowsToCopy & ", Available: " & (LastRow - 1) & ", Copied: " & dataRows & " rows from rows 2 to " & (2 + dataRows - 1)
                    
                    ' Move currentRow to last pasted data row
                    currentRow = dataStart + dataRows - 1
                Else
                    Debug.Print "Sheet: " & ws.Name & ", Skipped: No data rows to copy"
                End If
            Else
                Debug.Print "Sheet: " & ws.Name & ", Skipped: No data rows after header"
            End If
        End If
    Next ws
    
    Application.CutCopyMode = False
    MsgBox "Consolidation complete. Copied up to " & rowsToCopy & " rows from each sheet.", vbInformation
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Cleanup
End Sub


