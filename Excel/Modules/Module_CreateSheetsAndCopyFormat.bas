Attribute VB_Name = "Module_Create_Sheets"
Option Explicit

Sub CreateSheetsAndCopyFormat()

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRowNames As Long
    Dim lastRowData As Long
    Dim lastCol As Long
    Dim i As Long
    Dim SheetName As String
    Dim formatRange As Range

    Set wb = ActiveWorkbook
    Set wsSource = wb.Sheets(1)

    ' Last row ONLY for names (Column A)
    lastRowNames = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Last row for ACTUAL DATA (any column)
    lastRowData = wsSource.UsedRange.Rows(wsSource.UsedRange.Rows.Count).Row

    ' Last column for data
    lastCol = wsSource.UsedRange.Columns(wsSource.UsedRange.Columns.Count).Column

    ' Define full format + data range (B onwards)
    Set formatRange = wsSource.Range(wsSource.Cells(1, 2), _
                                     wsSource.Cells(lastRowData, lastCol))

    Application.ScreenUpdating = False

    For i = 2 To lastRowNames

        SheetName = Trim(wsSource.Cells(i, 1).Value)

        If SheetName <> "" Then

            On Error Resume Next
            Set wsNew = wb.Sheets(SheetName)
            On Error GoTo 0

            If wsNew Is Nothing Then
                Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                wsNew.Name = SheetName
            Else
                ' ? Clear ONLY data, keep formats
                wsNew.Cells.ClearContents
            End If

            formatRange.Copy
            wsNew.Range("A1").PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            wsNew.Rows.RowHeight = 25
            wsNew.Columns.AutoFit

            Set wsNew = Nothing
        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox "Done! Formats preserved correctly.", vbInformation

End Sub


