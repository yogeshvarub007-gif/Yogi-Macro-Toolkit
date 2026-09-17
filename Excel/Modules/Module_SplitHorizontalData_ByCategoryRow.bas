Attribute VB_Name = "Module_rowdatasplit"
Sub SplitHorizontalData_ByCategoryRow()
    Dim ws As Worksheet
    Dim categoryRow As Long
    Dim headerRowCount As Long
    Dim lastCol As Long, LastRow As Long
    Dim i As Long, j As Long
    Dim category As String
    Dim newWS As Worksheet

    ' ====== Setup ======
    Set ws = ActiveSheet
    categoryRow = 3        ' <<<<< Change this to the row where your category is (e.g., 2 for Row 2)
    headerRowCount = 3     ' <<<<< Number of rows to copy as field names (e.g., 1 if Row 1 is Name)

    lastCol = ws.Cells(categoryRow, ws.Columns.Count).End(xlToLeft).Column
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    ' ====== Loop through each column ======
    For i = 1 To lastCol
        category = Trim(ws.Cells(categoryRow, i).Value)

        If category <> "" Then
            ' Create or get sheet with that category name
            On Error Resume Next
            Set newWS = Worksheets(category)
            If newWS Is Nothing Then
                Set newWS = Worksheets.Add(After:=Sheets(Sheets.Count))
                newWS.Name = category
            End If
            On Error GoTo 0

            ' Add headers if sheet is new
            If newWS.Cells(1, 1).Value = "" Then
                newWS.Cells(1, 1).Value = "Field"
                newWS.Cells(1, 2).Value = "Value"
            End If

            ' Copy all field names (row labels) and values in this column
            For j = 1 To LastRow
                Dim label As String
                Dim val As Variant
                label = ws.Cells(j, 1).Value    ' Field names in column A
                val = ws.Cells(j, i).Value

                If j <> categoryRow Then ' Skip the category row
                    Dim targetRow As Long
                    targetRow = newWS.Cells(newWS.Rows.Count, 1).End(xlUp).Row + 1
                    newWS.Cells(targetRow, 1).Value = label
                    newWS.Cells(targetRow, 2).Value = val
                End If
            Next j
        End If

        Set newWS = Nothing
    Next i

    Application.ScreenUpdating = True
    MsgBox "Data split by row category successfully!"
End Sub

