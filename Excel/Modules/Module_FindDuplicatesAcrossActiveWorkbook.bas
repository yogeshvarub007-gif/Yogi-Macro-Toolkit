Attribute VB_Name = "Module_FindDuplicates"
Option Explicit

Sub FindDuplicatesAcrossActiveWorkbook()

    Dim wb As Workbook
    Dim ws As Worksheet, wsResult As Worksheet
    Dim colInput As Variant, colNum As Long
    Dim LastRow As Long, lastCol As Long
    Dim lastRowResult As Long
    Dim dict As Object
    Dim dataArray As Variant, srcRow As Variant
    Dim r As Long, i As Long
    Dim cellVal As String, normalized As String
    Dim arrWords As Variant, sortedVal As String
    Dim tempList As Collection
    Dim key As Variant, items As Variant, item As Variant
    Dim parts As Variant
    Dim foundAny As Boolean
    Dim headerArray As Variant
    Dim headerWritten As Boolean

    On Error GoTo ErrorHandler
    Set wb = ActiveWorkbook

    ' Ask column
    colInput = Application.InputBox( _
        "Enter the column letter (e.g., B) or number (e.g., 2):", _
        "Column to Check", Type:=2)

    If colInput = False Or Trim(colInput) = "" Then Exit Sub

    If IsNumeric(colInput) Then
        colNum = CLng(colInput)
    Else
        colNum = Range(colInput & "1").Column
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Remove old Duplicates sheet
    On Error Resume Next
    wb.Worksheets("Duplicates").Delete
    On Error GoTo ErrorHandler

    ' Create result sheet
    Set wsResult = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsResult.Name = "Duplicates"

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    headerWritten = False

    ' ===== SCAN ALL SHEETS =====
    For Each ws In wb.Worksheets
        If ws.Name <> wsResult.Name Then

            LastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
            If LastRow < 2 Then GoTo NextSheet

            ' Capture headers ONCE (from first valid sheet)
            If Not headerWritten Then
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                headerArray = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value

                wsResult.Cells(1, 1).Value = "Sheet Name"
                wsResult.Cells(1, 2).Resize(1, lastCol).Value = headerArray

                headerWritten = True
            End If

            dataArray = ws.Range(ws.Cells(2, colNum), ws.Cells(LastRow, colNum)).Value

            For r = 1 To UBound(dataArray, 1)

                cellVal = CStr(dataArray(r, 1))
                normalized = NormalizeText(cellVal)

                If Len(normalized) > 0 Then
                    arrWords = Split(normalized, " ")
                    Set tempList = New Collection

                    For i = LBound(arrWords) To UBound(arrWords)
                        If arrWords(i) <> "" Then tempList.Add arrWords(i)
                    Next i

                    If tempList.Count > 0 Then
                        ReDim arrWords(0 To tempList.Count - 1)
                        For i = 1 To tempList.Count
                            arrWords(i - 1) = tempList(i)
                        Next i

                        QuickSort arrWords, LBound(arrWords), UBound(arrWords)
                        sortedVal = Join(arrWords, " ")

                        If dict.Exists(sortedVal) Then
                            dict(sortedVal) = dict(sortedVal) & "|" & ws.Name & "," & (r + 1)
                        Else
                            dict(sortedVal) = ws.Name & "," & (r + 1)
                        End If
                    End If
                End If
            Next r
        End If
NextSheet:
    Next ws

    ' ===== OUTPUT DUPLICATES =====
    lastRowResult = 2

    For Each key In dict.Keys
        items = Split(dict(key), "|")

        If UBound(items) >= 1 Then
            foundAny = True

            For Each item In items
                parts = Split(item, ",")

                wsResult.Cells(lastRowResult, 1).Value = parts(0)

                lastCol = wb.Worksheets(parts(0)).Cells(CLng(parts(1)), _
                           wb.Worksheets(parts(0)).Columns.Count).End(xlToLeft).Column

                srcRow = wb.Worksheets(parts(0)).Range( _
                         wb.Worksheets(parts(0)).Cells(CLng(parts(1)), 1), _
                         wb.Worksheets(parts(0)).Cells(CLng(parts(1)), lastCol)).Value

                wsResult.Cells(lastRowResult, 2).Resize(1, lastCol).Value = srcRow

                lastRowResult = lastRowResult + 1
            Next item

            lastRowResult = lastRowResult + 1 ' blank row between groups
        End If
    Next key

    If Not foundAny Then
        wsResult.Delete
        MsgBox "No duplicates found.", vbInformation
        GoTo Cleanup
    End If

    ' ===== FINAL FORMATTING (CORRECT RANGE) =====
    Dim finalLastRow As Long, finalLastCol As Long
    finalLastRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row
    finalLastCol = wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column

    With wsResult.Range(wsResult.Cells(1, 1), wsResult.Cells(finalLastRow, finalLastCol))
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Borders.LineStyle = xlContinuous
    End With

    wsResult.Rows(1).Font.Bold = True
    wsResult.Columns.AutoFit

    MsgBox "Duplicates sheet created successfully.", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

' ===== TEXT NORMALIZATION =====
Private Function NormalizeText(ByVal txt As String) As String
    Dim chars, titles, c, t

    txt = Replace(txt, Chr(160), " ")
    txt = Replace(txt, vbTab, " ")

    chars = Array(",", ".", ";", ":", "-", "_", "(", ")", "/", "\", "'", """", "&", "@", "#", "!")
    For Each c In chars
        txt = Replace(txt, c, " ")
    Next c

    txt = " " & LCase(txt) & " "
    titles = Array("prof", "professor", "dr", "mr", "ms", "mrs", "miss")
    For Each t In titles
        txt = Replace(txt, " " & t & " ", " ")
    Next t

    NormalizeText = Application.WorksheetFunction.Trim(txt)
End Function

' ===== QUICKSORT =====
Private Sub QuickSort(arr As Variant, first As Long, last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String

    If first >= last Then Exit Sub
    pivot = arr((first + last) \ 2)
    i = first: j = last

    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop

    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last
End Sub


