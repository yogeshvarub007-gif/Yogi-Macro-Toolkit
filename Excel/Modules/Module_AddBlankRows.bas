Attribute VB_Name = "Module_addblankrows"
Sub AddBlankRowsBetweenSelection()

    Dim rng As Range
    Dim i As Long
    Dim startRow As Long, endRow As Long
    Dim insertAfter As Long
    Dim blankRows As Long

    ' Ensure a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If

    ' Ask user inputs
    insertAfter = Application.InputBox( _
        Prompt:="Insert blank rows after how many rows?", _
        Title:="Row Interval", Type:=1)

    If insertAfter <= 0 Then Exit Sub

    blankRows = Application.InputBox( _
        Prompt:="How many blank rows to insert each time?", _
        Title:="Blank Rows Count", Type:=1)

    If blankRows <= 0 Then Exit Sub

    Set rng = Selection
    startRow = rng.Rows(1).Row
    endRow = rng.Rows(rng.Rows.Count).Row

    Application.ScreenUpdating = False

    ' Loop bottom to top to avoid shifting issues
    For i = endRow To startRow Step -1
        If (i - startRow + 1) Mod insertAfter = 0 Then
            Rows(i + 1).Resize(blankRows).Insert Shift:=xlDown
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Blank rows inserted successfully!", vbInformation

End Sub


