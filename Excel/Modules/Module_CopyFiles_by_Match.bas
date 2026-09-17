Attribute VB_Name = "Module_CopyFiles_by_Match"
Option Explicit

Private Sub ScanFolderRecursive(ByVal folderPath As String, ByRef dict As Object)
    Dim fso As Object, folder As Object, subFolder As Object, file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        dict(file.Name) = file.Path
    Next file

    For Each subFolder In folder.SubFolders
        ScanFolderRecursive subFolder.Path, dict
    Next subFolder
End Sub
Sub CopyOrMoveFiles_WithSubfolders()

    Dim ws As Worksheet
    Dim srcFolder As String, destFolder As String
    Dim LastRow As Long, i As Long
    Dim baseName As String
    Dim colLetter As String, colNum As Long
    Dim statusCol As Long
    Dim dictFiles As Object
    Dim f As Variant
    Dim found As Boolean
    Dim actionText As String
    Dim processedCount As Long
    Dim confirmMove As VbMsgBoxResult

    Set ws = ActiveWorkbook.ActiveSheet
    Set dictFiles = CreateObject("Scripting.Dictionary")

    ' Ask column
    colLetter = InputBox("Enter column letter containing Base IDs (e.g. A):", "Select Column")
    If colLetter = "" Then Exit Sub
    colNum = Columns(colLetter).Column

    ' Copy / Move buttons
    frmCopyMove.SelectedAction = ""
    frmCopyMove.Show vbModal
    If frmCopyMove.SelectedAction = "" Then Exit Sub

    ' Confirm only for Move
    If frmCopyMove.SelectedAction = "MOVE" Then
        confirmMove = MsgBox( _
            "Files will be MOVED from all subfolders." & vbCrLf & _
            "Do you want to continue?", _
            vbYesNo + vbExclamation, _
            "Confirm Move")
        If confirmMove = vbNo Then Exit Sub
    End If

    actionText = IIf(frmCopyMove.SelectedAction = "COPY", "Copied", "Moved")

    ' Select folders
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select SOURCE folder"
        If .Show <> -1 Then Exit Sub
        srcFolder = .SelectedItems(1)
    End With

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select DESTINATION folder"
        If .Show <> -1 Then Exit Sub
        destFolder = .SelectedItems(1) & "\"
    End With

    ' Safe status column
    statusCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, statusCol).Value = "Status"

    ' Scan source folder + subfolders
    ScanFolderRecursive srcFolder, dictFiles

    LastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 2 To LastRow
        baseName = Trim(ws.Cells(i, colNum).Value)
        found = False
        processedCount = 0

        If baseName <> "" Then
            For Each f In dictFiles.Keys
                If Left(f, Len(baseName)) = baseName Then
                    If Dir(destFolder & f) = "" Then
                        If frmCopyMove.SelectedAction = "COPY" Then
                            FileCopy dictFiles(f), destFolder & f
                        Else
                            Name dictFiles(f) As destFolder & f
                        End If
                        processedCount = processedCount + 1
                    End If
                    found = True
                End If
            Next f

            ws.Cells(i, statusCol).Value = IIf(found, _
                actionText & " (" & processedCount & ")", _
                "No Match Found")
        End If
    Next i

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Completed successfully (subfolders included).", vbInformation

End Sub


