Attribute VB_Name = "Module_splitsheetstonNewbooks"
Sub SplitSheetsToNewWorkbooks()

    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim saveFolder As String
    Dim fileName As String
    Dim safeSheetName As String
    
    ' Ask user to choose folder to save output files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save Split Workbooks"
        If .Show <> -1 Then Exit Sub
        saveFolder = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each ws In ActiveWorkbook.Worksheets
        
        ' Fix invalid characters in file name
        safeSheetName = ws.Name
        safeSheetName = Replace(safeSheetName, "/", "-")
        safeSheetName = Replace(safeSheetName, "\", "-")
        safeSheetName = Replace(safeSheetName, ":", "-")
        safeSheetName = Replace(safeSheetName, "*", "-")
        safeSheetName = Replace(safeSheetName, "?", "")
        safeSheetName = Replace(safeSheetName, """", "'")
        safeSheetName = Replace(safeSheetName, "<", "(")
        safeSheetName = Replace(safeSheetName, ">", ")")
        safeSheetName = Replace(safeSheetName, "|", "-")
        
        ' Create a new workbook and copy sheet
        ws.Copy
        Set newWb = ActiveWorkbook
        
        ' Build file path
        fileName = saveFolder & "\" & safeSheetName & ".xlsx"
        
        ' Save workbook
        newWb.SaveAs fileName, FileFormat:=xlOpenXMLWorkbook
        
        ' Close workbook
        newWb.Close False
        
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Sheets successfully split into individual workbooks!", vbInformation

End Sub

