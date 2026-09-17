Attribute VB_Name = "Module_CopySheetsFromFolder"
Option Explicit

Sub CopySheetsFromFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim destWb As Workbook
    Dim newSheetName As String
    Dim counter As Long
    
    ' Destination is the active workbook
    Set destWb = ActiveWorkbook
    
    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Workbooks"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & "\"
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    fileName = Dir(folderPath & "*.xls*") ' picks .xlsx, .xlsm, .xls
    
    Do While fileName <> ""
        
        ' Skip the destination workbook to avoid copying into itself
        If folderPath & fileName <> destWb.FullName Then
            
            Set wbSource = Workbooks.Open(folderPath & fileName)
            
            For Each ws In wbSource.Worksheets
                
                ' Handle duplicate sheet names
                newSheetName = ws.Name
                counter = 1
                While SheetExists(newSheetName, destWb)
                    newSheetName = ws.Name & "_" & counter
                    counter = counter + 1
                Wend
                
                ws.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                destWb.Sheets(destWb.Sheets.Count).Name = newSheetName
            Next ws
            
            wbSource.Close SaveChanges:=False
        End If
        
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "All sheets copied successfully!", vbInformation
End Sub

Function SheetExists(SheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(SheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

