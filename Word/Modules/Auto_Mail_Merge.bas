Attribute VB_Name = "Auto_Mail_Merge"
Option Explicit

Sub MailMerge_Save_FullOptimized()

    Dim mainDoc As Document
    Dim mm As MailMerge
    Dim recCount As Long, i As Long

    Dim FileName As String
    Dim Folder As String
    Dim savePath As String
    Dim MainFolder As String
    Dim FolderFolder As String

    Dim NewDoc As Document
    Dim fDialog As FileDialog

    Dim nameCol As String
    Dim folderCol As String
    Dim nameIndex As Long
    Dim folderIndex As Long

    Dim useFolders As VbMsgBoxResult
    Dim saveMode As Long
    Dim modeInput As String

    Dim logFile As String
    Dim logNum As Integer

    On Error GoTo ErrHandler

    ' -------------------------
    ' Silent mode
    ' -------------------------
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.Visible = False

    ' -------------------------
    ' Validate Mail Merge doc
    ' -------------------------
    If Documents.Count = 0 Then
        MsgBox "Open Mail Merge main document first.", vbExclamation
        GoTo SafeExit
    End If

    Set mainDoc = ActiveDocument

    If mainDoc.MailMerge Is Nothing Then
        MsgBox "This is not a Mail Merge document.", vbExclamation
        GoTo SafeExit
    End If

    Set mm = mainDoc.MailMerge

    If mm.State <> wdMainAndDataSource Then
        MsgBox "Open the ORIGINAL merge template.", vbExclamation
        GoTo SafeExit
    End If

    recCount = mm.DataSource.RecordCount

    If recCount < 1 Then
        MsgBox "No records found.", vbExclamation
        GoTo SafeExit
    End If

    ' -------------------------
    ' Filename column
    ' -------------------------
    nameCol = InputBox( _
        "Enter Excel column LETTER for file names" & vbCrLf & _
        "Example: A", _
        "Filename Column", "A")

    If Trim(nameCol) = "" Then GoTo SafeExit

    nameIndex = ColumnLetterToNumber(nameCol)

    ' -------------------------
    ' Folder option
    ' -------------------------
    useFolders = MsgBox( _
        "Create folders using another column?", _
        vbYesNo + vbQuestion)

    If useFolders = vbYes Then

        folderCol = InputBox( _
            "Enter Excel column LETTER for folders", _
            "Folder Column", "B")

        If Trim(folderCol) = "" Then GoTo SafeExit

        folderIndex = ColumnLetterToNumber(folderCol)

    End If

    ' -------------------------
    ' Save mode
    ' -------------------------
    modeInput = InputBox( _
        "Choose output:" & vbCrLf & vbCrLf & _
        "1 = DOCX only" & vbCrLf & _
        "2 = PDF only" & vbCrLf & _
        "3 = DOCX + PDF", _
        "Save Mode", _
        "1")

    If modeInput = "" Then GoTo SafeExit

    If modeInput <> "1" _
    And modeInput <> "2" _
    And modeInput <> "3" Then

        MsgBox "Enter 1, 2 or 3.", vbExclamation
        GoTo SafeExit

    End If

    saveMode = CLng(modeInput)

    ' -------------------------
    ' Output folder
    ' -------------------------
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    fDialog.Title = "Select output folder"

    If fDialog.Show <> -1 Then GoTo SafeExit

    MainFolder = fDialog.SelectedItems(1)

    If Right(MainFolder, 1) <> "\" Then
        MainFolder = MainFolder & "\"
    End If

    ' -------------------------
    ' Logging
    ' -------------------------
    logFile = MainFolder & "merge_log.txt"

    logNum = FreeFile

    Open logFile For Output As #logNum

    Print #logNum, "Merge Run: " & Now
    Print #logNum, "Filename Column: " & UCase(nameCol)

    If useFolders = vbYes Then
        Print #logNum, "Folder Column: " & UCase(folderCol)
    End If

    Print #logNum, _
        "Save Mode: " & _
        Choose(saveMode, _
        "DOCX Only", _
        "PDF Only", _
        "DOCX + PDF")

    Print #logNum, String(40, "-")

    ' =========================
    ' PROCESS RECORDS
    ' =========================

    For i = 1 To recCount

        mm.DataSource.ActiveRecord = i

        FileName = GetMergeFieldByIndex(mm, nameIndex)
        FileName = SanitizeFileName(FileName)

        ' Skip blanks
        If Trim(FileName) = "" _
        Or FileName = "Unnamed" Then

            Print #logNum, _
            "Record " & i & _
            " skipped (blank filename)"

            GoTo ContinueLoop

        End If

        ' Folder
        If useFolders = vbYes Then

            Folder = GetMergeFieldByIndex(mm, folderIndex)
            Folder = SanitizeFileName(Folder)

        Else

            Folder = ""

        End If

        ' Build path
        If useFolders = vbYes _
        And Folder <> "" Then

            FolderFolder = MainFolder & Folder & "\"

            If Dir(FolderFolder, vbDirectory) = "" Then
                MkDir FolderFolder
            End If

            savePath = FolderFolder & FileName

        Else

            savePath = MainFolder & FileName

        End If

        savePath = EnsureUniquePath(savePath)

        ' Merge one record
        mm.DataSource.FirstRecord = i
        mm.DataSource.LastRecord = i
        mm.Destination = wdSendToNewDocument
        mm.Execute Pause:=False

        Set NewDoc = ActiveDocument

        ' SAFE cleanup
        RemoveTrailingBreaksSafe NewDoc

        ' Header shape handling
        LockHeaderShapes NewDoc

        NewDoc.Repaginate
        DoEvents

        ' Save
        Select Case saveMode

            Case 1
                NewDoc.SaveAs2 _
                    savePath & ".docx", _
                    wdFormatXMLDocument

            Case 2
                NewDoc.ExportAsFixedFormat _
                    savePath & ".pdf", _
                    wdExportFormatPDF

            Case 3
                NewDoc.SaveAs2 _
                    savePath & ".docx", _
                    wdFormatXMLDocument

                NewDoc.ExportAsFixedFormat _
                    savePath & ".pdf", _
                    wdExportFormatPDF

        End Select

        NewDoc.Close False

        Print #logNum, _
            "Record " & i & _
            " saved -> " & savePath

ContinueLoop:
    Next i

    Close #logNum

    MsgBox "Completed successfully.", vbInformation

SafeExit:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.Visible = True

    Exit Sub

ErrHandler:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.Visible = True

    If logNum <> 0 Then Close #logNum

    MsgBox _
        "Error " & Err.Number & vbCrLf & Err.Description, _
        vbCritical

End Sub


' ==========================================
' SAFE BREAK CLEANUP
' ==========================================
Private Sub RemoveTrailingBreaksSafe(d As Document)

    Dim rng As Range

    Set rng = d.Content
    rng.Collapse wdCollapseEnd

    rng.MoveStart wdCharacter, -1

    Select Case AscW(rng.Text)

        Case 11, 12
            rng.Delete

    End Select

End Sub


' ==========================================
' HEADER SHAPE FIX FOR PDF
' ==========================================
Private Sub LockHeaderShapes(d As Document)

    Dim shp As Shape

    For Each shp In d.Shapes

        On Error Resume Next

        If shp.Anchor.StoryType = wdPrimaryHeaderStory Then

            shp.LockAnchor = True

            shp.Line.Visible = msoTrue
            shp.Line.Transparency = 0

            If shp.Line.Weight < 0.5 Then
                shp.Line.Weight = 0.5
            End If

        End If

        On Error GoTo 0

    Next shp

End Sub


' ==========================================
Private Function GetMergeFieldByIndex( _
    mm As MailMerge, _
    idx As Long) As String

    On Error Resume Next

    GetMergeFieldByIndex = _
    Trim(mm.DataSource.DataFields(idx).Value & "")

End Function


Private Function ColumnLetterToNumber( _
    col As String) As Long

    Dim i As Long

    col = UCase(col)

    For i = 1 To Len(col)

        ColumnLetterToNumber = _
        ColumnLetterToNumber * 26 + _
        (Asc(Mid(col, i, 1)) - 64)

    Next i

End Function


Private Function SanitizeFileName( _
    s As String) As String

    Dim bad
    Dim ch

    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each ch In bad
        s = Replace(s, ch, "_")
    Next ch

    If Trim(s) = "" Then s = "Unnamed"

    SanitizeFileName = Trim(s)

End Function


Private Function EnsureUniquePath( _
    p As String) As String

    Dim base As String
    Dim ext As String
    Dim idx As Long
    Dim pos As Long

    pos = InStrRev(p, ".")

    If pos > 0 Then
        base = Left(p, pos - 1)
        ext = Mid(p, pos)
    Else
        base = p
        ext = ""
    End If

    EnsureUniquePath = p

    idx = 1

    Do While Dir(EnsureUniquePath) <> ""

        EnsureUniquePath = _
        base & "(" & idx & ")" & ext

        idx = idx + 1

    Loop

End Function

