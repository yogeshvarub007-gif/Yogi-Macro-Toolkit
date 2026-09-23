Attribute VB_Name = "FindReplace_AllDocs"
Option Explicit

Sub FastSafe_FindReplace_AllDocs()

    Dim fd As FileDialog
    Dim RootFolder As String
    Dim FindText As String
    Dim ReplaceText As String

    FindText = InputBox("Enter text to FIND")
    If Trim(FindText) = "" Then Exit Sub

    ReplaceText = InputBox("Enter text to REPLACE WITH")

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select Main Folder"
        If .Show <> -1 Then Exit Sub
        RootFolder = .SelectedItems(1)
    End With


    'Performance boost
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    Options.CheckSpellingAsYouType = False
    Options.CheckGrammarAsYouType = False
    Options.BackgroundSave = False
    Options.Pagination = False


    ProcessFolder RootFolder, FindText, ReplaceText


    'Restore settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll

    Options.CheckSpellingAsYouType = True
    Options.CheckGrammarAsYouType = True
    Options.BackgroundSave = True
    Options.Pagination = True

    MsgBox "Find & Replace Completed", vbInformation

End Sub



Sub ProcessFolder(ByVal FolderPath As String, _
                  ByVal FindText As String, _
                  ByVal ReplaceText As String)

    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim doc As Document

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(FolderPath)


    For Each File In Folder.Files

        Select Case LCase(FSO.GetExtensionName(File.Name))

            Case "doc", "docx", "docm"

                On Error Resume Next

                Set doc = Documents.Open( _
                    FileName:=File.Path, _
                    Visible:=False, _
                    AddToRecentFiles:=False, _
                    ReadOnly:=False, _
                    ConfirmConversions:=False)

                If Err.Number = 0 Then

                    ReplaceEverywhere _
                    doc, FindText, ReplaceText


                    If Not doc.Saved Then
                        doc.Save
                    End If

                    doc.Close SaveChanges:=False

                End If

                Err.Clear
                On Error GoTo 0

        End Select

    Next File


    'Subfolders
    For Each SubFolder In Folder.SubFolders

        ProcessFolder _
        SubFolder.Path, _
        FindText, _
        ReplaceText

    Next SubFolder

End Sub




Sub ReplaceEverywhere(ByVal doc As Document, _
                      ByVal FindText As String, _
                      ByVal ReplaceText As String)

    Dim rng As Range
    Dim shp As Shape


    'Main stories + tables + headers + footers
    For Each rng In doc.StoryRanges

        Do

            With rng.Find

                .ClearFormatting
                .Replacement.ClearFormatting

                .Text = FindText
                .Replacement.Text = ReplaceText

                .Forward = True
                .Wrap = wdFindContinue

                .MatchCase = False
                .MatchWholeWord = True     'SAFER
                .MatchWildcards = False

                .Execute Replace:=wdReplaceAll

            End With


            Set rng = rng.NextStoryRange

        Loop Until rng Is Nothing

    Next rng



    'Text boxes / shapes
    For Each shp In doc.Shapes

        On Error Resume Next

        If shp.TextFrame.HasText Then

            shp.TextFrame.TextRange.Find.Execute _
                FindText:=FindText, _
                ReplaceWith:=ReplaceText, _
                MatchWholeWord:=True, _
                Replace:=wdReplaceAll

        End If

        On Error GoTo 0

    Next shp

End Sub

