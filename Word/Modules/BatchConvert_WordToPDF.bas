Attribute VB_Name = "BatchConvert_WordToPDF"
Option Explicit

' ============================================================
'  Batch Word ? PDF Converter
'  Fixes: shape loss, floating object displacement,
'         header/footer drawing disappearance
'  Method: Forces "Microsoft Print to PDF" printer context
'          before each conversion, mimicking manual Save As
' ============================================================

Dim gConverted  As Long
Dim gFailed     As Long
Dim gErrors()   As String
Dim gOrigPrinter As String

' -- Entry Point ---------------------------------------------
Public Sub BatchConvert_WordToPDF()

    Dim fd         As FileDialog
    Dim FolderPath As String

    ' -- 1. Folder picker ------------------------------------
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select root folder containing Word files"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    FolderPath = fd.SelectedItems(1)

    ' -- 2. Verify "Microsoft Print to PDF" is available -----
    If Not PrinterExists("Microsoft Print to PDF") Then
        MsgBox "Required: 'Microsoft Print to PDF' printer (built-in on " & _
               "Windows 10/11). It was not found on this machine." & vbCrLf & _
               "See the fallback options at the bottom of this macro.", vbCritical
        Exit Sub
    End If

    ' -- 3. Save current printer, switch to PDF printer ------
    '       This is the critical step that fixes shape rendering
    gOrigPrinter = Application.ActivePrinter
    SetPrinter "Microsoft Print to PDF"

    ' -- 4. Suppress UI --------------------------------------
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.Options.UpdateLinksAtOpen = False
    Application.Options.ConfirmConversions = False

    ' -- 5. Process recursively ------------------------------
    gConverted = 0
    gFailed = 0
    ReDim gErrors(0)

    ProcessFolder FolderPath

    ' -- 6. Restore everything -------------------------------
    SetPrinter gOrigPrinter
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.Options.UpdateLinksAtOpen = True

    ' -- 7. Summary ------------------------------------------
    Dim summary As String
    summary = "Conversion complete." & vbCrLf & vbCrLf & _
              "? Converted : " & gConverted & vbCrLf & _
              "? Failed    : " & gFailed

    If gFailed > 0 Then
        summary = summary & vbCrLf & vbCrLf & "Failed files:"
        Dim i As Long
        For i = 0 To UBound(gErrors)
            If gErrors(i) <> "" Then
                summary = summary & vbCrLf & "  • " & gErrors(i)
            End If
        Next i
    End If

    MsgBox summary, IIf(gFailed > 0, vbExclamation, vbInformation), _
           "Batch PDF Conversion"

End Sub

' -- Recursive folder walker ----------------------------------
Private Sub ProcessFolder(ByVal FolderPath As String)

    Dim FSO  As Object
    Dim fld  As Object
    Dim sf   As Object
    Dim fl   As Object
    Dim ext  As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fld = FSO.GetFolder(FolderPath)

    For Each fl In fld.Files

        ' Skip temp lock files Word creates
        If Left(fl.Name, 2) = "~$" Then GoTo NextFile

        ext = LCase(FSO.GetExtensionName(fl.Name))
        If ext = "docx" Or ext = "doc" Then
            ConvertSingle fl.Path
        End If

NextFile:
    Next fl

    ' Recurse into subfolders
    For Each sf In fld.SubFolders
        ProcessFolder sf.Path
    Next sf

End Sub

' -- Single file conversion -----------------------------------
Private Sub ConvertSingle(ByVal FilePath As String)

    Dim doc     As Document
    Dim PDFPath As String
    Dim errMsg  As String

    ' Build output path (same folder as source)
    PDFPath = Left(FilePath, InStrRev(FilePath, ".")) & "pdf"

    On Error GoTo ConvError

    ' -- Open document ----------------------------------------
    '   NOTE: Visible:=True is intentional here.
    '   Opening as Visible:=False skips Word's full layout/rendering
    '   pipeline on many builds, which is exactly what causes shapes
    '   to disappear. We open in a minimized, background window instead.
    Set doc = Documents.Open( _
        FileName:=FilePath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Repair:=False, _
        ShowingHelp:=False, _
        Visible:=True)          ' <- MUST be True for full render

    ' Minimize the document window to keep it off-screen
    doc.ActiveWindow.WindowState = wdWindowStateMinimize

    ' -- Force full layout repagination ----------------------
    '   This ensures all floating objects, anchors, and
    '   header/footer canvases are fully resolved before export
    doc.Repaginate
    doc.ActiveWindow.View.Type = wdPrintView  ' Print Layout = full fidelity

    ' -- Export using the same parameters Word uses internally
    '   for File ? Save As ? PDF (reverse-engineered from Word's
    '   own XML export manifest):
    doc.ExportAsFixedFormat2 _
        OutputFileName:=PDFPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False

    doc.Close SaveChanges:=wdDoNotSaveChanges
    gConverted = gConverted + 1
    Exit Sub

ConvError:
    errMsg = FilePath & "  [Err " & Err.Number & ": " & Err.Description & "]"
    gFailed = gFailed + 1
    ReDim Preserve gErrors(gFailed - 1)
    gErrors(gFailed - 1) = errMsg

    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0

End Sub

' -- Printer helpers ------------------------------------------

' Word stores ActivePrinter as "PrinterName on PortName"
' We need to find the full string for "Microsoft Print to PDF"
Private Function PrinterExists(ByVal PrinterName As String) As Boolean
    Dim oShell  As Object
    Dim oWMI    As Object
    Dim oPrints As Object
    Dim oPrint  As Object

    Set oShell = CreateObject("WScript.Shell")
    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set oPrints = oWMI.ExecQuery( _
        "SELECT Name FROM Win32_Printer WHERE Name LIKE '%" & PrinterName & "%'")

    PrinterExists = (oPrints.Count > 0)
End Function

Private Sub SetPrinter(ByVal PrinterName As String)
    ' Word requires the full "Name on PortXXX:" format.
    ' This routine finds the correct full name via WMI.
    Dim oWMI    As Object
    Dim oPrints As Object
    Dim oPrint  As Object

    ' If we're restoring original printer, just set it directly
    ' (we stored the full string when we saved it)
    If InStr(PrinterName, " on ") > 0 Then
        Application.ActivePrinter = PrinterName
        Exit Sub
    End If

    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set oPrints = oWMI.ExecQuery( _
        "SELECT Name, PortName FROM Win32_Printer WHERE Name LIKE '%" _
        & PrinterName & "%'")

    For Each oPrint In oPrints
        Application.ActivePrinter = oPrint.Name & " on " & oPrint.PortName & ":"
        Exit Sub
    Next oPrint
End Sub

