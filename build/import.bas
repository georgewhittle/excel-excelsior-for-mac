Attribute VB_Name = "modBuild"
Option Explicit
' ============================================================================
' Excel Excelsior For Mac builder. Lives in build/Builder.xlsm (not committed).
' BuildAddin reads src/*, builds ExcelExcelsiorForMac.xlam in Excel's Add-ins folder,
' then reloads it so the new version is live straight away.
' Needs: Tools > Macro > Security > "Trust access to the VBA project object model".
' ============================================================================

Private Const ADDIN_FILE As String = "ExcelExcelsiorForMac.xlam"
Private Const PROJECT_NAME As String = "ExcelExcelsiorForMac"
Private Const ADDIN_TITLE As String = "Excel Excelsior For Mac"

' Every file in src/, in import order. Add new modules here.
Private Function SourceFiles() As Variant
    SourceFiles = Array( _
        "modConfig.bas", "modUtil.bas", "modCycle.bas", "modUndo.bas", _
        "modNumbers.bas", "modColour.bas", "modBorders.bas", "modAlign.bas", _
        "modMain.bas", "clsAppEvents.cls", "ThisWorkbook.cls")
End Function

Public Sub BuildAddin()
    Dim srcDir As String, outPath As String, errMsg As String
    Dim files As Variant, f As Variant
    Dim wb As Workbook
    On Error GoTo Fail

    srcDir = RepoFolder() & "src" & Application.PathSeparator
    files = SourceFiles()

#If Mac Then
    ' Mac sandbox: Excel must be given access to the source files once.
    Dim grantList() As Variant, i As Long
    ReDim grantList(0 To UBound(files) - LBound(files))
    For i = LBound(files) To UBound(files)
        grantList(i - LBound(files)) = srcDir & files(i)
    Next i
    If Not GrantAccessToMultipleFiles(grantList) Then
        MsgBox "Excel needs access to the files in src/ to build.", vbExclamation, ADDIN_TITLE
        Exit Sub
    End If
#End If

    ' Unload the current copy so it can be overwritten.
    SafeUnloadAddin

    Set wb = Workbooks.Add
    For Each f In files
        ImportSource wb, srcDir & CStr(f)
    Next f
    wb.VBProject.Name = PROJECT_NAME
    SafeSetTitle wb

    outPath = AddinFolder() & ADDIN_FILE
    wb.IsAddin = True
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=outPath, FileFormat:=55   ' 55 = xlOpenXMLAddIn (.xlam)
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' Install (or reinstall) so Excel loads it now and at every startup.
    If SafeInstallAddin(outPath) Then
        MsgBox "Built and loaded:" & vbNewLine & outPath, vbInformation, ADDIN_TITLE
    Else
        MsgBox "Built:" & vbNewLine & outPath & vbNewLine & vbNewLine & _
            "Now tick it in Tools > Excel Add-ins (use Browse if it isn't listed).", _
            vbInformation, ADDIN_TITLE
    End If
    Exit Sub

Fail:
    errMsg = Err.Description
    Application.DisplayAlerts = True
    SafeCloseWorkbook wb
    MsgBox "Build failed: " & errMsg & vbNewLine & vbNewLine & _
        "Check that Tools > Macro > Security > 'Trust access to the VBA project " & _
        "object model' is ticked, and that Builder.xlsm is saved in the repo's build folder.", _
        vbExclamation, ADDIN_TITLE
End Sub

' ---- helpers -----------------------------------------------------------------

' The repo root, worked out from where this builder workbook is saved (build/).
Private Function RepoFolder() As String
    Dim p As String, sep As String
    sep = Application.PathSeparator
    p = ThisWorkbook.Path
    If p = "" Then Err.Raise vbObjectError + 1, , "Save this workbook as build/Builder.xlsm in the repo first."
    If LCase$(Left$(p, 4)) = "http" Then Err.Raise vbObjectError + 2, , "The repo is in a cloud folder. Keep it in a local folder."
    If LCase$(Mid$(p, InStrRev(p, sep) + 1)) <> "build" Then
        Err.Raise vbObjectError + 3, , "Builder.xlsm must be saved in the repo's build folder."
    End If
    RepoFolder = Left$(p, InStrRev(p, sep))
End Function

Private Function AddinFolder() As String
#If Mac Then
    ' HOME is the sandbox container path; strip it back to the real home folder.
    Dim home As String, pos As Long
    home = Environ("HOME")
    pos = InStr(home, "/Library/Containers/")
    If pos > 0 Then home = Left$(home, pos - 1)
    AddinFolder = home & "/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/"
#Else
    AddinFolder = Application.UserLibraryPath
#End If
End Function

' Adds one source file to wb. Code is added as text rather than through the
' VBA editor's importer, which is unreliable on Mac.
Private Sub ImportSource(ByVal wb As Workbook, ByVal path As String)
    Dim code As String, modName As String, ext As String
    Dim comp As Object, cm As Object
    code = CleanSource(ReadTextFile(path))
    modName = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    ext = LCase$(Mid$(modName, InStrRev(modName, ".") + 1))
    modName = Left$(modName, InStrRev(modName, ".") - 1)

    If modName = "ThisWorkbook" Then
        Set comp = WorkbookComponent(wb)
        If comp Is Nothing Then Err.Raise vbObjectError + 4, , "Couldn't find the ThisWorkbook module."
    ElseIf ext = "cls" Then
        Set comp = wb.VBProject.VBComponents.Add(2)   ' class module
        comp.Name = modName
    Else
        Set comp = wb.VBProject.VBComponents.Add(1)   ' standard module
        comp.Name = modName
    End If

    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString code
End Sub

Private Function WorkbookComponent(ByVal wb As Workbook) As Object
    On Error Resume Next
    Set WorkbookComponent = wb.VBProject.VBComponents("ThisWorkbook")
    If WorkbookComponent Is Nothing Then Set WorkbookComponent = wb.VBProject.VBComponents(wb.CodeName)
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim fnum As Integer, s As String
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    s = Space$(LOF(fnum))
    Get #fnum, , s
    Close #fnum
    ReadTextFile = s
End Function

' Normalises line endings and drops export headers and Attribute lines.
Private Function CleanSource(ByVal s As String) As String
    Dim lines As Variant, i As Long, ln As String, result As String
    Dim inHeader As Boolean
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    lines = Split(s, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If i = LBound(lines) And Left$(ln, 8) = "VERSION " Then inHeader = True
        If inHeader Then
            If Trim$(ln) = "END" Then inHeader = False
        ElseIf Left$(ln, 10) = "Attribute " Then
            ' skip
        Else
#If Mac Then
            result = result & ln & vbLf   ' Mac editor counts CR and LF as two line breaks
#Else
            result = result & ln & vbCrLf
#End If
        End If
    Next i
    CleanSource = result
End Function

Private Sub SafeUnloadAddin()
    On Error Resume Next
    Workbooks(ADDIN_FILE).Close SaveChanges:=False
    Err.Clear
End Sub

Private Function SafeInstallAddin(ByVal path As String) As Boolean
    Dim ai As Object
    On Error GoTo Done
    Set ai = Application.AddIns.Add(path, False)
    ai.Installed = False
    ai.Installed = True
    SafeInstallAddin = True
Done:
End Function

Private Sub SafeSetTitle(ByVal wb As Workbook)
    On Error Resume Next
    wb.BuiltinDocumentProperties("Title").Value = ADDIN_TITLE
    wb.BuiltinDocumentProperties("Comments").Value = "Keyboard shortcuts for formatting financial models in Excel for Mac."
End Sub

Private Sub SafeCloseWorkbook(ByVal wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub
