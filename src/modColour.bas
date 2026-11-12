Attribute VB_Name = "modColour"
Option Explicit
' Blue-black toggle, font and fill colour cycles, AutoColor.

' ---- blue-black toggle (Ctrl+;) -----------------------------------------------

' Active cell blue -> selection black, otherwise selection blue.
Public Sub ToggleBlueBlack()
    Dim areas As Collection, cur As Variant, target As Long
    On Error GoTo Fail
    ResetCycle
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub

    target = InputBlue()
    cur = ActiveCell.Font.Color
    If Not IsNull(cur) Then
        If CLng(cur) = InputBlue() Then target = FormulaBlack()
    End If
    ApplyFontColour areas, target, "Blue-black", IIf(target = InputBlue(), "Blue", "Black")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ShowStatus "Blue-black failed (" & Err.Description & ")"
End Sub

' ---- font colour cycle (Ctrl+') -------------------------------------------------

Public Sub CycleFontColour()
    Dim areas As Collection, cols As Variant, stepNo As Long, n As Long
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    cols = FontColourCycle()
    n = UBound(cols) - LBound(cols) + 1
    stepNo = NextStep("fontcolour", n)
    ApplyFontColour areas, CLng(cols(LBound(cols) + stepNo - 1)), "Font colour", _
        "Font colour " & stepNo & "/" & n
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Font colour failed (" & Err.Description & ")"
End Sub

' ---- fill colour cycle (Ctrl+Shift+K) -------------------------------------------

Public Sub CycleFillColour()
    Dim areas As Collection, cols As Variant, stepNo As Long, n As Long
    Dim colour As Long, area As Range
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    cols = FillColourCycle()
    n = UBound(cols) - LBound(cols) + 1
    stepNo = NextStep("fill", n)
    colour = CLng(cols(LBound(cols) + stepNo - 1))

    BeginSnapshot "Fill colour", CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_FILL
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        If colour = NO_FILL Then
            area.Interior.Pattern = xlNone
        Else
            area.Interior.Color = colour
        End If
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote("Fill colour " & stepNo & "/" & n)
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Fill colour failed (" & Err.Description & ")"
End Sub

' ---- AutoColor (Ctrl+Shift+A on Mac) ------------------------------------------

' Colours each cell by what it contains. Text, blanks and cells already in a
' colour outside the scheme (e.g. red flags) are left alone.
Public Sub AutoColorSelection()
    Dim areas As Collection, area As Range, c As Range
    Dim colour As Long, changed As Long
    On Error GoTo Fail
    ResetCycle
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub

    BeginSnapshot "AutoColor", CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_FONTCOLOR
    Next area

    Application.ScreenUpdating = False
    For Each area In areas
        For Each c In area.Cells
            If AutoColourFor(c, colour) Then
                If c.Font.Color <> colour Then c.Font.Color = colour
                changed = changed + 1
            End If
        Next c
    Next area
    Application.ScreenUpdating = True

    CommitSnapshot
    ShowStatus WithUndoNote("AutoColor: " & changed & " cells coloured")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ShowStatus "AutoColor failed (" & Err.Description & ")"
End Sub

' True (and colour set) if AutoColor should colour this cell.
Private Function AutoColourFor(ByVal c As Range, ByRef colour As Long) As Boolean
    Dim cur As Variant, f As String
    cur = c.Font.Color
    If IsNull(cur) Then Exit Function
    If Not InScheme(CLng(cur)) Then Exit Function

    If c.HasFormula Then
        f = CStr(c.Formula)
    Else
        f = SpillParentFormula(c)
    End If

    If f <> "" Then
        colour = FormulaColour(f, c.Worksheet.Name)
        AutoColourFor = True
    ElseIf VarType(c.Value2) = vbDouble Then
        colour = InputBlue()
        AutoColourFor = True
    End If
End Function

Private Function InScheme(ByVal colour As Long) As Boolean
    InScheme = (colour = InputBlue() Or colour = FormulaBlack() Or _
                colour = SheetLinkGreen() Or colour = BookLinkPurple())
End Function

' For a cell filled by a dynamic-array spill, the formula of the cell it spills from.
Private Function SpillParentFormula(ByVal c As Range) As String
    On Error Resume Next
    If c.HasSpill Then SpillParentFormula = CStr(c.SpillParent.Formula)
End Function

' Colour for a formula. First matching rule wins, in the spec's order:
' hardcode inside -> blue, other workbook -> purple, other sheet -> green, else black.
Public Function FormulaColour(ByVal f As String, ByVal ownSheet As String) As Long
    Dim hardcode As Boolean, sheetLink As Boolean, bookLink As Boolean
    ScanFormula f, ownSheet, hardcode, sheetLink, bookLink
    If hardcode Then
        FormulaColour = InputBlue()
    ElseIf bookLink Then
        FormulaColour = BookLinkPurple()
    ElseIf sheetLink Then
        FormulaColour = SheetLinkGreen()
    Else
        FormulaColour = FormulaBlack()
    End If
End Function

' ---- formula tokeniser ------------------------------------------------------------

' Reads an English (A1-style) formula once and reports whether it has a numeric
' hardcode, a link to another sheet, or a link to another workbook.
Private Sub ScanFormula(ByVal f As String, ByVal ownSheet As String, _
        ByRef hardcode As Boolean, ByRef sheetLink As Boolean, ByRef bookLink As Boolean)
    Dim i As Long, j As Long, n As Long, lvl As Long, depth As Long
    Dim ch As String, nx As String, tok As String
    Dim funcs() As String, args() As Long

    n = Len(f)
    ReDim funcs(0 To 16)
    ReDim args(0 To 16)
    args(0) = 1
    i = 1
    If Left$(f, 1) = "=" Then i = 2

    Do While i <= n
        ch = Mid$(f, i, 1)

        If ch = """" Then
            ' String literal: skip it ("" inside is an escaped quote).
            i = i + 1
            Do While i <= n
                If Mid$(f, i, 1) = """" Then
                    If Mid$(f, i + 1, 1) = """" Then
                        i = i + 2
                    Else
                        Exit Do
                    End If
                Else
                    i = i + 1
                End If
            Loop
            i = i + 1

        ElseIf ch = "'" Then
            ' Quoted sheet or workbook name: 'Sheet name'! or '...[Book.xlsx]Sheet'!
            j = i + 1
            Do While j <= n
                If Mid$(f, j, 1) = "'" Then
                    If Mid$(f, j + 1, 1) = "'" Then
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    j = j + 1
                End If
            Loop
            tok = Replace(Mid$(f, i + 1, j - i - 1), "''", "'")
            i = j + 1
            If Mid$(f, i, 1) = "!" Then
                NoteSheetRef tok, ownSheet, sheetLink, bookLink
                i = i + 1
            End If

        ElseIf ch = "[" Then
            ' [Book.xlsx]Sheet! is a workbook link; Table[Column] is not.
            j = i
            lvl = 0
            Do While j <= n
                If Mid$(f, j, 1) = "[" Then lvl = lvl + 1
                If Mid$(f, j, 1) = "]" Then
                    lvl = lvl - 1
                    If lvl = 0 Then Exit Do
                End If
                j = j + 1
            Loop
            i = j + 1
            j = i
            Do While j <= n
                If Not IsIdentChar(Mid$(f, j, 1)) Then Exit Do
                j = j + 1
            Loop
            If j > i And Mid$(f, j, 1) = "!" Then bookLink = True

        ElseIf IsIdentStart(ch) Then
            ' Name, function or cell reference (A1, $B$2, Sheet1, SUM ...).
            j = i
            Do While j <= n
                If Not IsIdentChar(Mid$(f, j, 1)) Then Exit Do
                j = j + 1
            Loop
            tok = Mid$(f, i, j - i)
            nx = Mid$(f, j, 1)
            If nx = "(" Then
                PushFunc funcs, args, depth, FunctionName(tok)
                j = j + 1
            ElseIf nx = "!" Then
                NoteSheetRef tok, ownSheet, sheetLink, bookLink
                j = j + 1
            End If
            i = j

        ElseIf IsDigit(ch) Or (ch = "." And IsDigit(Mid$(f, i + 1, 1))) Then
            ' Numeric literal.
            j = i
            Do While j <= n
                nx = Mid$(f, j, 1)
                If Not (IsDigit(nx) Or nx = ".") Then Exit Do
                j = j + 1
            Loop
            If UCase$(Mid$(f, j, 1)) = "E" Then
                nx = Mid$(f, j + 1, 1)
                If nx = "+" Or nx = "-" Or IsDigit(nx) Then
                    j = j + 2
                    Do While j <= n
                        If Not IsDigit(Mid$(f, j, 1)) Then Exit Do
                        j = j + 1
                    Loop
                End If
            End If
            tok = Mid$(f, i, j - i)
            If Not IsRowRange(f, i, j) Then
                If Val(tok) <> 0 And Val(tok) <> 1 Then
                    If Not IsPositionArg(funcs(depth), args(depth)) Then hardcode = True
                End If
            End If
            i = j

        Else
            Select Case ch
                Case "("
                    PushFunc funcs, args, depth, ""
                Case "{"
                    PushFunc funcs, args, depth, "{"
                Case ")", "}"
                    If depth > 0 Then depth = depth - 1
                Case ","
                    args(depth) = args(depth) + 1
            End Select
            i = i + 1
        End If
    Loop
End Sub

Private Sub PushFunc(ByRef funcs() As String, ByRef args() As Long, ByRef depth As Long, ByVal fn As String)
    depth = depth + 1
    If depth > UBound(funcs) Then
        ReDim Preserve funcs(0 To depth * 2)
        ReDim Preserve args(0 To depth * 2)
    End If
    funcs(depth) = fn
    args(depth) = 1
End Sub

' Drops the _xlfn. / _xlws. prefixes Excel adds to newer functions.
Private Function FunctionName(ByVal tok As String) As String
    tok = UCase$(tok)
    tok = Replace(tok, "_XLFN.", "")
    tok = Replace(tok, "_XLWS.", "")
    FunctionName = tok
End Function

Private Sub NoteSheetRef(ByVal sheetRef As String, ByVal ownSheet As String, _
        ByRef sheetLink As Boolean, ByRef bookLink As Boolean)
    If InStr(sheetRef, "[") > 0 Or InStr(sheetRef, "]") > 0 Or InStr(1, sheetRef, ".xl", vbTextCompare) > 0 Then
        bookLink = True
    ElseIf StrComp(sheetRef, ownSheet, vbTextCompare) <> 0 Then
        sheetLink = True
    End If
End Sub

' True for the 3 in 3:3 or 1:5 (whole-row references), not a hardcode.
Private Function IsRowRange(ByVal f As String, ByVal startPos As Long, ByVal endPos As Long) As Boolean
    Dim prevCh As String
    If startPos > 1 Then prevCh = Mid$(f, startPos - 1, 1)
    IsRowRange = (prevCh = ":" Or prevCh = "!" Or prevCh = "$" Or Mid$(f, endPos, 1) = ":")
End Function

Private Function IsPositionArg(ByVal fn As String, ByVal argNo As Long) As Boolean
    Dim spec As String, p As Long, e As Long
    If fn = "" Or fn = "{" Then Exit Function
    spec = "|" & POSITION_ARGS & "|"
    p = InStr(1, spec, "|" & fn & ":", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(fn) + 2
    e = InStr(p, spec, "|")
    IsPositionArg = InStr("," & Mid$(spec, p, e - p) & ",", "," & CStr(argNo) & ",") > 0
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Function IsIdentStart(ByVal ch As String) As Boolean
    IsIdentStart = (Len(ch) = 1) And ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
        Or ch = "_" Or ch = "$" Or ch = "\")
End Function

Private Function IsIdentChar(ByVal ch As String) As Boolean
    IsIdentChar = IsIdentStart(ch) Or IsDigit(ch) Or ch = "." Or ch = "?"
End Function

' ---- shared ----------------------------------------------------------------------

Private Sub ApplyFontColour(ByVal areas As Collection, ByVal colour As Long, _
                            ByVal actionLabel As String, ByVal msg As String)
    Dim area As Range
    BeginSnapshot actionLabel, CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_FONTCOLOR
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        area.Font.Color = colour
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote(msg)
End Sub
