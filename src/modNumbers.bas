Attribute VB_Name = "modNumbers"
Option Explicit
' Number, percent, multiple and date cycles; increase / decrease decimals.

Public Sub CycleNumber()
    RunFormatCycle "number", "Number format", NumberCycle()
End Sub

Public Sub CyclePercent()
    RunFormatCycle "percent", "Percent format", PercentCycle()
End Sub

Public Sub CycleMultiple()
    RunFormatCycle "multiple", "Multiple format", MultipleCycle()
End Sub

Public Sub CycleDate()
    RunFormatCycle "date", "Date format", DateCycle()
End Sub

Public Sub IncreaseDecimals()
    ShiftSelectionDecimals 1, "Increase decimals"
End Sub

Public Sub DecreaseDecimals()
    ShiftSelectionDecimals -1, "Decrease decimals"
End Sub

' ---- cycles --------------------------------------------------------------------

' Applies the next format in a cycle to the selection, with undo.
Private Sub RunFormatCycle(ByVal cycleId As String, ByVal actionLabel As String, ByVal formats As Variant)
    Dim areas As Collection, stepNo As Long, n As Long, fmt As String
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub

    n = UBound(formats) - LBound(formats) + 1
    stepNo = NextStep(cycleId, n)
    fmt = ExpandTokens(CStr(formats(LBound(formats) + stepNo - 1)))
    ApplyNumberFormat areas, fmt, actionLabel, actionLabel & " " & stepNo & "/" & n
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus actionLabel & " failed (" & Err.Description & ")"
End Sub

' ---- decimals ------------------------------------------------------------------

' Adds or removes one decimal in the active cell's format and applies the
' result to the whole selection.
Private Sub ShiftSelectionDecimals(ByVal delta As Long, ByVal actionLabel As String)
    Dim areas As Collection, oldFmt As String, newFmt As String
    On Error GoTo Fail
    ResetCycle
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub

    oldFmt = CStr(ActiveCell.NumberFormat)
    newFmt = ShiftDecimals(oldFmt, delta)
    If newFmt = oldFmt Then
        ShowStatus "no decimals to change"
        Exit Sub
    End If
    ApplyNumberFormat areas, newFmt, actionLabel, actionLabel
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ShowStatus actionLabel & " failed (" & Err.Description & ")"
End Sub

' Returns fmt with one decimal place added (delta = 1) or removed (delta = -1)
' in every numeric section. Brackets, padding and literal text are kept.
' General + increase gives step 2 of the number cycle.
Public Function ShiftDecimals(ByVal fmt As String, ByVal delta As Long) As String
    Dim mask As String, i As Long, startPos As Long, result As String
    Dim arr As Variant

    If LCase$(fmt) = "general" Then
        If delta > 0 Then
            arr = NumberCycle()
            ShiftDecimals = ExpandTokens(CStr(arr(LBound(arr) + 1)))
        Else
            ShiftDecimals = fmt
        End If
        Exit Function
    End If

    ' Split into sections on ; (ignoring any ; inside quotes) and shift each.
    mask = CodeMask(fmt)
    startPos = 1
    For i = 1 To Len(fmt)
        If Mid$(fmt, i, 1) = ";" And Mid$(mask, i, 1) = "c" Then
            result = result & ShiftSection(Mid$(fmt, startPos, i - startPos), _
                                           Mid$(mask, startPos, i - startPos), delta) & ";"
            startPos = i + 1
        End If
    Next i
    result = result & ShiftSection(Mid$(fmt, startPos), Mid$(mask, startPos), delta)
    ShiftDecimals = result
End Function

' Shifts decimals in one format section. m is the matching CodeMask slice.
' Sections with no digit placeholders (text, "-" zero sections) and date,
' time or fraction sections are returned unchanged.
Private Function ShiftSection(ByVal sec As String, ByVal m As String, ByVal delta As Long) As String
    Dim i As Long, ch As String, regionEnd As Long
    Dim dotPos As Long, lastDigit As Long, decCount As Long, trimmed As String

    ShiftSection = sec
    regionEnd = Len(sec)

    ' Leave dates, times and fractions alone; stop before any exponent.
    For i = 1 To Len(sec)
        If Mid$(m, i, 1) = "c" Then
            ch = Mid$(sec, i, 1)
            If InStr(1, "dmyhsDMYHS/", ch, vbBinaryCompare) > 0 Then Exit Function
            If (ch = "E" Or ch = "e") And regionEnd = Len(sec) Then regionEnd = i - 1
        End If
    Next i

    For i = 1 To regionEnd
        If Mid$(m, i, 1) = "c" Then
            ch = Mid$(sec, i, 1)
            If ch = "." And dotPos = 0 Then
                dotPos = i
            ElseIf ch = "0" Or ch = "#" Or ch = "?" Then
                lastDigit = i
                If dotPos > 0 Then decCount = decCount + 1
            End If
        End If
    Next i
    If lastDigit = 0 Then Exit Function

    If delta > 0 Then
        If dotPos = 0 Then
            ShiftSection = Left$(sec, lastDigit) & ".0" & Mid$(sec, lastDigit + 1)
        ElseIf decCount = 0 Then
            ShiftSection = Left$(sec, dotPos) & "0" & Mid$(sec, dotPos + 1)
        Else
            ' Repeat the last decimal placeholder (0 stays 0, # stays #).
            ShiftSection = Left$(sec, lastDigit) & Mid$(sec, lastDigit, 1) & Mid$(sec, lastDigit + 1)
        End If
    Else
        If decCount = 0 Then Exit Function
        trimmed = Left$(sec, lastDigit - 1) & Mid$(sec, lastDigit + 1)
        If decCount = 1 Then trimmed = Left$(trimmed, dotPos - 1) & Mid$(trimmed, dotPos + 1)
        ShiftSection = trimmed
    End If
End Function

' Returns a string as long as fmt: "c" where the character is format code,
' "l" where it is literal (quoted text, [brackets], or the character after
' \ _ or *).
Private Function CodeMask(ByVal fmt As String) As String
    Dim i As Long, ch As String, m As String
    Dim inQuote As Boolean, inBracket As Boolean
    i = 1
    Do While i <= Len(fmt)
        ch = Mid$(fmt, i, 1)
        If inQuote Then
            m = m & "l"
            If ch = """" Then inQuote = False
        ElseIf inBracket Then
            m = m & "l"
            If ch = "]" Then inBracket = False
        ElseIf ch = """" Then
            m = m & "l"
            inQuote = True
        ElseIf ch = "[" Then
            m = m & "l"
            inBracket = True
        ElseIf ch = "\" Or ch = "_" Or ch = "*" Then
            m = m & "ll"
            i = i + 1
        Else
            m = m & "c"
        End If
        i = i + 1
    Loop
    CodeMask = Left$(m, Len(fmt))
End Function

' ---- shared ----------------------------------------------------------------------

' Snapshots, applies fmt to every area, registers undo and shows msg.
' Errors go to the caller's handler.
Private Sub ApplyNumberFormat(ByVal areas As Collection, ByVal fmt As String, _
                              ByVal actionLabel As String, ByVal msg As String)
    Dim area As Range
    BeginSnapshot actionLabel, CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_NUMBERFORMAT
    Next area

    Application.ScreenUpdating = False
    For Each area In areas
        area.NumberFormat = fmt
    Next area
    Application.ScreenUpdating = True

    CommitSnapshot
    ShowStatus WithUndoNote(msg)
End Sub
