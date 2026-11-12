Attribute VB_Name = "modAlign"
Option Explicit
' Horizontal alignment and left indent cycles.

Public Sub CycleHorizontalAlign()
    Dim areas As Collection, area As Range, modes As Variant
    Dim stepNo As Long, n As Long, mode As String, alignValue As Long
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    modes = HAlignCycle()
    n = UBound(modes) - LBound(modes) + 1
    stepNo = NextStep("halign", n)
    mode = CStr(modes(LBound(modes) + stepNo - 1))
    alignValue = AlignConst(mode)

    ' Changing alignment can reset the indent, so both are saved for undo.
    BeginSnapshot "Alignment", CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_HALIGN
        SnapshotProperty area, PROP_INDENT
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        area.HorizontalAlignment = alignValue
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote("Alignment: " & mode & " (" & stepNo & "/" & n & ")")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Alignment failed (" & Err.Description & ")"
End Sub

Public Sub CycleIndent()
    Dim areas As Collection, area As Range, levels As Variant
    Dim stepNo As Long, n As Long, lvl As Long
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    levels = IndentCycle()
    n = UBound(levels) - LBound(levels) + 1
    stepNo = NextStep("indent", n)
    lvl = CLng(levels(LBound(levels) + stepNo - 1))

    BeginSnapshot "Indent", CellCount(areas)
    For Each area In areas
        SnapshotProperty area, PROP_HALIGN
        SnapshotProperty area, PROP_INDENT
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        ' Excel only indents left-, right- or distributed-aligned cells.
        If lvl > 0 Then
            If Not CanIndent(area.HorizontalAlignment) Then area.HorizontalAlignment = xlLeft
        End If
        area.IndentLevel = lvl
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote("Indent " & lvl & " (" & stepNo & "/" & n & ")")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Indent failed (" & Err.Description & ")"
End Sub

Private Function CanIndent(ByVal alignValue As Variant) As Boolean
    If IsNull(alignValue) Then Exit Function
    Select Case CLng(alignValue)
        Case xlLeft, xlRight, xlDistributed
            CanIndent = True
    End Select
End Function

Private Function AlignConst(ByVal mode As String) As Long
    Select Case LCase$(mode)
        Case "centre across selection", "center across selection"
            AlignConst = xlCenterAcrossSelection
        Case "centre", "center"
            AlignConst = xlCenter
        Case "left"
            AlignConst = xlLeft
        Case "right"
            AlignConst = xlRight
        Case Else
            AlignConst = xlGeneral
    End Select
End Function
