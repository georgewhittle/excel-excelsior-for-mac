Attribute VB_Name = "modBorders"
Option Explicit
' Bottom and outside border cycles. Inside borders are never touched.

Public Sub CycleBottomBorder()
    Dim areas As Collection, area As Range, styles As Variant
    Dim stepNo As Long, n As Long, style As String
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    styles = BottomBorderCycle()
    n = UBound(styles) - LBound(styles) + 1
    stepNo = NextStep("bottomborder", n)
    style = CStr(styles(LBound(styles) + stepNo - 1))

    BeginSnapshot "Bottom border", CellCount(areas)
    For Each area In areas
        SnapshotProperty area.Rows(area.Rows.Count), PROP_BORDER_BASE + xlEdgeBottom
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        SetEdge area, xlEdgeBottom, style
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote("Bottom border: " & style & " (" & stepNo & "/" & n & ")")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Bottom border failed (" & Err.Description & ")"
End Sub

Public Sub CycleOutsideBorder()
    Dim areas As Collection, area As Range, styles As Variant
    Dim stepNo As Long, n As Long, style As String
    On Error GoTo Fail
    Set areas = FormattableAreas()
    If areas Is Nothing Then Exit Sub
    styles = OutsideBorderCycle()
    n = UBound(styles) - LBound(styles) + 1
    stepNo = NextStep("outsideborder", n)
    style = CStr(styles(LBound(styles) + stepNo - 1))

    BeginSnapshot "Outside border", CellCount(areas)
    For Each area In areas
        SnapshotProperty area.Rows(1), PROP_BORDER_BASE + xlEdgeTop
        SnapshotProperty area.Rows(area.Rows.Count), PROP_BORDER_BASE + xlEdgeBottom
        SnapshotProperty area.Columns(1), PROP_BORDER_BASE + xlEdgeLeft
        SnapshotProperty area.Columns(area.Columns.Count), PROP_BORDER_BASE + xlEdgeRight
    Next area
    Application.ScreenUpdating = False
    For Each area In areas
        SetEdge area, xlEdgeTop, style
        SetEdge area, xlEdgeBottom, style
        SetEdge area, xlEdgeLeft, style
        SetEdge area, xlEdgeRight, style
    Next area
    Application.ScreenUpdating = True
    CommitSnapshot
    ShowStatus WithUndoNote("Outside border: " & style & " (" & stepNo & "/" & n & ")")
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    ResetCycle
    ShowStatus "Outside border failed (" & Err.Description & ")"
End Sub

' Applies one style (thin, medium, double, none) to one edge of rng.
Private Sub SetEdge(ByVal rng As Range, ByVal edge As Long, ByVal style As String)
    With rng.Borders(edge)
        Select Case LCase$(style)
            Case "none"
                .LineStyle = xlNone
            Case "double"
                .LineStyle = xlDouble
                .Weight = xlThick
                .Color = BORDER_COLOUR
            Case "medium"
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = BORDER_COLOUR
            Case Else
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = BORDER_COLOUR
        End Select
    End With
End Sub
