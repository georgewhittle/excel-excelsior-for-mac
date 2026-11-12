Attribute VB_Name = "modUtil"
Option Explicit
' Selection guards, status bar messages and small helpers.

Private mStatusClearAt As Date
Private mTimerPending As Boolean

' Full name for OnKey / OnUndo / OnTime, e.g. 'ExcelExcelsiorForMac.xlam'!CycleNumber
Public Function QualifiedProc(ByVal procName As String) As String
    QualifiedProc = "'" & ThisWorkbook.Name & "'!" & procName
End Function

' The selection's areas, with whole rows/columns clipped to the used range.
' Returns an empty Collection if the selection isn't cells.
Public Function TargetAreas() As Collection
    Dim result As Collection
    Dim sel As Range, area As Range, clipped As Range
    Set result = New Collection
    Set TargetAreas = result
    If TypeName(Selection) <> "Range" Then Exit Function
    Set sel = Selection
    For Each area In sel.Areas
        Set clipped = ClipToUsed(area)
        If Not clipped Is Nothing Then result.Add clipped
    Next area
End Function

Private Function ClipToUsed(ByVal area As Range) As Range
    Dim ws As Worksheet
    Set ws = area.Worksheet
    If area.Rows.Count = ws.Rows.Count Or area.Columns.Count = ws.Columns.Count Then
        Set ClipToUsed = Application.Intersect(area, ws.UsedRange)
    Else
        Set ClipToUsed = area
    End If
End Function

Public Function CellCount(ByVal areas As Collection) As Double
    Dim area As Range, total As Double
    For Each area In areas
        total = total + area.CountLarge
    Next area
    CellCount = total
End Function

' TargetAreas, or Nothing (with a status message) if there's nothing to
' format or the sheet is protected.
Public Function FormattableAreas() As Collection
    Dim areas As Collection
    Set areas = TargetAreas()
    If areas.Count = 0 Then
        ShowStatus "select some cells first"
    ElseIf Not CanFormat(areas(1).Worksheet) Then
        ShowStatus "sheet is protected"
    Else
        Set FormattableAreas = areas
    End If
End Function

' Adds a note to a status message when the last action had no undo snapshot.
Public Function WithUndoNote(ByVal msg As String) As String
    If SnapshotAvailable() Then
        WithUndoNote = msg
    Else
        WithUndoNote = msg & " (undo unavailable: over " & Format(UNDO_CAP, "#,##0") & " cells)"
    End If
End Function

' True if cells on this sheet can be formatted.
Public Function CanFormat(ByVal ws As Worksheet) As Boolean
    If Not ws.ProtectContents Then
        CanFormat = True
    Else
        CanFormat = ws.Protection.AllowFormattingCells
    End If
End Function

' Shows a message on the status bar for STATUS_SECONDS, then clears it.
Public Sub ShowStatus(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = ADDIN_NAME & ": " & msg
    CancelStatusTimer
    mStatusClearAt = Now + TimeSerial(0, 0, STATUS_SECONDS)
    Application.OnTime mStatusClearAt, QualifiedProc("ClearStatus")
    mTimerPending = True
End Sub

Public Sub ClearStatus()
    On Error Resume Next
    mTimerPending = False
    Application.StatusBar = False
End Sub

' Cancels the pending clear, so no timer can reopen the add-in after it unloads.
Public Sub CancelStatusTimer()
    On Error Resume Next
    If mTimerPending Then
        Application.OnTime mStatusClearAt, QualifiedProc("ClearStatus"), , False
    End If
    mTimerPending = False
End Sub
