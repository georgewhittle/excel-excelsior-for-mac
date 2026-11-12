Attribute VB_Name = "modUndo"
Option Explicit
' One-level undo: snapshot the properties an action changes, then register
' Application.OnUndo so Edit > Undo / Cmd+Z / Ctrl+Z restores them.

Public Const PROP_NUMBERFORMAT As Long = 1
Public Const PROP_FONTCOLOR As Long = 2
Public Const PROP_FILL As Long = 3      ' stored as "pattern|colour"
Public Const PROP_HALIGN As Long = 4
Public Const PROP_INDENT As Long = 5
' Borders: PROP_BORDER_BASE + edge (xlEdgeLeft 7 .. xlEdgeRight 10),
' stored as "linestyle|weight|colour".
Public Const PROP_BORDER_BASE As Long = 100

Private mEntries As Collection    ' each item: Array(worksheet, address, propId, value)
Private mLabel As String
Private mAvailable As Boolean

' Call before changing anything. Over UNDO_CAP cells, no snapshot is kept.
Public Sub BeginSnapshot(ByVal actionLabel As String, ByVal nCells As Double)
    Set mEntries = New Collection
    mLabel = actionLabel
    mAvailable = (nCells <= UNDO_CAP)
End Sub

Public Function SnapshotAvailable() As Boolean
    SnapshotAvailable = mAvailable
End Function

' Records one property for a range: one entry if uniform, else one per cell.
Public Sub SnapshotProperty(ByVal rng As Range, ByVal propId As Long)
    Dim v As Variant, c As Range
    If Not mAvailable Then Exit Sub
    v = ReadProp(rng, propId)
    If Not IsNull(v) Then
        AddEntry rng, propId, v
    Else
        For Each c In rng.Cells
            AddEntry c, propId, ReadProp(c, propId)
        Next c
    End If
End Sub

' Call after the change. Registers the undo, or drops the snapshot if over the cap.
Public Sub CommitSnapshot()
    If mAvailable Then
        Application.OnUndo "Undo " & mLabel, QualifiedProc("UndoLastAction")
    Else
        Set mEntries = Nothing
    End If
End Sub

' Entry point for OnUndo (Cmd+Z / Edit > Undo).
Public Sub UndoLastAction()
    Dim i As Long, e As Variant, ws As Worksheet
    On Error GoTo Fail
    ResetCycle
    If mEntries Is Nothing Then
        ShowStatus "nothing to undo"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For i = mEntries.Count To 1 Step -1
        e = mEntries(i)
        Set ws = e(0)
        WriteProp ws.Range(e(1)), CLng(e(2)), e(3)
    Next i
    Application.ScreenUpdating = True
    Set mEntries = Nothing
    ShowStatus "undone: " & mLabel
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Set mEntries = Nothing
    ShowStatus "couldn't undo (" & Err.Description & ")"
End Sub

Private Sub AddEntry(ByVal rng As Range, ByVal propId As Long, ByVal v As Variant)
    mEntries.Add Array(rng.Worksheet, rng.Address(False, False), propId, v)
End Sub

' Returns Null when the range is mixed, so SnapshotProperty goes cell by cell.
Private Function ReadProp(ByVal rng As Range, ByVal propId As Long) As Variant
    Dim p As Variant, c As Variant, w As Variant
    Select Case propId
        Case PROP_NUMBERFORMAT
            ReadProp = rng.NumberFormat
        Case PROP_FONTCOLOR
            ReadProp = rng.Font.Color
        Case PROP_FILL
            p = rng.Interior.Pattern
            c = rng.Interior.Color
            If IsNull(p) Or IsNull(c) Then
                ReadProp = Null
            Else
                ReadProp = CStr(p) & "|" & CStr(c)
            End If
        Case PROP_HALIGN
            ReadProp = rng.HorizontalAlignment
        Case PROP_INDENT
            ReadProp = rng.IndentLevel
        Case PROP_BORDER_BASE + 1 To PROP_BORDER_BASE + 12
            With rng.Borders(propId - PROP_BORDER_BASE)
                p = .LineStyle
                w = .Weight
                c = .Color
            End With
            If IsNull(p) Or IsNull(w) Or IsNull(c) Then
                ReadProp = Null
            Else
                ReadProp = CStr(p) & "|" & CStr(w) & "|" & CStr(c)
            End If
    End Select
End Function

Private Sub WriteProp(ByVal rng As Range, ByVal propId As Long, ByVal v As Variant)
    Dim parts() As String
    Select Case propId
        Case PROP_NUMBERFORMAT
            rng.NumberFormat = v
        Case PROP_FONTCOLOR
            rng.Font.Color = v
        Case PROP_FILL
            parts = Split(CStr(v), "|")
            If CLng(parts(0)) = xlNone Then
                rng.Interior.Pattern = xlNone
            Else
                rng.Interior.Pattern = CLng(parts(0))
                rng.Interior.Color = CDbl(parts(1))
            End If
        Case PROP_HALIGN
            rng.HorizontalAlignment = v
        Case PROP_INDENT
            rng.IndentLevel = v
        Case PROP_BORDER_BASE + 1 To PROP_BORDER_BASE + 12
            parts = Split(CStr(v), "|")
            With rng.Borders(propId - PROP_BORDER_BASE)
                If CLng(parts(0)) = xlNone Then
                    .LineStyle = xlNone
                Else
                    .LineStyle = CLng(parts(0))
                    .Weight = CLng(parts(1))
                    .Color = CDbl(parts(2))
                End If
            End With
    End Select
End Sub
