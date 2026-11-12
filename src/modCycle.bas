Attribute VB_Name = "modCycle"
Option Explicit
' Generic cycle engine.
' The first press always gives step 1. Pressing the same shortcut again on the
' same selection advances one step and wraps. Any selection change, sheet or
' workbook switch (see clsAppEvents), different shortcut or undo resets it.

Private mLastId As String
Private mLastSel As String
Private mStep As Long

' Returns the step to apply now, from 1 to itemCount.
Public Function NextStep(ByVal cycleId As String, ByVal itemCount As Long) As Long
    Dim sel As String
    sel = SelectionKey()
    If mStep >= 1 And cycleId = mLastId And sel = mLastSel And sel <> "" Then
        mStep = (mStep Mod itemCount) + 1
    Else
        mStep = 1
    End If
    mLastId = cycleId
    mLastSel = sel
    NextStep = mStep
End Function

Public Sub ResetCycle()
    mLastId = ""
    mLastSel = ""
    mStep = 0
End Sub

Private Function SelectionKey() As String
    Dim k As String
    On Error Resume Next
    k = ActiveWorkbook.Name & "|" & ActiveSheet.Name & "|" & Selection.Address(False, False)
    If k = "" Then
        ' Very long multi-area addresses can fail; fall back to a shorter key.
        Err.Clear
        k = ActiveWorkbook.Name & "|" & ActiveSheet.Name & "|" & Selection.Areas.Count & _
            "|" & Selection.Cells(1, 1).Address & "|" & Selection.CountLarge
    End If
    SelectionKey = k
End Function
