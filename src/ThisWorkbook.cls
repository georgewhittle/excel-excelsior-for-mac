Option Explicit
' The add-in's own workbook events. Workbook_Open (not Auto_Open) is used
' because Auto_Open doesn't run reliably for add-ins loaded at startup on Mac.

Private Sub Workbook_Open()
    RegisterShortcuts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UnregisterShortcuts
End Sub
