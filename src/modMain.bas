Attribute VB_Name = "modMain"
Option Explicit
' Shortcut registration. Called from ThisWorkbook on load and close.

Private mEvents As clsAppEvents

' Every shortcut: Array(key, procedure). Keys are set in modConfig.
Private Function Bindings() As Variant
    Bindings = Array( _
        Array(KEY_NUMBER_CYCLE, "CycleNumber"), _
        Array(KEY_PERCENT_CYCLE, "CyclePercent"), _
        Array(KEY_MULTIPLE_CYCLE, "CycleMultiple"), _
        Array(KEY_DATE_CYCLE, "CycleDate"), _
        Array(KEY_DECIMALS_UP, "IncreaseDecimals"), _
        Array(KEY_DECIMALS_DOWN, "DecreaseDecimals"), _
        Array(KEY_BLUE_BLACK, "ToggleBlueBlack"), _
        Array(KEY_AUTOCOLOR, "AutoColorSelection"), _
        Array(KEY_FONT_CYCLE, "CycleFontColour"), _
        Array(KEY_FILL_CYCLE, "CycleFillColour"), _
        Array(KEY_BOTTOM_BORDER, "CycleBottomBorder"), _
        Array(KEY_OUTSIDE_BORDER, "CycleOutsideBorder"), _
        Array(KEY_HALIGN_CYCLE, "CycleHorizontalAlign"), _
        Array(KEY_INDENT_CYCLE, "CycleIndent"))
End Function

Public Sub RegisterShortcuts()
    Dim b As Variant
    On Error Resume Next
    For Each b In Bindings()
        Application.OnKey CStr(b(0)), QualifiedProc(CStr(b(1)))
    Next b
    If UNDO_FALLBACK_ON Then Application.OnKey KEY_UNDO_FALLBACK, QualifiedProc("UndoLastAction")

    ' Watches selection and sheet changes so cycles reset.
    Set mEvents = New clsAppEvents
    Set mEvents.App = Application

    ShowStatus "loaded, shortcuts on"
End Sub

' Passing no procedure to OnKey gives the key back to Excel.
Public Sub UnregisterShortcuts()
    Dim b As Variant
    On Error Resume Next
    For Each b In Bindings()
        Application.OnKey CStr(b(0))
    Next b
    If UNDO_FALLBACK_ON Then Application.OnKey KEY_UNDO_FALLBACK
    Set mEvents = Nothing
    CancelStatusTimer
    Application.StatusBar = False
End Sub
