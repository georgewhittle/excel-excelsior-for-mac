Attribute VB_Name = "clsAppEvents"
Option Explicit
' Application events: any selection change or sheet/workbook switch resets
' the current cycle, as the spec requires.

Public WithEvents App As Application

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ResetCycle
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    ResetCycle
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ResetCycle
End Sub
