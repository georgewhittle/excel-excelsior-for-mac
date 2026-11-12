Attribute VB_Name = "modConfig"
Option Explicit
' ============================================================================
' Excel Excelsior For Mac: every editable value lives in this module.
' What each cycle does: docs/cycles.md. How to edit: docs/customising.md.
'
' To change a format: edit the string below, then rebuild (docs/building.md).
'   - Inside a VBA string, a double quote is written twice: "" means ".
'   - {dash} is replaced with an en dash (ChrW(8211)) when the add-in runs.
'     Never type the dash itself into this file: the Mac editor mangles it.
' ============================================================================

Public Const ADDIN_NAME As String = "Excel Excelsior"

' ---- Shortcuts (OnKey syntax: ^ = Ctrl, + = Shift, % = Option/Alt) --------
' On Mac, ^ is the Control key, not Cmd.
Public Const KEY_NUMBER_CYCLE As String = "^+1"     ' Ctrl+Shift+1
Public Const KEY_PERCENT_CYCLE As String = "^+5"    ' Ctrl+Shift+5
Public Const KEY_MULTIPLE_CYCLE As String = "^+8"   ' Ctrl+Shift+8
Public Const KEY_DATE_CYCLE As String = "^+2"       ' Ctrl+Shift+2
Public Const KEY_DECIMALS_UP As String = "^,"       ' Ctrl+,
Public Const KEY_DECIMALS_DOWN As String = "^."     ' Ctrl+.
Public Const KEY_BLUE_BLACK As String = "^;"        ' Ctrl+;
#If Mac Then
' Excel for Mac ignores add-in shortcuts that use Cmd (*) or Option (%) with a
' letter, so Mac shortcuts use Ctrl and Ctrl+Shift only.
Public Const KEY_AUTOCOLOR As String = "^+a"        ' Ctrl+Shift+A
#Else
Public Const KEY_AUTOCOLOR As String = "^%a"        ' Ctrl+Alt+A
#End If
Public Const KEY_FONT_CYCLE As String = "^'"        ' Ctrl+'
Public Const KEY_FILL_CYCLE As String = "^+k"       ' Ctrl+Shift+K
Public Const KEY_BOTTOM_BORDER As String = "^+b"    ' Ctrl+Shift+B
Public Const KEY_OUTSIDE_BORDER As String = "^+7"   ' Ctrl+Shift+7
Public Const KEY_HALIGN_CYCLE As String = "^+h"     ' Ctrl+Shift+H
Public Const KEY_INDENT_CYCLE As String = "^+i"     ' Ctrl+Shift+I

' ---- Undo -------------------------------------------------------------------
' Selections bigger than this are formatted without an undo snapshot.
Public Const UNDO_CAP As Long = 50000

' Backup undo key, in case Cmd+Z (Application.OnUndo) doesn't work on Mac.
Public Const KEY_UNDO_FALLBACK As String = "^%z"    ' Ctrl+Option+Z
' Off: Cmd+Z works on Mac, and Option shortcuts don't register.
#If Mac Then
Public Const UNDO_FALLBACK_ON As Boolean = False
#Else
Public Const UNDO_FALLBACK_ON As Boolean = False
#End If

' ---- Status bar ---------------------------------------------------------------
Public Const STATUS_SECONDS As Long = 4

' AutoColor: numbers in these function arguments are positions or indexes,
' not hardcodes (e.g. the 3 in VLOOKUP(x, r, 3, 0)). Format: NAME:arg,arg|...
' Numbers 0 and 1 are never treated as hardcodes.
Public Const POSITION_ARGS As String = _
    "INDEX:2,3,4|MATCH:3|XMATCH:3,4|VLOOKUP:3,4|HLOOKUP:3,4|XLOOKUP:5,6|" & _
    "OFFSET:2,3,4,5|CHOOSE:1|SUBTOTAL:1|AGGREGATE:1,2|ROUND:2|ROUNDUP:2|ROUNDDOWN:2|" & _
    "LEFT:2|RIGHT:2|MID:2,3|LARGE:2|SMALL:2|EOMONTH:2|EDATE:2|WEEKDAY:2|WEEKNUM:2|" & _
    "MOD:2|RANK:3|RANK.EQ:3|RANK.AVG:3|FIND:3|SEARCH:3|SUBSTITUTE:4|INDIRECT:2|" & _
    "ADDRESS:1,2,3,4|SEQUENCE:1,2,3,4|TAKE:2,3|DROP:2,3|CHOOSECOLS:2,3,4,5,6|" & _
    "CHOOSEROWS:2,3,4,5,6|SORT:2,3,4|FILTER:3|NETWORKDAYS:3|WORKDAY:3|YEARFRAC:3|" & _
    "TEXTBEFORE:3,4|TEXTAFTER:3,4"

' Fill cycle entry meaning "no fill".
Public Const NO_FILL As Long = -1

' Colour for every border the add-in draws (RGB 0,0,0 = black).
Public Const BORDER_COLOUR As Long = 0

' ---- Procedures below this line. VBA needs every Const above them. ---------

' ---- 1. Number cycle (Ctrl+Shift+1) -------------------------------------------
Public Function NumberCycle() As Variant
    NumberCycle = Array( _
        "#,##0_);(#,##0);""{dash}""_);@_)", _
        "#,##0.0_);(#,##0.0);""{dash}""_);@_)", _
        "#,##0.00_);(#,##0.00);""{dash}""_);@_)", _
        "0_);(0);0_);@_)")
End Function

' ---- 2. Percent cycle (Ctrl+Shift+5) ------------------------------------------
' Steps 4 and 5 expect the cell to hold basis points (150 = 150bps).
Public Function PercentCycle() As Variant
    PercentCycle = Array( _
        "0.0%_);(0.0%);""{dash}""_)", _
        "0.00%_);(0.00%);""{dash}""_)", _
        "0%_);(0%);""{dash}""_)", _
        "#,##0""bps""_);(#,##0""bps"");""{dash}""_)", _
        """BBSW+""0_);""BBSW-""0_);""BBSW""_)")
End Function

' ---- 3. Multiple cycle (Ctrl+Shift+8) -----------------------------------------
Public Function MultipleCycle() As Variant
    MultipleCycle = Array( _
        "0.0""x""_);(0.0""x"");""{dash}""_)", _
        "0.00""x""_);(0.00""x"");""{dash}""_)", _
        "0""x""_);(0""x"");""{dash}""_)")
End Function

' ---- 4. Date cycle (Ctrl+Shift+2) ---------------------------------------------
' "FY"yy is only right for June year-end dates.
Public Function DateCycle() As Variant
    DateCycle = Array( _
        "d-mmm-yy", _
        "mmm-yy", _
        "dd/mm/yyyy", _
        """FY""yy", _
        "yyyy")
End Function

' ---- 7-8. Colours used by blue-black and AutoColor --------------------------------
Public Function InputBlue() As Long
    InputBlue = RGB(0, 0, 255)
End Function
Public Function FormulaBlack() As Long
    FormulaBlack = RGB(0, 0, 0)
End Function
Public Function SheetLinkGreen() As Long
    SheetLinkGreen = RGB(0, 128, 0)
End Function
Public Function BookLinkPurple() As Long
    BookLinkPurple = RGB(112, 48, 160)
End Function

' ---- 9. Font colour cycle (Ctrl+') ---------------------------------------------
Public Function FontColourCycle() As Variant
    FontColourCycle = Array( _
        RGB(0, 0, 255), _
        RGB(0, 0, 0), _
        RGB(0, 128, 0), _
        RGB(192, 0, 0), _
        RGB(112, 48, 160))
End Function

' ---- 10. Fill colour cycle (Ctrl+Shift+K) --------------------------------------
Public Function FillColourCycle() As Variant
    FillColourCycle = Array( _
        RGB(242, 242, 242), _
        RGB(221, 235, 247), _
        RGB(255, 242, 204), _
        RGB(226, 239, 218), _
        NO_FILL)
End Function

' ---- 11. Bottom border cycle (Ctrl+Shift+B) ------------------------------------
' Styles: thin, medium, double, none. Bottom edge of the selection only.
Public Function BottomBorderCycle() As Variant
    BottomBorderCycle = Array("thin", "medium", "double", "none")
End Function

' ---- 12. Outside border cycle (Ctrl+Shift+7) -----------------------------------
' Styles: thin, medium, double, none. Around the selection only.
Public Function OutsideBorderCycle() As Variant
    OutsideBorderCycle = Array("thin", "medium", "none")
End Function

' ---- 13. Horizontal alignment cycle (Ctrl+Shift+H) -----------------------------
' Choices: centre across selection, centre, left, right, general.
Public Function HAlignCycle() As Variant
    HAlignCycle = Array("centre across selection", "centre", "left", "right", "general")
End Function

' ---- 14. Left indent cycle (Ctrl+Shift+I) --------------------------------------
Public Function IndentCycle() As Variant
    IndentCycle = Array(1, 2, 3, 0)
End Function

' Turns {dash} into a real en dash.
Public Function ExpandTokens(ByVal s As String) As String
    ExpandTokens = Replace(s, "{dash}", ChrW(8211))
End Function
