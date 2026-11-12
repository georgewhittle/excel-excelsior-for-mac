# Changing a format, colour or shortcut

Everything you'd want to change is in [`src/modConfig.bas`](../src/modConfig.bas), a plain text file. Edit it in any text editor, then rebuild the add-in ([building.md](building.md)).

## Formats
Each cycle is a list of Excel number formats, for example:
```vb
Public Function MultipleCycle() As Variant
    MultipleCycle = Array( _
        "0.0""x""_);(0.0""x"");""{dash}""_)", _
        "0.00""x""_);(0.00""x"");""{dash}""_)", _
        "0""x""_);(0""x"");""{dash}""_)")
End Function
```
- Write each format as you would in Excel's Format Cells box, then **double every `"`**: `"x"` becomes `""x""`.
- `{dash}` turns into the en dash shown for zeros. **Never type the dash, or any other non-ASCII character, into a `.bas` file**: the Mac VBA editor mangles it.
- Add or remove lines to make a cycle longer or shorter. Every line but the last ends with `, _`.

## Colours
Colours are written as `RGB(red, green, blue)`. The AutoColor and blue-black colours are the `InputBlue`, `FormulaBlack`, `SheetLinkGreen` and `BookLinkPurple` functions.

## Shortcuts
Shortcuts are the `KEY_` constants near the top, in Excel's OnKey codes:

| Code | Key |
|---|---|
| `^` | Control |
| `+` | Shift |
| `%` | Option |

So `"^+1"` is Control+Shift+1. On Mac, stick to Control (`^`) and Control+Shift (`^+`): Excel for Mac ignores add-in shortcuts that use Cmd or Option.

## Rules that keep it compiling
- Keep every `Const` above the first `Function` or `Sub` in the file. VBA refuses to compile otherwise.
- Use plain ASCII characters only.
- After rebuilding, check it compiles: open the VBA editor (Tools > Macro > Visual Basic Editor), click **VBAProject (ExcelExcelsiorForMac.xlam)**, then Debug > Compile VBAProject. Nothing happening means it's fine.
