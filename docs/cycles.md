# What each shortcut does

Every format, colour and shortcut below is a default. Change any of them in [`src/modConfig.bas`](../src/modConfig.bas) (see [customising.md](customising.md)).

## Rules for every feature
- **Cycles:** the first press gives the first option. Pressing the same shortcut again on the same selection moves one step on, and wraps after the last. Changing the selection, the sheet or the shortcut starts again from the first option.
- **Undo:** one level. Cmd+Z (or Edit > Undo) restores exactly what the add-in's last action changed. Selections over 50,000 cells are formatted without undo, and the status bar says so.
- **Selections:** whole rows or columns are clipped to the used range. Multi-area selections (Cmd-click) work.
- **Errors:** protected sheets and other failures show a status bar message. You never get a VBA error box.

## Number formats
Numbers are right-aligned with negatives in brackets. Positives are padded so they line up with the brackets, zeros show as an en dash, and text is padded to the same edge.

### Number cycle (Ctrl+Shift+1)
| Step | Format | 1234.5 | −1234.5 | 0 |
|---|---|---|---|---|
| 1 | `#,##0_);(#,##0);"–"_);@_)` | 1,235 | (1,235) | – |
| 2 | `#,##0.0_);(#,##0.0);"–"_);@_)` | 1,234.5 | (1,234.5) | – |
| 3 | `#,##0.00_);(#,##0.00);"–"_);@_)` | 1,234.50 | (1,234.50) | – |
| 4 | `0_);(0);0_);@_)` | 1235 | (1235) | 0 |

Step 4 has no thousands separator and shows a real zero, for years, IDs and counts.

### Percent cycle (Ctrl+Shift+5)
| Step | Format | Cell value | Shows |
|---|---|---|---|
| 1 | `0.0%_);(0.0%);"–"_)` | 0.0525 | 5.3% |
| 2 | `0.00%_);(0.00%);"–"_)` | 0.0525 | 5.25% |
| 3 | `0%_);(0%);"–"_)` | 0.0525 | 5% |
| 4 | `#,##0"bps"_);(#,##0"bps");"–"_)` | 150 | 150bps |
| 5 | `"BBSW+"0_);"BBSW-"0_);"BBSW"_)` | 350 | BBSW+350 |

Steps 4 and 5 expect the cell to hold basis points (150 = 150bps). A cell holding 1.50% cycled to step 4 shows "2bps", so keep bps inputs in their own cells. BBSW is the Australian bank bill rate; change the label in `modConfig` to SOFR, SONIA or whatever you use.

### Multiple cycle (Ctrl+Shift+8)
| Step | Format | 8.94 shows |
|---|---|---|
| 1 | `0.0"x"_);(0.0"x");"–"_)` | 8.9x |
| 2 | `0.00"x"_);(0.00"x");"–"_)` | 8.94x |
| 3 | `0"x"_);(0"x");"–"_)` | 9x |

### Date cycle (Ctrl+Shift+2)
| Step | Format | 30 Jun 2027 shows | Use |
|---|---|---|---|
| 1 | `d-mmm-yy` | 30-Jun-27 | Default date |
| 2 | `mmm-yy` | Jun-27 | Monthly or quarterly headers |
| 3 | `dd/mm/yyyy` | 30/06/2027 | Full date (day first) |
| 4 | `"FY"yy` | FY27 | Annual headers. Only right for June year ends |
| 5 | `yyyy` | 2027 | Calendar year headers |

### Increase / decrease decimals (Ctrl+, / Ctrl+.)
Adds or removes one decimal place in every numeric part of the active cell's format, then applies the result to the whole selection. Brackets, padding and text like `x`, `%`, `bps` and `BBSW+` are kept. On a General cell, increase gives step 2 of the number cycle. Decreasing at zero decimals does nothing.

## Colours

### Blue-black toggle (Ctrl+;)
If the active cell's font is input blue, the selection turns black. Otherwise it turns input blue.

### AutoColor (Ctrl+Shift+A)
| Cell contains | Colour | RGB |
|---|---|---|
| A typed number | Input blue | 0, 0, 255 |
| A formula with a typed number in it (`=A1*1.1`) | Input blue | 0, 0, 255 |
| A formula linking to another workbook | Purple | 112, 48, 160 |
| A formula linking to another sheet | Green | 0, 128, 0 |
| Any other formula | Black | 0, 0, 0 |

- When a formula fits more than one row, the first matching row in the table wins.
- Text, blanks and cells already in a colour outside this scheme (such as red) are left alone.
- Numbers that are positions or switches rather than assumptions don't count as typed numbers: `0`, `1`, and arguments like the column number in `VLOOKUP` or the digits in `ROUND`. The list is `POSITION_ARGS` in `modConfig`. When in doubt, the add-in treats the cell as a plain formula.

### Font colour cycle (Ctrl+')
Input blue (0,0,255) → black (0,0,0) → green (0,128,0) → red (192,0,0) → purple (112,48,160)

### Fill colour cycle (Ctrl+Shift+K)
| Step | Colour | RGB | Typical use |
|---|---|---|---|
| 1 | Light grey | 242, 242, 242 | Headers, subtotal rows |
| 2 | Pale blue | 221, 235, 247 | Input or assumption blocks |
| 3 | Pale yellow | 255, 242, 204 | Flags, items to review |
| 4 | Pale green | 226, 239, 218 | Checks that pass |
| 5 | No fill | | Clear |

## Borders, alignment and indent

### Bottom border cycle (Ctrl+Shift+B)
Thin → medium → double → none, in black, on the bottom edge of the selection only. Double is for grand totals.

### Outside border cycle (Ctrl+Shift+7)
Thin → medium → none, in black, around the selection only. Inside borders are never touched.

### Alignment cycle (Ctrl+Shift+H)
Centre across selection → centre → left → right → general. Centre across selection comes first because it's the usual choice for headers and avoids merged cells.

### Indent cycle (Ctrl+Shift+I)
Indent 1 → 2 → 3 → 0.
