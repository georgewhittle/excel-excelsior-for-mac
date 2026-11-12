# Excel Excelsior For Mac

A free Excel add-in for Mac that puts the formatting you use most in a financial model on a single keyboard shortcut each. Press a shortcut once for the first format, and keep pressing to cycle through the options.

- Number, percent, multiple and date format cycles
- Increase and decrease decimals without breaking brackets or suffixes
- Blue-black toggle, font and fill colour cycles
- **AutoColor**: colours a block of cells by what's in them (hardcodes blue, formulas black, links to other sheets green, links to other workbooks purple)
- Bottom border, outside border, alignment and indent cycles
- Cmd+Z undoes the add-in's last action

> **Mac only.** Built and tested on Excel for Mac (Microsoft 365). Windows isn't supported yet.

## Install

1. Download **ExcelExcelsiorForMac.xlam** from the [latest release](../../releases/latest).
2. In Finder, choose **Go > Go to Folder**, paste the path below and press Return:
   ```
   ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized
   ```
   Move the downloaded file into that folder.
3. Open Excel. Choose **Tools > Excel Add-ins**, tick **ExcelExcelsiorForMac** and click OK. If it isn't listed, click **Browse** and pick the file.
4. If Excel asks about macros, click **Enable Macros**.

You'll see **Excel Excelsior: loaded, shortcuts on** in the status bar at the bottom of the window. From now on it loads every time Excel starts.

**If Excel blocks the macros:** macOS marks downloaded files as coming from the internet. Open Terminal and run this, then restart Excel:
```bash
xattr -d com.apple.quarantine ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content.localized/Add-Ins.localized/ExcelExcelsiorForMac.xlam
```
Also check **Excel > Preferences > Security** (or **Tools > Macro > Security**): macros must be set to at least "Disable all macros with notification".

**To turn it off**, untick it in Tools > Excel Add-ins. Excel's own shortcuts come straight back. To remove it completely, delete the file from the Add-ins folder.

## Shortcuts

"Ctrl" means the **Control** key, not Cmd. Excel for Mac ignores add-in shortcuts that use Cmd or Option, so every shortcut uses Control or Control+Shift.

| Feature | Shortcut | What it does |
|---|---|---|
| Number cycle | Ctrl+Shift+1 | `1,235` → `1,234.5` → `1,234.50` → `1235` |
| Percent cycle | Ctrl+Shift+5 | `5.3%` → `5.25%` → `5%` → `150bps` → `BBSW+350` |
| Multiple cycle | Ctrl+Shift+8 | `8.9x` → `8.94x` → `9x` |
| Date cycle | Ctrl+Shift+2 | `30-Jun-27` → `Jun-27` → `30/06/2027` → `FY27` → `2027` |
| Increase decimals | Ctrl+, | One more decimal place, keeping brackets and suffixes |
| Decrease decimals | Ctrl+. | One fewer decimal place |
| Blue-black toggle | Ctrl+; | Input blue ↔ black |
| AutoColor | Ctrl+Shift+A | Colours each cell by what it contains |
| Font colour cycle | Ctrl+' | Blue → black → green → red → purple |
| Fill colour cycle | Ctrl+Shift+K | Grey → pale blue → pale yellow → pale green → none |
| Bottom border cycle | Ctrl+Shift+B | Thin → medium → double → none |
| Outside border cycle | Ctrl+Shift+7 | Thin → medium → none |
| Alignment cycle | Ctrl+Shift+H | Centre across selection → centre → left → right → general |
| Indent cycle | Ctrl+Shift+I | Indent 1 → 2 → 3 → 0 |
| Undo | Cmd+Z | Undoes the add-in's last action |

Numbers show negatives in brackets and zeros as a dash. Every cycle starts at the first option. Pressing the same shortcut again on the same cells moves to the next option, and changing the selection or the sheet starts again from the first.

Full details of every format, colour and rule: [docs/cycles.md](docs/cycles.md).

### Excel shortcuts it replaces

While the add-in is loaded, these shortcuts do the add-in's job instead of Excel's: Ctrl+; (insert today's date), and any other Excel shortcut on the list above. Untick the add-in to get them back.

## Make it your own

The formats, colours and shortcuts are house-style defaults (Australian dates and a BBSW margin format, for example). They all live in one file, [`src/modConfig.bas`](src/modConfig.bas), so changing one is a one-line edit followed by a rebuild. See [docs/customising.md](docs/customising.md) and [docs/building.md](docs/building.md).

## Good to know

- **Undo is one level.** Cmd+Z undoes the add-in's last action only. For selections over 50,000 cells there's no undo, and the status bar tells you so.
- **Whole rows and columns** are clipped to the used range, so selecting column A is fast.
- **Protected sheets** get a status bar message, never an error box.
- **AutoColor** leaves text, blank cells and cells already in another colour (like a red flag) alone.
- **The bps and BBSW formats** expect the cell to hold basis points (150 means 150bps), not a percentage.

## Problems and ideas

Please [open an issue](../../issues). Include your Excel version (Excel > About Excel) and what you pressed.

## Licence

[MIT](LICENSE). Free to use, change and share.

Not affiliated with or endorsed by Microsoft. Excel is a trademark of Microsoft Corporation.
