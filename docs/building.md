# Building from source

You only need this if you've changed the code. To just use the add-in, download it from the [releases page](../../../releases/latest).

The VBA source is plain text in `src/`. A small builder workbook reads those files and assembles `ExcelExcelsiorForMac.xlam`. The `.xlam` and the builder workbook are build outputs and aren't committed.

Keep the repo in a local folder, not OneDrive, iCloud Drive or Dropbox: Excel's sandbox sees cloud folders as web addresses and the build fails.

## One-time setup
1. In Excel: **Tools > Macro > Security**, tick **Trust access to the VBA project object model**. (Only the builder needs this. You can untick it after building.)
2. **File > New Blank Workbook**.
3. **Tools > Macro > Visual Basic Editor**. **File > Import File**, choose `build/import.bas`. Close the editor.
   If the import looks garbled, instead use **Insert > Module** and paste in the contents of `build/import.bas`, leaving out the first `Attribute` line.
4. **File > Save As**, format **Excel Macro-Enabled Workbook (.xlsm)**, name it `Builder.xlsm`, in this repo's `build` folder.

## Every build
1. Open `build/Builder.xlsm` and enable macros if asked.
2. **Tools > Macro > Macros**, pick **BuildAddin**, click Run.
3. The first time, Excel asks for access to the files in `src`: click Select or Grant.
4. It builds `ExcelExcelsiorForMac.xlam` into Excel's Add-ins folder, loads it, and says "Built and loaded".
5. Check it compiles: in the VBA editor, click **VBAProject (ExcelExcelsiorForMac.xlam)**, then **Debug > Compile VBAProject**. Nothing happening means it's fine.

If `build/import.bas` changes, re-import it into `Builder.xlsm`: in the VBA editor, right-click **modBuild > Remove** (don't export), then File > Import File and save.

## Adding a module
Add the new file to `SourceFiles()` in `build/import.bas`. That file keeps Windows (CRLF) line endings, so edit it with an editor that preserves them.

## How the code is laid out
```
src/
  modConfig.bas      every editable format, colour and shortcut
  modMain.bas        registers and removes the shortcuts
  modCycle.bas       the cycle engine (which step comes next)
  modNumbers.bas     number, percent, multiple and date cycles; decimals
  modColour.bas      blue-black, font and fill cycles, AutoColor
  modBorders.bas     bottom and outside border cycles
  modAlign.bas       alignment and indent cycles
  modUndo.bas        one-level undo via Application.OnUndo
  modUtil.bas        selection handling and status bar messages
  clsAppEvents.cls   resets the cycle when the selection or sheet changes
  ThisWorkbook.cls   registers shortcuts when the add-in opens
build/import.bas     the builder (module modBuild)
```

## Excel for Mac lessons
Each of these broke a build at some point:
- Register shortcuts from `Workbook_Open`, not `Auto_Open`, which doesn't run reliably for add-ins loaded at startup.
- Add-in shortcuts only work with Control (`^`) and Control+Shift (`^+`). Cmd (`*`) and Option (`%`) are ignored.
- The Mac VBA editor mangles non-ASCII characters. Build them in code, e.g. `ChrW(8211)` for the en dash.
- The builder adds code as text with LF line endings. The editor's own importer is unreliable on Mac, and CRLF double-spaces every line, which breaks ` _` line continuations.
- Every `Const` and module-level variable must come before the first procedure in a module.
- VBA only compiles a module when something in it first runs, so always run Debug > Compile after a build.
- Platform differences go in `#If Mac Then` blocks rather than separate modules.

## Making a release
1. Build as above and run through [testing.md](testing.md).
2. Copy `ExcelExcelsiorForMac.xlam` out of the Add-ins folder.
3. Update [CHANGELOG.md](../CHANGELOG.md), commit, and tag the version (e.g. `v0.9.0`).
4. On GitHub, create a release from the tag and attach the `.xlam`.
