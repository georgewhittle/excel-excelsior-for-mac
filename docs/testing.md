# Testing checklist

Run this on a fresh build before tagging a release.

## Cycles
- [ ] Pressing a shortcut N times gives the Nth option, whatever the cell started with.
- [ ] One more press after the last option wraps to the first.
- [ ] Changing the selection, switching sheets or using a different shortcut starts again from the first option.
- [ ] Each format looks as in [cycles.md](cycles.md), with positive, negative, zero and text cells side by side.
- [ ] Zero cells show an en dash, not `?` or garbage.

## Increase / decrease decimals
- [ ] Works on every format in the number, percent and multiple cycles, keeping brackets, padding and suffixes.
- [ ] General → number cycle step 2. Decrease at zero decimals does nothing.

## Colours
- [ ] Blue-black flips both ways, based on the active cell.
- [ ] AutoColor on a test block: typed number, text, blank, same-sheet formula, `=A1*1.1`, other-sheet link, other-workbook link, error value, and a cell already coloured red. Each gets the right colour; the red cell, text and blank are untouched.

## Borders, alignment, indent
- [ ] Bottom border changes only the bottom edge. Outside border changes only the perimeter. Inside borders are untouched.
- [ ] Centre across selection works on several columns without merging.

## Selections and robustness
- [ ] Multi-area selections (Cmd-click) work for every feature.
- [ ] A whole column or row is clipped to the used range and finishes in under 2 seconds.
- [ ] Merged cells don't cause an error.
- [ ] A protected sheet gives a status bar message, not an error box.
- [ ] No feature ever shows a VBA error box.

## Undo
- [ ] Cmd+Z after every feature restores exactly what was there.
- [ ] Undo still restores the right cells after switching sheets in between.
- [ ] A selection over 50,000 cells gives the "undo unavailable" message and the action still runs.

## Loading
- [ ] Installing the downloaded release (not just a local build) works by following the README.
- [ ] Every shortcut works in a fresh Excel session, with no manual step.
- [ ] Unticking the add-in gives Excel's own shortcuts back.
