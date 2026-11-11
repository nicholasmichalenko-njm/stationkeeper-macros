# Testing

## Automated: `test.ps1` (177 checks, all passed in the 1.0.0 run; 2026-11-10, Word 16.0.17932, Windows)

Runs every step in its own hidden Word and always closes it. Changes no Word or Windows settings. A guard looks for a dialog every 15 seconds; a dialog still open at the next look is written into the results and that Word is ended, so a VBA error never stalls the run.

| Area | Checks |
|---|---|
| Install | Kit copied into a template with Word's Organizer; installing again updates it and keeps the catalog |
| Scan | 8 macros found in 2 modules; Private subs, subs with arguments, Functions, AutoOpen and the kit are left out; comments above and inside a Sub are read; separator lines skipped; readable titles (`TOCFix` becomes "TOC Fix") |
| Catalog | Saves into the template with a dated backup; survives close and reopen; saves from a document based on the template; a macro moved to another module keeps its description |
| Menu | Lists only macros put in the menu; search by keyword prefix, description and affects; no-match message; runs a macro by Module.Macro |
| Admin | Tabs are Macros, Categories, Settings, Bulk tools, History; labels read "What this macro affects" and "Does this change the document?"; no Also in Stationkeeper box; opens, edits, counts pending changes; list shows In menu and Description columns; Save is off with nothing to save and needs Ready to save ticked |
| CSV | Export and update round trip with quotes, commas and line breaks |
| Lists | Excel export read back; a primary list edited in Excel (changed description and category, blank title left alone, Show in menu Yes, a row for a macro not in the template skipped, preview counts); a Word table; JSON with escapes, unicode, true and a bare macro name; the primary list shown in the admin |
| Opening the menu | Alt+Shift+M and the Add-ins tab button |
| Backups | Folder chosen per person and kept; saves back up into it |
| Health | Missing libraries; notes for saves/closes/deletes files, hides its own errors, alerts off, fixed path; clean macros get no notes |
| Test runs | Passes on a scratch document and closes it; a missing macro fails; a run that never finished is recorded as Failed; Test all runs 12 safe macros and skips the file-changing one |
| Undo | One Ctrl+Z undoes a two-edit macro run from the menu |
| Needs | Table at the cursor and Selection checked before running |
| Stationkeeper Tools: install | Six SKW_ modules copied with the Organizer; the 12 tools arrive in the catalog described, categorised, with the right "Works on", and in the menu; menu search finds them; the template's own catalog is untouched |
| Standards workbook | Made in Excel with 25 allowed styles, 2 table types, 10 placeholders and the settings, then read back; reports name it |
| Defect document | Built in VBA with direct font, size and spacing, a custom style, a good and a broken REF, an INCLUDETEXT to a missing file, a link to a missing bookmark, a comment, a tracked insertion, hidden text, a highlight, two placeholders, a captioned good table, a long wide table with no caption, a floating table, a TOC and a PAGE field |
| Reports | Broken references (3, good REF not listed); field code report by type; leftovers (comment, tracked change, hidden, highlight, 2 placeholders); style inspector (font and size by hand); non-allowed styles with and without a workbook; direct formatting (bold and subscript treated as emphasis); table audit (good table clean, 6 findings on the bad one, floating table). Running every report leaves the document's Undo list exactly as it was |
| Fixes | Clear direct formatting: back to the style, bold kept, one Ctrl+Z restores; apply table type: style, repeating header, rows whole, 100% width, then one Ctrl+Z (pressed from outside the macro, as a person would) restores it; safe update: the included file is not touched, the broken REF now shows its error, track changes put back |
| Batch 2 (0.5.0) | A second defect set: banned phrases, typed bullet/number/letter beside a real list, three empty paragraphs, a line break, a page break, a heading without keep-with-next, two headings added after the contents, two pictures (one captioned with SEQ, one with a typed number), a landscape section with its own footer. Checks: all 13 reports leave Undo as it was; section map; caption audit (2 missing, 1 typed, the figure caption below not taken for the next figure's); keep-with-next; break hygiene (section break not counted as a page break); typed lists; TOC check (2 missing); field codes toggled and back; banned phrases (2 tracked replacements, 1 flag, 3 comments, tracking put back, one Ctrl+Z removes it all); remap style (one Ctrl+Z restores); import team styles (Normal restored from the template; one Ctrl+Z puts the changed size back) |
| Batch 3 (0.6.0) | A third defect set, saved as a .docx with an author and company: RFP spelled out at first use and on the team list, SOW used before it is spelled out, typing slips, two misspellings, two unused custom styles, alt text on one picture. Checks: all 18 reports leave Undo as it was; acronym quick check says text only and gets RFP and SOW right; typography sweep finds each kind of slip; spelling lists the misspellings with suggestions; empty dictionary list explained; fonts not allowed (Arial, Courier New); page limit parts; graphics (one picture without alt text); the clean copy has no author or company and the open document stays active; the PDF is saved; unused custom styles go, Fancy Body stays, one Ctrl+Z brings them back |
| Utilities (0.7.0) | Two numbered lists: continue makes B1 3, restart at B2 makes it 1; lock, unlock and unlink the two REF fields (one Ctrl+Z brings them back); typed bullet, number and letter become List Bullet and List Number with markers gone; the 9-inch table fits 6.5 inches and one Ctrl+Z restores it; Go to Next Finding walks the table audit; Insert Acronym List adds RFP with its meaning and one Ctrl+Z removes it; Reset Paragraph to Style removes Arial and bold; shortcuts set in the template |
| PowerPoint (0.8.0) | A four-slide windowless test deck (notes, TBD, utilize, RFP spelled out, a duplicate title, an empty body, a box parked off the slide, 8 pt Arial, a picture without alt text, a hidden slide) handed to the tools directly. Each report finds exactly its defect; the delivery copy has 3 slides and no notes while the open deck keeps 4. Test every macro skips the deck tools, so it never touches a person's PowerPoint |
| Excel (0.9.0) | A three-sheet windowless test workbook (#DIV/0! and #N/A, a name pointing at #REF!, a hidden sheet, a hidden row, a note, TBD, utilize, RFP spelled out, Arial, a merged A10:B10) handed to the tools directly. Each report finds exactly its defect; the delivery copy has no notes while the open workbook keeps its note. Test every macro skips the workbook tools, so it never touches a person's Excel |
| Stationkeeper kinds and link (0.10.0) | Four menu macros marked built, needs rules, needs the template and not related; the kinds survive a reload; the tools come filled in (Style Inspector built, Toggle Field Codes not related). Each Show filter lists the right macros with the right badge. Not linked: the built macro asks gently, with Run it, Link a folder, Cancel and Don't ask again; the needs-data macro names the file and where Stationkeeper writes it and offers the picker; a not-related macro just runs; Test this macro skips the needs-data one. Don't ask again stops the question. A stand-in effort folder goes through every state: Folder missing, No data yet, Current (picking the effort folder or its Stationkeeper Data folder), Incomplete (rules.json a revision behind), Stale (5 days), Unreadable (not Stationkeeper's manifest). Linked: the needs-rules macro runs and the needs-template one says the template is missing. The person's own link is put back after the run |
| Effort rules (1.0.0) | Not linked, the tools read the standards workbook as before. A stand-in effort (rules, template, dictionary) is linked: the report stamp names the effort and revision; effort settings win one by one (CaptionStyle) while unset ones (TableLabel) come from the workbook; an effort list replaces the sheet, even an empty one (placeholders 0); a list the effort leaves out (acronyms) stays from the workbook; StyleSource and the dictionary point into the local copy. Banned Phrases Review catches the effort's phrase and not the workbook's. Mid-save (rules.json a revision behind), the last complete copy is used and the menu says so; a new revision is picked up on the next run and the old revision folder is removed; with the folder renamed away, the copy is used; Unlink removes the copy. The test clears this PC's local copy; the tools make it again on the next run |
| Co-authoring | With AutoSave on a cloud file (test switch), fixes are refused and reports still run |
| Tools in the menu engine | A report tool runs through the menu's run (one Undo record open) and makes its report; Test every macro passes all 22 safe macros, skips the two that need a selection or table, closes every report and clears the quiet flag |
| No VBA access | Catalog trusted (nothing marked missing); menu lists, searches and runs; admin opens; .bas files read (one module or many) with their comments; the opening message explains the two ways forward |

The no-access checks use a test switch (`SK_ForceNoAccess`). Word started by automation can always read VBA projects whatever the Trust Center says, so a script cannot reproduce the real setting.

## By hand, before giving it to someone

On a PC where "Trust access to the VBA project object model" is **off** (Word's default):

1. Install the kit into a copy of the template (Way 1 in the README).
2. Open the template. Run `StationkeeperMacrosAdmin`. Expect a message with the two ways forward, then the Bulk tools tab.
3. Use **Read macros from .bas files** on an exported module. Describe two macros, show them in the menu, tick Ready to save, save. Check the In menu and Description columns update as you go.
4. New document from the template. Alt+Shift+M (after setting it on the Settings tab) or View > Macros > `StationkeeperMacros`. Search, run one macro, press Ctrl+Z once. Run a table macro with the cursor outside a table: the menu says what to do.
6. In the admin: Test this macro on one that fails on purpose (for example 1/0). Click End in Word's dialog, reopen the admin: it shows Failed.
5. Turn the setting **on**, open the admin again: every macro is listed live, comments are read.

Stationkeeper Tools by hand:

1. Run `StationkeeperToolsInstall` from the kit on a copy of a real template, then `MakeStandardsWorkbook`. Fill in a few of your styles and one table type; save.
2. On a real proposal: Pre-Delivery Leftovers, Find Broken References, Table Standards Audit. Check a few findings against the page numbers.
3. Select some hand-formatted text, Clear Direct Formatting (Selection), press Ctrl+Z once.
4. Put the cursor in a table, Apply Team Table Type with two types in the workbook: it asks which.

Also look at: the forms at 100% and 150% display scaling, and a macro that shows its own form while the menu stays open.
