# Testing

## Automated: `test.ps1` (58 checks, 2026-11-10, Word 16.0.17932, Windows)

Runs every step in its own hidden Word and always closes it. Changes no Word or Windows settings.

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

Also look at: the forms at 100% and 150% display scaling, and a macro that shows its own form while the menu stays open.
