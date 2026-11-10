# Testing

## Automated: `test.ps1` (36 checks, 2026-11-10, Word 16.0.17932, Windows)

Runs every step in its own hidden Word and always closes it. Changes no Word or Windows settings.

| Area | Checks |
|---|---|
| Install | Kit copied into a template with Word's Organizer; installing again updates it and keeps the catalog |
| Scan | 8 macros found in 2 modules; Private subs, subs with arguments, Functions, AutoOpen and the kit are left out; comments above and inside a Sub are read; separator lines skipped; readable titles (`TOCFix` becomes "TOC Fix") |
| Catalog | Saves into the template with a dated backup; survives close and reopen; saves from a document based on the template; a macro moved to another module keeps its description |
| Menu | Lists only macros put in the menu; search by keyword prefix, description and affects; no-match message; runs a macro by Module.Macro |
| Admin | Opens, edits, counts pending changes; list shows In menu and Description columns; Save is off with nothing to save and needs Ready to save ticked |
| CSV | Export and import round trip with quotes, commas and line breaks |
| Opening the menu | Alt+Shift+M and the Add-ins tab button |
| No VBA access | Catalog trusted (nothing marked missing); menu lists, searches and runs; admin opens; .bas files read (one module or many) with their comments; Check tab explains the two ways forward |

The no-access checks use a test switch (`SK_ForceNoAccess`). Word started by automation can always read VBA projects whatever the Trust Center says, so a script cannot reproduce the real setting.

## By hand, before giving it to someone

On a PC where "Trust access to the VBA project object model" is **off** (Word's default):

1. Install the kit into a copy of the template (Way 1 in the README).
2. Open the template. Run `StationkeeperMacrosAdmin`. Expect a message, then the Check tab with the two ways forward.
3. Use **Read macros from .bas files** on an exported module. Describe two macros, show them in the menu, tick Ready to save, save. Check the In menu and Description columns update as you go.
4. New document from the template. Alt+Shift+M (after setting it on the Look tab) or View > Macros > `StationkeeperMacros`. Search, run one macro.
5. Turn the setting **on**, open the admin again: every macro is listed live, comments are read.

Also look at: the forms at 100% and 150% display scaling, and a macro that shows its own form while the menu stays open.
