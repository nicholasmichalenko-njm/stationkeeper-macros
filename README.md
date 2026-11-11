# Stationkeeper Macros

A menu and a bookshelf for the macros in a Word template.

- **The menu** (`StationkeeperMacros`): search as you type, sort by category, A to Z, most used or recent, filter by what a macro affects, read what it does, and run it. It can stay open beside the document.
- **The admin** (`StationkeeperMacrosAdmin`): see every macro in the template, then give each one a title, description, category, keywords, what each one affects and whether it changes the document, and choose which ones show in the menu. Set and order your own categories, export the catalog to Excel and back, and check for macros that were added, moved or removed.

It works with whatever macros are already in the template: one module or many. It never edits your code. Part of the Stationkeeper suite and separate from the Stationkeeper app.

## What is in a release

| File | What it is |
|---|---|
| `Stationkeeper Macros Kit.dotm` | The kit, plus installers that copy the menu and admin, and the Stationkeeper Tools, into any template |
| `Import files/` | The same three pieces as files (`SK_Macros.bas`, `frmSKMenu.frm/.frx`, `frmSKAdmin.frm/.frx`) for importing in the VBA editor, and `Stationkeeper Tools/` with the tool modules |
| `Read me first.txt` | Short version of this page |

## Add it to a template

**Work on a copy of the template first.** Both ways make a dated backup anyway.

**Way 1, no VBA editor (Word's Organizer):**
1. Save `Stationkeeper Macros Kit.dotm` somewhere Word trusts for macros. If it came from email, Teams or the web, right-click it > Properties > tick **Unblock** > OK first.
2. In Word: File > Open the kit. Enable macros if asked.
3. View > Macros > `StationkeeperMacrosInstall` > Run. Pick your template (it must be closed, or open and saved).
4. The kit adds three items to the template: module `SK_Macros` and forms `frmSKMenu` and `frmSKAdmin`. Running it again updates them. `StationkeeperMacrosRemove` takes them back out.

**Way 2, VBA editor:** open the template, Alt+F11, File > Import File, and import the three files from `Import files/`. Keep each `.frx` beside its `.frm`. Save the template.

## First time in a template (the admin)

1. Open the template itself (File > Open the .dotm).
2. View > Macros > `StationkeeperMacrosAdmin` > Run.
3. On **Macros**, the list has three columns: the macro, **In menu** (Shown, Hidden, Not set, or Macro gone) and **Description** (Written or Missing). Use **Show** to list only the ones still missing a description or not yet in the menu. Pick each macro and fill in Title, Description, Category, Keywords, **What this macro affects** and **Does this change the document?** Tick **Show in the menu** for the ones people should see. **Use code comment** copies the comment above the macro into the description.
4. **Bulk tools** work on the whole list at once: **Fill blank descriptions from code comments**, **Show every described macro in the menu**, **Remove macros that are gone**, export and import for Excel, and **Read macros from .bas files**. Each button has a sentence beside it saying what it does.
5. Faster for many macros: keep a **primary list** (see below).
6. **Categories** sets your own categories and their order.
7. **Settings** sets the menu title, the default sort, two ways to open the menu (**Alt+Shift+M** and a button on the **Add-ins** tab), and **where backups go**.
8. Tick **Ready to save**, then **Save to template**. The first time, the admin asks where backups should go (a folder beside the template is suggested). You can change it on **Settings**. Save stays off until there is a change and the box is ticked. Each save makes a dated backup of the template first, adds a line to **History**, and clears the tick.

## Using the menu

Run `StationkeeperMacros` (or press Alt+Shift+M, or use the Add-ins tab button) in any document based on the template.

- Type to search. The list narrows with every letter and says where each match came from (title, keywords, what it affects, category, description).
- Up and Down move, Enter runs, Esc closes.
- **One Ctrl+Z undoes a whole macro.** Each run is one step on Word's Undo list, named "Stationkeeper: <title>".
- **The menu checks what a macro needs first.** A macro set to work on a selection, a table at the cursor or a list at the cursor will not run until that is true, and the menu says what to do.
- Pins, Most used and Recent belong to each person on their own PC and are never saved in the template.

## The primary list: update many macros at once

Keep one list with the latest titles, categories, descriptions and the rest, and update any template from it in one step.

1. **Bulk tools > Export the list to Excel** makes a workbook (sheet "Macros") with one row per macro. Say Yes to use it as the primary list, or move it somewhere shared first (OneDrive or SharePoint) and choose it on **Settings > Choose the primary list**.
2. Edit the list in Excel. Add rows for macros other templates have: rows for macros a template doesn't contain are skipped and listed, so one list can serve many templates.
3. **Bulk tools > Update from the primary list** shows what will change (how many titles, categories, descriptions and so on, which rows were skipped, which macros aren't on the list) and changes nothing until you say Yes. Then tick **Ready to save** and save.

**Update from a list file** does the same from any file: Excel (.xlsx, .xlsm, .xls), CSV, a **table in a Word document**, or **JSON**.

| Column | Notes |
|---|---|
| Macro | Required. `Module.Macro`, or just the macro name if no other module has one with the same name. |
| Title, Category, Description, Keywords | Text. A new category is added to the category list. |
| What it affects | Comma-separated. "Affects" also works. |
| Works on | For example Selection, Table at the cursor, Whole document. |
| Does it change the document | For example Yes; No, it only reports. |
| Show in menu | Yes or No (also True/False, 1/0). |

- **Blank cells leave values as they are.** Clear a value in the admin itself.
- Other columns (Status, Test, your own notes) are ignored. In Excel the sheet named "Macros" is read (otherwise the first sheet), and the header can sit below a title row.
- In Word, the first table whose first row has a Macro column is read (keep it a simple table, no merged cells).
- JSON: a list of objects, or `{"macros": [ ... ]}`, with the same names as keys (`"Show in menu": true` works).

## Testing macros

- **Test column** in the admin's list: Passed, Failed or Not tested (on this PC). **Show > Test failed / Not tested** lists them.
- **Test this macro** (Macros tab) runs it on a scratch document made from the template, or on a fresh copy of a sample document you pick (**Bulk tools > Choose a test document**). Your real files and the sample are never changed.
- **Test every macro** (Bulk tools) runs them one after another. Macros whose code saves, closes or deletes files, or ends Word, are skipped (test them one at a time).
- If a macro hits an error it doesn't handle, Word shows its own error dialog: click **End**. That macro is then marked **Failed**. The same happens when a macro from the menu stops with an error.
- **Check the code health** (Bulk tools) runs nothing. It finds missing reference libraries (the usual cause of "works on my PC, not hers") and notes risky lines in each macro: saves, closes or deletes files; a fixed file path; Word alerts or screen updating switched off and never back on; On Error Resume Next (a test can't see errors the macro hides); a Stop line. To check the code compiles, use Debug > Compile in the VBA editor.
- Limit: a test shows the macro runs without an error. It doesn't check the result is right.

## Stationkeeper Tools (0.4.0 and later)

Checking and fixing tools for Word documents, in the same kit. They work on their own (View > Macros) and need nothing from the Stationkeeper app. With the menu installed, they show up in it already described.

**Add them:** open the kit, View > Macros > `StationkeeperToolsInstall`, pick the template. That copies six modules (`SKW_Lib`, `SKW_Setup`, `SKW_Fields`, `SKW_Review`, `SKW_Styles`, `SKW_Tables`). Running it again updates them; `StationkeeperToolsRemove` takes them out. Or import the files in `Import files/Stationkeeper Tools/`. Open the template once afterwards: the menu lists the tools with titles, descriptions and categories. Save once in the admin to keep that in the catalog.

**Team rules live in one Excel workbook.** `MakeStandardsWorkbook` makes one with starting values (allowed styles, table types, placeholders, settings, plus sheets later tools use) and opens it to fill in. Each person runs `ChooseStandardsWorkbook` once per template to point the tools at the shared copy. Without a workbook the tools use sensible defaults and say so on every report.

| Tool | What it does | Changes the document? |
|---|---|---|
| Safe Update All Fields | Updates every field in every part (body, headers, footers, notes, text boxes), then contents and figure tables. Fields that pull in outside files (LINK, INCLUDETEXT, INCLUDEPICTURE) are kept as they are. Track changes is paused for the update and put back. | Yes, one Ctrl+Z |
| Find Broken References | Cross-references to bookmarks that are gone, fields showing Error!, internal links that go nowhere, included files that are missing | No, report |
| Field Code Report | Every field: type, code, result, where, page, locked | No, report |
| Pre-Delivery Leftovers | Comments, tracked changes, hidden text, highlighting, placeholders (TBD, XXX, [insert...), author and company saved in the file | No, report |
| Style Inspector | The paragraph and character style at the cursor and anything formatted by hand on top; several paragraphs make a report | No |
| Non-Allowed Style Report | Styles in use that are not on the team list, required styles that are missing. Without a list: custom styles | No, report |
| Direct Formatting Report | Paragraphs whose font or spacing was set by hand. Bold, italic, superscript and subscript count as emphasis (setting `KeepOnClear`). In tables only font and size are checked, since table styles set the rest | No, report |
| Clear Direct Formatting (Selection) | Clears hand formatting in the selection; keeps emphasis and list numbering | Yes, one Ctrl+Z |
| Table Standards Audit | Each table against the team table types: style, repeating header row, rows breaking across pages, width, floating or nested tables, cell styles, a caption above or below | No, report |
| Apply Team Table Type | Sets the table at the cursor to a team type: style, header row, rows kept whole, width, cell styles, padding | Yes, one Ctrl+Z |

How they behave:
- **Reports** open as a new Word document with one table (Rule, Finding, Where, Page, Text) and the standards used at the top. They change nothing in your document, not even its Undo list.
- **Fixes** are one step on Undo, named "Stationkeeper: ...".
- **Co-authored files:** on a SharePoint or OneDrive file with AutoSave on, the fixes do not run (to protect co-authors' edits); the reports still do.
- **Test every macro** in the admin runs the report tools quietly and closes their reports; the setup tools wait for a person.

## "Trust access to the VBA project object model"

Word blocks a template from reading its own code unless this setting is on (File > Options > Trust Center > Trust Center Settings > Macro Settings). It is off by default.

- **The menu never needs it.** It runs from the saved catalog.
- **The admin is best with it on, for the person running the admin only.** With it on, the admin lists every macro live, reads code comments, and spots macros that were added, moved or removed.
- **Without it,** the admin says so, opens on **Bulk tools**, and offers **Read macros from .bas files**: export your module(s) from the VBA editor (right-click > Export File) and pick the file(s).

## What counts as a macro

A `Public Sub` (or plain `Sub`) with no arguments, in a standard module. These are left out: `Private` subs, Functions, subs that take arguments, AutoOpen and the other Auto macros, modules marked `Option Private Module`, and the kit itself (`SK_*`, `frmSK*`). That is the same rule Word uses for its Macros list.

## Where things are kept

- **The catalog** (titles, descriptions, categories, keywords, what it affects, show or hide, history): a hidden XML part inside the template, namespace `urn:stationkeeper:macros:1`. Changing it never touches code, so a code signature stays valid.
- **Backups, a readable copy of the catalog, and Excel exports:** the folder each person chooses the first time they save or export (kept in their Windows profile, for every template).
- **Each person's pins, counts and preferences:** their Windows profile (`HKCU\Software\VB and VBA Program Settings\Stationkeeper Macros`).

## Build and test (developers)

Needs Windows, Word, and "Trust access to the VBA project object model" on.

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1   # makes build\ (kit + import files), compiles first
powershell -ExecutionPolicy Bypass -File test.ps1    # 81 checks on a sample template; changes no settings
```

`test.ps1` builds a sample template with all of its macros in one module (plus one in a second module), installs the kit, and checks scan, catalog save and reload, menu search, admin, CSV round trip, the shortcut and button, use from a document based on the template, macros that move between modules, updating the kit, and the no-access path; then installs the tools and runs each one on a document built with one of every defect. See `TESTING.md` for the checks run by hand.

## Layout

```
src/            SK_Macros.bas (engine + entry macros), SK_Install.bas (kit only),
                frmSKMenu.code.vba, frmSKAdmin.code.vba (forms build their controls in code)
src/tools/      SKW_*.bas: Stationkeeper Tools (SKW_Lib is the shared plumbing)
test/           sample macros and the self-test module
tools/          wordjob.ps1 (runs one hidden Word and always closes it), watch.ps1 (reads stuck dialogs)
build.ps1, test.ps1
```

## Next

- Stationkeeper Tools, ten at a time, a release for each: captions, banned phrases, pagination and lists (0.5.0); acronym quick check, typography, dictionaries, fonts, PDF (0.6.0); utilities (0.7.0); PowerPoint (0.8.0); Excel (0.9.0).

- Library mode: add or remove whole macro modules from a master Macro Library template, and export one macro family as its own small template.
- Load Nicholas's ProposalEngine macros as `SKM_` modules with their catalog filled.
- Stationkeeper app: show "Macros in this template" in the template library.
