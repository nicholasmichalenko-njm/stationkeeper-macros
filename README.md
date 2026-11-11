# Stationkeeper Macros

A menu and a bookshelf for the macros in a Word template.

- **The menu** (`StationkeeperMacros`): search as you type, sort by category, A to Z, most used or recent, filter by what a macro affects or by its tie to Stationkeeper, read what it does, and run it. It can stay open beside the document, and shows which Stationkeeper effort folder it reads.
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

### Three kinds of macro

The menu is a general macro menu. Each macro is one of three kinds (the admin's **Stationkeeper** box; the Stationkeeper Tools come filled in):

| Kind | Badge | What Run does |
|---|---|---|
| Not related | none | Runs. Your own macros, and general tools such as Toggle Field Codes. |
| Built for Stationkeeper | ◇ SK | Runs on its own. If no effort folder is linked, it asks first: "This is built for Stationkeeper but is still functional. Would you like to run it?" with **Run it**, **Link a folder...**, **Cancel** and **Don't ask again**. |
| Needs Stationkeeper data | ◆ SK | Runs only when the linked effort has the data it needs (the effort rules, template or dictionary). If not, it says "This macro is used with the Stationkeeper application. Please point to the Stationkeeper data folder for your effort.", names the files it needs and where Stationkeeper writes them, says what is wrong right now, and offers **Point to the folder...**. |

**Show** filters the list: All macros, My macros, Stationkeeper tools, Needs effort data, and Ready to run now (what can run right now, given the selection and the link).

### The Stationkeeper link

The strip under the list shows where the Stationkeeper data comes from, with **Link a folder...** (or **Change...**), **Refresh** and **Unlink**.

- **One folder pick.** Pick the effort folder that Stationkeeper saves to, or the `Stationkeeper Data` folder inside it. It can be a local folder or a synced SharePoint or OneDrive folder. Each person links once on their own PC, and the link works for every template.
- **Read fresh every time.** The menu reads the folder each time it opens, before each run, and on **Refresh**. A save in Stationkeeper, or a colleague's save once OneDrive has synced it, shows up on the next run, with no reopening.
- **What the strip says:** Not linked; Folder missing; No data yet; Incomplete (Stationkeeper is mid-save or OneDrive is mid-sync, so try again in a minute); Current; Aging (not saved in Stationkeeper for 24 hours); Stale (72 hours). Each effort sets its own warning times.
- **Read only.** The macros never write to the effort folder; Stationkeeper is the only writer. The file layout is in the effort data contract (project doc `stationkeeper-effort-data-contract-v1.md`).
- **The tools read the effort's rules first (1.0.0).** Once linked, the tools that are built for Stationkeeper use the effort's rules, then the standards workbook for anything the effort does not set, then their own defaults:
  - A list the effort sets (banned phrases, placeholders, allowed styles, table types, fonts, acronyms and the rest) replaces that sheet of the workbook, even when the effort's list is empty.
  - Settings come over one at a time; a setting the effort leaves out (or sets to null) comes from the workbook.
  - The effort's template becomes StyleSource (Import Team Styles), and its dictionary becomes the Dictionaries list (Load Team Dictionaries).
  - Every report says where its rules came from: "Rules: effort <name> rev <n>, then <workbook>".
- **A local copy keeps working when the folder can't.** Each complete new revision is copied to this person's profile (`%LOCALAPPDATA%\Stationkeeper Tools\Effort cache`), one effort only. While Stationkeeper is mid-save, OneDrive is mid-sync, or the folder is out of reach, the tools use the last complete copy and say so on the report. Unlink, or linking another folder, removes the copy.
- **SharePoint and OneDrive checks (1.1.0).** The strip warns when:
  - the effort folder is synced but not set to **Always keep on this device**, so it would not work offline;
  - a file in `Stationkeeper Data` is online-only right now (it names the file);
  - OneDrive kept conflicting copies (`rules-PCNAME.json`, `manifest (1).json`) because two PCs saved the effort at once. The real files are still the ones used.
- **The link comes to you (1.1.0).** When the document you opened the menu from sits inside an effort folder (that folder or one above it holds `Stationkeeper Data`) and the menu reads another effort or none, the strip offers it: "This document is in Effort B, but the menu reads Effort A. Switch?" with **Use this folder**. Nothing changes until you click it.
- **Your own macros that need effort data** can call `SK_EffortFile("rules")` (or `"template"`, `"dictionary"`): it returns the file to read now (the live file, else the last complete copy), or "" when there is none. The template is always a local copy (`%LOCALAPPDATA%\Stationkeeper Tools\Attach`, one folder per effort and revision), so a document attached to it never holds Stationkeeper's own file open (1.2.0).

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

## Stationkeeper Tools (0.4.0 and later; 61 tools in 0.9.0)

Checking and fixing tools for Word documents, PowerPoint decks and Excel workbooks, in the same kit. They work on their own (View > Macros) and need nothing from the Stationkeeper app. With the menu installed, they show up in it already described.

**Add them:** open the kit, View > Macros > `StationkeeperToolsInstall`, pick the template. That copies fifteen modules (`SKW_Lib`, `SKW_Setup`, `SKW_Fields`, `SKW_Review`, `SKW_Styles`, `SKW_Tables`, `SKW_Layout`, `SKW_Captions`, `SKW_Lists`, `SKW_Acronyms`, `SKW_Text`, `SKW_Output`, `SKW_Utils`, `SKP_Deck`, `SKX_Book`). Running it again updates them; `StationkeeperToolsRemove` takes them out. Or import the files in `Import files/Stationkeeper Tools/`. Open the template once afterwards: the menu lists the tools with titles, descriptions and categories. Save once in the admin to keep that in the catalog.

**Team rules live in one Excel workbook, or come from the linked Stationkeeper effort (see The Stationkeeper link).** `MakeStandardsWorkbook` makes one with starting values (allowed styles, table types, placeholders, settings, plus sheets later tools use) and opens it to fill in. Each person runs `ChooseStandardsWorkbook` once per template to point the tools at the shared copy. Without a workbook the tools use sensible defaults and say so on every report.

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
| Section Map | Every section: first page, orientation, paper, margins, columns, how it starts, headers and footers linked or own, page number restarts | No, report |
| Banned Phrases Review | The team's banned phrases (Banned phrases sheet) in the selection or the whole body: replacements as tracked changes, a comment naming the rule on each | Yes, one Ctrl+Z |
| Caption Audit | Tables and figures without captions, typed caption numbers, wrong labels, captions on the wrong side (settings TableCaptionPosition, FigureCaptionPosition), repeated captions | No, report |
| Import Team Styles | Brings the style definitions in from the team template (setting StyleSource, else the attached template) and lists which changed | Yes, one Ctrl+Z |
| Remap Style | Changes every use of one style to another, or every pair on the Style map sheet | Yes, one Ctrl+Z |
| Keep-With-Next Audit | Headings and captions above tables that can be stranded at a page bottom, long kept chains | No, report |
| Break Hygiene | Manual page and column breaks, empty paragraphs used as spacing, line breaks ending paragraphs, body text starting a new page | No, report |
| Typed Lists Finder | Paragraphs starting with a typed bullet or number instead of a real list | No, report |
| Toggle Field Codes | Shows or hides field codes (Alt+F9) | View only |
| TOC Check | Headings missing from the contents, entries whose heading changed, page numbers out of date | No, report |
| Acronym Quick Check (text only) | Acronyms in the typed text: uses, first use in running text (not tables or headings), spelled out there as "Words (ACR)" or not, team meaning from the Acronyms sheet or the Stationkeeper app's export (setting AcronymList). Pictures are not scanned and the report says so; the app does the full check | No, report |
| Typography Sweep | Double spaces, space before punctuation, straight quotes, double or spaced hyphens, three dots, spaces at paragraph ends | No, report |
| Load Team Dictionaries | Adds the dictionaries on the Dictionaries sheet to Word on this PC | Word settings only |
| Check Selection Spelling | Word's spelling and grammar issues in the selection, with its first suggestion | No, report |
| Font Not Allowed | Every font used, flagging ones not on the Fonts sheet | No, report |
| Page Limit Counter | Pages per top-level part and in total against the Page limits sheet and PageLimitTotal | No, report |
| Graphics Audit | Pictures, charts and shapes: inline or floating, size, too wide, scaled up, no alt text, missing linked file | No, report |
| Metadata Clean Copy | Saves "<name> - clean.docx" beside the document without author, company, template path, variables and personal details | No, saves a copy |
| PDF with Bookmarks | Saves a tagged PDF with heading bookmarks, without file details | No, saves a PDF |
| Remove Unused Custom Styles | Deletes custom styles not applied anywhere, not on the team list and not depended on | Yes, one Ctrl+Z |
| Unlink / Lock / Unlock Fields (Selection) | Fields in the selection to plain text, locked, or unlocked (contents and indexes are never unlinked) | Yes, one Ctrl+Z |
| Restart Numbering / Continue Numbering | The list at the cursor starts again at 1, or continues the numbered list before it | Yes, one Ctrl+Z |
| Convert to Team List | Selected typed bullets and numbers become the team list styles (BulletStyle, NumberStyle); the typed markers go | Yes, one Ctrl+Z |
| Fit All Tables | Tables wider than the margins fit the page width (or only those in the selection) | Yes, one Ctrl+Z |
| Go to Next Finding | Selects the next finding of the last report in the document it checked | View only |
| Install Tool Shortcuts | Ctrl+Alt+Shift+N, U, C, R and L for the most used tools, leaving keys already in use alone | Template |
| Insert Acronym List | An Acronym and Definition table at the cursor, A to Z, from the team meanings (only those used, unless AcronymListUsedOnly is No) | Yes, one Ctrl+Z |
| Reset Paragraph to Style | The paragraph at the cursor or the selected ones back to their style, emphasis included | Yes, one Ctrl+Z |

**PowerPoint (0.8.0).** These check the deck open in PowerPoint. They never start PowerPoint or open a file on their own, and their reports open in Word.

| Tool | What it does | Changes the deck? |
|---|---|---|
| Deck Leftovers | Speaker notes, comments, hidden slides, placeholder text | No, report |
| Deck Banned Phrases | Banned phrases in slides and notes, with replacements | No, report |
| Deck Acronym Check (text only) | Acronyms, first slide, spelled out or not, team meaning | No, report |
| Deck Font Audit | Fonts used, flagged against the Fonts sheet | No, report |
| Deck Small Text | Text under DeckMinFontSize (10 pt) | No, report |
| Deck Alt Text Audit | Graphics without alt text (Decorative ones pass) | No, report |
| Deck Off-Slide Objects | Shapes partly or wholly off the slide | No, report |
| Deck Titles Check | Slides without a title, slides sharing one | No, report |
| Deck Empty Placeholders | Empty placeholders left on slides | No, report |
| Deck Delivery Copy | "<name> - delivery.pptx" without notes, comments, hidden slides or personal details | No, saves a copy |

**Excel (0.9.0).** These check the workbook open in Excel (the active one, or the only one open). They never start Excel or open a file on their own, and their reports open in Word. On very large sheets they read the first 20,000 text cells of each sheet.

| Tool | What it does | Changes the workbook? |
|---|---|---|
| Book Formula Errors | Cells showing #DIV/0!, #N/A, #REF!, #VALUE! and the rest, with the formula | No, report |
| Book Broken Names | Defined names that point at #REF!, at nothing, or at another file | No, report |
| Book External Links | Other workbooks this one links to, and the cells that use them | No, report |
| Book Hidden Content | Hidden and very hidden sheets, hidden rows and columns, hidden names | No, report |
| Book Leftovers | Notes, comments and placeholder text (TBD, XXX, [insert) | No, report |
| Book Banned Phrases | Banned phrases in cell text, with replacements | No, report |
| Book Acronym Check (text only) | Acronyms, first cell, spelled out or not, team meaning | No, report |
| Book Font Audit | Fonts used, flagged against the Fonts sheet | No, report |
| Book Merged Cells | Merged areas (they break sorting, filtering and screen readers) | No, report |
| Book Delivery Copy | "<name> - delivery.xlsx" without notes, comments, links to other workbooks (kept as values) or personal details. Hidden sheets stay; Book Hidden Content lists them | No, saves a copy |

How they behave:
- **Reports** open as a new Word document with one table and the standards used at the top. They change nothing in your document, not even its Undo list (the report is written as a web page and opened, so Word edits nothing while making it).
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
- **The Stationkeeper link** (the effort folder): their Windows profile (`...\Stationkeeper Tools\Link`), one per person per PC.
- **The local copy of the linked effort's data:** `%LOCALAPPDATA%\Stationkeeper Tools\Effort cache`, on that PC only, one effort, removed on Unlink.

## Build and test (developers)

Needs Windows, Word, and "Trust access to the VBA project object model" on.

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1   # makes build\ (kit + import files), compiles first
powershell -ExecutionPolicy Bypass -File test.ps1    # 187 checks on a sample template; changes no settings
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

- 1.2.0, shared findings; 1.3.0, both directions (a hand-filled setup workbook the app imports).

- Library mode: add or remove whole macro modules from a master Macro Library template, and export one macro family as its own small template.
- Load Nicholas's ProposalEngine macros as `SKM_` modules with their catalog filled.
- Stationkeeper app: show "Macros in this template" in the template library.
