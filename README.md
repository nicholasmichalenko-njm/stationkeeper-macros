# Stationkeeper Macros

A menu and a bookshelf for the macros in a Word template.

- **The menu** (`StationkeeperMacros`): search as you type, sort by category, A to Z, most used or recent, filter by what a macro affects or by its tie to Stationkeeper, read what it does, and run it. It can stay open beside the document, and shows which Stationkeeper effort folder it reads.
- **The admin** (`StationkeeperMacrosAdmin`): see every macro in the template, then give each one a title, description, category, keywords, what each one affects and whether it changes the document, and choose which ones show in the menu. Set and order your own categories, export the catalog to Excel and back, and check for macros that were added, moved or removed.

It works with whatever macros are already in the template: one module or many. It never edits your code. Part of the Stationkeeper suite and separate from the Stationkeeper app.

## What is in a release

| File | What it is |
|---|---|
| `Stationkeeper Macros Kit.dotm` | The kit, plus installers that copy the menu and admin, and the Stationkeeper Tools, into any template |
| `Stationkeeper Ribbon.dotm` | Optional (1.6.0): a small add-in that puts a **Stationkeeper** tab on Word's ribbon. Leave it out where add-ins are not allowed |
| `Import files/` | The same pieces as files (`SK_Macros.bas`, `frmSKMenu.frm/.frx`, `frmSKAdmin.frm/.frx`, `frmSKBrand.frm/.frx`) for importing in the VBA editor, and `Stationkeeper Tools/` with the tool modules and forms |
| `Read me first.txt` | Short version of this page |

## Add it to a template

**Work on a copy of the template first.** Both ways make a dated backup anyway.

**Way 1, no VBA editor (Word's Organizer):**
1. Save `Stationkeeper Macros Kit.dotm` somewhere Word trusts for macros. If it came from email, Teams or the web, right-click it > Properties > tick **Unblock** > OK first.
2. In Word: File > Open the kit. Enable macros if asked.
3. View > Macros > `StationkeeperMacrosInstall` > Run. Pick your template (it must be closed, or open and saved).
4. The kit adds four items to the template: module `SK_Macros` and forms `frmSKMenu`, `frmSKAdmin` and `frmSKBrand` (the Stationkeeper look). Running it again updates them. `StationkeeperMacrosRemove` takes them back out.

The kit's own menu (open the kit, then `StationkeeperMacros`) lists the installers and **Add Macros From Template**, described, under the category **Install and add**.

**Way 2, VBA editor:** open the template, Alt+F11, File > Import File, and import the four files from `Import files/`. Keep each `.frx` beside its `.frm`. Save the template.

## The Stationkeeper look and the ribbon tab (1.6.0)

**Finding a macro (1.16.0).** One row of filters:

| Filter | Choices |
|---|---|
| Search | Type any words: name, description, category, what it affects |
| Application | All, Word, Excel, PowerPoint |
| Category | The categories set in the admin |
| Sort by | Category, Name, Most used, Recent |

The "What it affects" filter is gone; Search still finds those words. The Show filter (Stationkeeper or Other) is gone; the diamond marks in the list still say which macros are made for Stationkeeper.

**The standards the tools read (1.16.0).** Two lines at the bottom of the menu, each with its own **Choose...**:

- **Global Standards:** the app's `Global Standards.xlsx`. Always needed.
- **A workbook says what it is (1.22.0).** Its Read me tab may carry the line **This workbook is** with Global Standards, Effort Standards or Program Standards, and **It belongs to** with a name. The tools go by that line, not by the file's name, so a workbook can be renamed or copied. A workbook that says it is an effort's or a program's is refused as the Global Standards, and the other way round. A program's workbook is chosen on the same line, which then reads "This program's standards". A workbook with no such line is read as before. Nothing is hidden in the workbook.
- **Found by its folder (1.23.0).** With no effort's workbook pointed to, the tools look in the folder of the document in front, then in up to three folders above it, for a workbook that says Effort Standards or Program Standards (or carries the older name Effort Formatting). The first one found is laid over the Global Standards; every report says it was found by the document's folder, and the menu's line names it. Pointing to a workbook always wins. Two such workbooks in one folder: none is used, and the report says so. A document that is not saved gets no search. The search never looks in a drive's top folder, a share's top folder, or your own profile folder. What each workbook says is kept on the PC by its date and size, so Excel is opened only for a new or changed workbook. The Global Standards are still chosen by pointing.
- **Plain column names are read too (1.24.0).** A workbook may name a column the older way (Align, Fmt, Edges: top) or the new, plain way (Alignment, Number format, Edge above). The tools read both and get the same values. A table type may say which rows it has in yes or no columns and name the look each uses. An effort's Phrases, Placeholders, Ignore, Compliance words and Page limits rows are now laid over the Global Standards. A Draft or Retired phrase is not flagged. On Fonts allowed a row is told apart by its Name and its Used for together. The tools still write the older names: no workbook changes shape.
- **A table type's own look is laid over the kind's look (1.27.0).** When a table type names the look one of its kinds of row uses (Header look: Shaded header), that kind first takes the look called by its own word (header), and the named look is laid over it: only its filled cells count. A blank cell means "not changed"; **none** in an edge cell takes the line away. So in an effort's workbook a type's look need hold only what the effort changes, as Stationkeeper 4.58.0 writes it; in the Global Standards it holds the whole look and the result is the same. Before 1.27.0 the named look replaced the kind's look, so a blank cell there meant no line.
- **An effort's change to a Global table type's look keeps the Global type's own change (1.29.0).** When an effort's workbook names another look for a kind of row of a type the Global Standards already gave a look of its own (Header look: "Kp header" where the Global said "Shaded header"), the kind's own look comes first, the Global type's look over it, and the effort's over that, filled cells only. So an effort that changes only a fill keeps the Global type's text color and the lines it takes away, as Stationkeeper does.
- **The effort's folder, in the app's words (1.28.0).** The strip and the question window say **Link the effort's folder...**; **Refresh** is now **Check now**; reports say "(the effort's folder: current)" where they said "effort data"; the kind of macro is **Needs the effort's folder**, and a macro marked in the older words keeps its mark. When the `Stationkeeper Data` folder inside the effort's folder is picked, the strip names the effort's folder. What the tools read is unchanged.
- **Choices in Word's own words (1.26.0).** Capture Styles writes Alignment as Left, Centered, Right or Justified and Line rule as Multiple, Exactly or At least; Capture Heading Numbering writes Number format as 1, 2, 3; I, II, III; i, ii, iii; A, B, C; a, b, c or Bullet, as Word shows them and as the Stationkeeper app writes them from 4.58.0. The older words (both, exact, atLeast, decimal, upperRoman and the rest) are still read. A cell is compared by what it means, not by its text, so a heading numbered 1, 2, 3 in the document and "decimal" on the sheet are the same, and writing leaves such a cell as it is. A value that goes into **Other Word settings** (a workbook with no column for it) keeps the app's older word there, because the app reads that cell as it is. Stationkeeper 4.57.0 or later reads Word's words; an older Stationkeeper does not, so update the app on a PC before it opens a workbook these tools have written to.
- **New Workbook and Add a Read Me Tab (1.25.0).** New Workbook makes a blank Global Standards, Effort Standards or Program Standards workbook with a Read me tab and every tab of the one layout, each with all its columns and no rows, under the plain column names. Stationkeeper 4.57.0 or later reads those names; an older Stationkeeper misreads them, so update the app on a PC before it opens such a workbook. Add a Read Me Tab gives an older workbook the tab that says what it is and changes nothing else.
- **This effort's standards:** optional. The app writes an effort's own values to `<effort name> Effort Formatting.xlsx`: only what the effort changes. Choose it and its values are laid over the Global Standards: where the effort sets a value it wins, a blank cell keeps the Global value, a row only the effort has is added. **Clear** goes back to the Global Standards alone. Rows the effort sets for one volume are not used, because a document does not say which volume it is. Every report names both workbooks.

Below them, **the effort's folder** link (Link the effort's folder...) is a third, separate thing: the folder the app saves an effort to, with the data it publishes there, used by the macros marked as needing the effort's folder.

**The menu's size (1.13.1).** The menu opens compact (500 by 600 points; it was 672 wide). Drag any edge or corner to resize it: the list, the description and the buttons follow, and the size is kept for the next time on that PC. The Category column in the list shows once the menu is wide enough.

**The look.** The menu, the admin and the Deck Media List open with a navy band carrying the Stationkeeper satellite and the form's name, an amber line under it, a light page, and navy main buttons (Run, Save to template, Rename). The satellite is also in each window's title bar. The pictures are stored inside the forms (`frmSKBrand` for the menu and admin, `frmSKWBrand` for the tools), so nothing extra goes on a PC. Windows keeps its own title bar color and list highlight, and message boxes stay Windows' standard ones.

**Which menu the tab opens (1.15.0).** The tab holds no menu. Each template carries its own copy, at the version last installed into it, and the tab opens the copy in the template your document is based on. When that copy is older than the release the tab came with, the tab says so once, names the version and the file, and says how to update it: open the new kit, run `StationkeeperMacrosInstall`, then `StationkeeperToolsInstall`, and pick that template.

**The ribbon tab (optional).** `Stationkeeper Ribbon.dotm` adds a **Stationkeeper** tab after View, with **Macro menu**, **Admin** and **About and removal**. Its buttons open the menu and admin of the template your document is based on; it holds no other code.

- **Add it:** keep `Stationkeeper Ribbon.dotm` beside the kit, open the kit, and run `StationkeeperRibbonInstall` (in the kit's menu as **Add the Stationkeeper Tab to Word's Ribbon**). It copies the add-in into Word's Startup folder (File > Options > Advanced > File Locations > Startup), so Word loads it every time it starts. By hand: copy the file into that folder and restart Word.
- **Turn it off for now:** File > Options > Add-ins > Manage: **Templates** > Go, and untick `Stationkeeper Ribbon.dotm` (Developer > Word Add-ins opens the same list). Tick it again to bring the tab back.
- **Remove it:** run `StationkeeperRibbonRemove` from the kit (**Remove the Stationkeeper Tab from Word's Ribbon**), or close Word and delete `Stationkeeper Ribbon.dotm` from the Startup folder.
- **SCIF or locked-down PCs:** don't install it. Nothing else depends on it; the menu still opens from View > Macros, Alt+Shift+M or the Add-ins tab button.

## First time in a template (the admin)

1. Open the template itself (File > Open the .dotm).
2. View > Macros > `StationkeeperMacrosAdmin` > Run.
3. On **Macros**, the list has three columns: the macro, **In menu** (Shown, Hidden, Not set, or Macro gone) and **Description** (Written or Missing). **Show** narrows the list: **All visible** (everything not hidden), **Stationkeeper** (made for Stationkeeper), **Hidden**, **Description missing**, and **Other** (not made for Stationkeeper). **Category** narrows it further to one of your categories, so **Stationkeeper** plus a category shows just those tools. Tests and the code health check are on **Bulk tools**. Pick each macro and fill in Title, Description, Category, Keywords, **What this macro affects** and **Does this change the document?** Tick **Show in the menu** for the ones people should see. **Use code comment** copies the comment above the macro into the description.
4. **Bulk tools** work on the whole list at once: **Fill blank descriptions from code comments**, **Show every described macro in the menu**, **Remove macros that are gone**, export and import for Excel, and **Read macros from .bas files**. Each button has a sentence beside it saying what it does.
5. Faster for many macros: keep a **primary list** (see below).
6. **Categories** sets your own categories and their order.
7. **Add from a template** copies whole modules of macros in from another template (Word's Organizer): **Pick a template...** lists its modules with their kind, how many macros, which macros, and whether a module of that name is already here. Tick modules and **Copy the ticked modules in**. A module already here is left alone unless you tick **Replace**; the kit's own modules are never copied. A dated backup of this template comes first, the new macros join the list described from the other template's catalog when it has one, and **Save** keeps them. The same tab opens from `AddMacrosFromTemplate` (View > Macros).
8. **Settings** sets the menu title, the default sort, two ways to open the menu (**Alt+Shift+M** and a button on the **Add-ins** tab), and **where backups go**.
9. Tick **Ready to save**, then **Save to template**. The first time, the admin asks where backups should go (a folder beside the template is suggested). You can change it on **Settings**. Save stays off until there is a change and the box is ticked. Each save makes a dated backup of the template first, adds a line to **History**, and clears the tick.

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
| Not related | none | Runs. Your own macros, and general tools such as Fit All Tables. |
| Built for Stationkeeper | ◇ SK | Runs on its own. If no effort folder is linked, it asks first: "This is built for Stationkeeper but is still functional. Would you like to run it?" with **Run it**, **Link the effort's folder...**, **Cancel** and **Don't ask again**. |
| Needs the effort's folder | ◆ SK | Runs only when the linked effort has the data it needs (the effort rules, template or dictionary). If not, it says "This macro is used with the Stationkeeper application. Please point to your effort's folder.", names the files it needs and where Stationkeeper writes them, says what is wrong right now, and offers **Point to the effort's folder...**. A macro marked before 1.28.0 as "Needs Stationkeeper data" keeps its mark and works the same. |

**Show** filters the list: All macros, Stationkeeper (made for Stationkeeper), and Other (not made for Stationkeeper). **Category** narrows it to one of your categories (the admin's order first), so you can see, for example, just the Stationkeeper tools in Tables.

### The Stationkeeper link

The strip under the list shows **the effort's folder**, with **Link the effort's folder...** (or **Change...**), **Check now** and **Unlink** (the same words as the app since 1.28.0: decisions 253 and 254).

- **One folder pick.** Pick the effort's folder that Stationkeeper saves to, or the `Stationkeeper Data` folder inside it; either way the strip names the effort's folder (1.28.0). It can be a local folder or a synced SharePoint or OneDrive folder. Each person links once on their own PC, and the link works for every template.
- **Read fresh every time.** The menu reads the folder each time it opens, before each run, and on **Check now**. A save in Stationkeeper, or a colleague's save once OneDrive has synced it, shows up on the next run, with no reopening.
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

**Update from a list file** does the same from any file: Excel (.xlsx, .xlsm, .xls), CSV or **JSON**. A Word file is refused as a list (1.9.1).

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

## Stationkeeper Tools (0.4.0 and later; 90 tools in 1.25.0)

Checking and fixing tools for Word documents, PowerPoint decks and Excel workbooks, in the same kit. They work on their own (View > Macros) and need nothing from the Stationkeeper app. With the menu installed, they show up in it already described.

**Add them:** open the kit, View > Macros > `StationkeeperToolsInstall`, pick the template. That copies seventeen modules and three forms (`SKW_Lib`, `SKW_Setup`, `SKW_Fields`, `SKW_Review`, `SKW_Styles`, `SKW_Tables`, `SKW_Layout`, `SKW_Captions`, `SKW_Lists`, `SKW_Acronyms`, `SKW_Text`, `SKW_Output`, `SKW_Utils`, `SKP_Deck`, `SKP_Fix`, `SKP_Tidy`, `frmSKPMedia`, `frmSKWBrand`, `SKX_Book`). Running it again updates them; `StationkeeperToolsRemove` takes them out. Or import the files in `Import files/Stationkeeper Tools/`. Open the template once afterwards: the menu lists the tools with titles, descriptions and categories. Save once in the admin to keep that in the catalog.

**Team rules live in one Excel workbook, or come from the linked Stationkeeper effort (see The Stationkeeper link).** `MakeStandardsWorkbook` makes one with starting values (allowed styles, table types, placeholders, settings, plus sheets later tools use) and opens it to fill in. Each person runs `ChooseStandardsWorkbook` once per template to point the tools at the shared copy. Without a workbook the tools use sensible defaults and say so on every report.

### Quick tools (1.10.0; one source in 1.11.0)

Fast fixes for tables and captions. **Every value comes from the Stationkeeper app's Global Standards workbook and from nowhere else.** With no workbook chosen, or a sheet missing, a tool says so and changes nothing. (1.10.0 and 1.10.1 kept their own "quick values" on the PC. That was a second source of truth and is gone.)

| Tool | What it does | Changes the document? |
|---|---|---|
| Quick Table Format (Table at Cursor) | Header row style, body rows style, cell margins | Yes, one Ctrl+Z |
| Quick Table Format (All Tables) | The same as a report for every table, then the fix in a copy | No, report; the fix works on a copy |
| Quick Caption Format (Caption at Cursor) | The caption's paragraph style, what follows the number, the style's tab stops and indents, and the action caption's character style | Yes, one Ctrl+Z |
| Quick Caption Format (All Captions) | Every caption shown as title and action caption, then the fix in a copy | No, report; the fix works on a copy |
| Caption Numbering Check | Every caption's field codes against the Captions sheet, then the repair in a copy | No, report; the fix works on a copy |
| Fix Caption Number (Caption at Cursor) | Repairs one caption's number; makes a typed number fields | Yes, one Ctrl+Z |

**What to fill in the workbook**

| Sheet | Row | Column | Used for |
|---|---|---|---|
| Table looks | `header` | Tag style | The header row's paragraph style |
| Table looks | `body` | Tag style | The body rows' paragraph style |
| Table looks | `keep` (a row you add) | Recognize (a column you add) | Styles left alone in body rows, separated by `;` |
| Table types | the first row | Cell margins: top, bottom, left, right | Cell margins, in points |
| Captions | one row per label and part (Table body, Table appendix, Figure body, Figure appendix) | Label | The word the caption starts with |
| Captions | | Style | The caption's paragraph style |
| Captions | | After number | What sits between the number and the title: `tab`, `period_tab`, `period_space`, `colon_space` |
| Captions | | Chapter level | The heading level whose number leads the caption number. Blank: counted straight through |
| Captions | | Separator | Between the heading number and the count, such as `-` |
| Captions | | Identifier | What the caption counts as (the SEQ name) |
| Captions | | **Action style** (a column you add) | The action caption's character style |
| Settings | `caption_abbreviations` (optional, a row you add) | Value | Periods that do not end a title, separated by `;`. Without it a built-in list is used (U.S., vs., No. and others) |

- Only **Approved** rows count. The app keeps the columns and rows you add and never sends them to Word.
- **The header row** is row 2 when row 1 holds a caption number field, else row 1. A caption row is left alone.
- **The action caption** starts at the first period followed by a space in the caption's title. The period goes with the action caption, so a list of tables or figures built from the title shows no period before its dot leader. A period after an abbreviation, after single letters, or before a lower-case word does not count. A caption with no period has no action caption. A caption that ends in a period with nothing after it is reported and not changed.
- **A caption follows the row for its label and the heading level its number takes.** A caption under a level-7 appendix heading follows the appendix row. A caption with no heading number follows the body row.
- **What the numbering check finds:** a typed number, a number counted under the wrong name, a missing heading number, a wrong separator, a count that does not restart under its heading, a heading number or restart where the row says counted straight through, and the cross-references that point at a caption the fix would change.
- **What it leaves alone and names:** a caption that takes its heading number from a level the sheet has no row for, or from a style by name, and a number that repeats the one before it (`\c`).
- **Tab stops and indents come from the caption style.** A caption in its style whose tab stops or indents were set by hand takes the style's.
- The repair changes field codes only. Caption text is not touched. In the copy, numbers, cross-references and lists of tables and figures are updated.
- Quick Table Format sets only the header row, body rows and cell margins. Apply Table Roles sets every role, fill and edge from the same workbook, and makes a missing Tag style.
- With an effort linked, reports say the values are the Global Standards and the effort may differ.
- Captions in text boxes, headers and footers are not read. Cell margins are set on the table; a cell with margins of its own keeps them.

### Headers and footers by page size (1.12.0)

Headers and footers belong to sections, and so do page size and orientation. These two tools work on every section of one page size and orientation together.

| Tool | What it does | Changes the document? |
|---|---|---|
| Header and Footer Map | Every section: the page it starts on, its page size and orientation, what its header and footer hold, whether each follows the section before, different first page and odd-even, and where page numbers restart | No, report |
| Replace Headers and Footers (by Page Size) | Gives every section of one page size and orientation the same header and footer. It asks which page size, where the new ones come from (the section the cursor is in, or a Quick Part by name), and whether to replace the header, the footer or both. It reports first, then works on a copy | No, report; the fix works on a copy |

- **Pagination is protected.** No section break is added, removed or moved. No section's page number settings (restart, start at, format) are touched.
- **Neighbours are protected.** A section of another page size whose header or footer follows a section about to change is unlinked first, so it keeps what it shows now. The report lists each one.
- **Proof.** The copy's page count and the page each section starts on are read before and after. The result says "Pagination is as it was", or gives both readings when a taller header or footer moved text.
- **From the section at the cursor:** the main header and footer are copied, and the first-page and even-page ones where both the source and the target section use them.
- **From a Quick Part:** the main header or footer is filled. First-page and even-page ones are left as they are, and the report says so. The Quick Part is found by name in the templates Word has loaded.
- A floating object in the source header (a text box, a picture set to float) that does not carry over is named in the result.
- The source section itself is not changed.

### Capture into the Global Standards workbook (1.17.0)

The Stationkeeper app reads its Global Standards workbook again whenever the file changes. These tools write rows to it in the app's own shape, so what you capture in Word reaches the app.

| Tool | What it writes | How |
|---|---|---|
| Capture Styles to the Global Standards | The document's styles that the **Styles** sheet lacks | Added as **Draft** rows. The app holds a Draft: it reaches no template until you approve it in the app. A style already on the sheet is never changed |
| Capture Banned Phrase | The selected words, or a phrase you type, on the **Phrases** sheet | Banned or Avoid, with what to use instead and why. A phrase already there is left as it is |
| Capture Heading Numbering | The document's numbering for Heading 1 to 9 on the **Heading numbering** sheet | A report of what differs first. Writing changes the rows that differ and keeps each row's status, so an Approved row changes the standard for everyone |

The rules every write keeps:
- **Report first, write on a yes.**
- **Nothing is removed.** Rows are added or changed, never taken out.
- **Only columns the workbook has are filled.** A style's other values go in the row's Other Word settings, where the app reads them.
- **Not while Excel has the workbook open.** The tool says so and writes nothing.
- **A dated copy first**, in `%LOCALAPPDATA%\Stationkeeper Macros\Workbook backups`.
- The app gives a new row its ID and stamps Changed when it reads the workbook.

**Banned Phrases Review** now reads the Phrases sheet first. The app's "Use instead" is advice, so those phrases are flagged with a comment and not replaced.

Not captured yet, and why: **acronyms** (the app syncs its acronym file whole, and when both sides changed the newer file replaces the other, so a macro writing to it could lose the app's own changes); **Quick Parts** (they live in Word files, and the app's Bring in parts reads them); **page numbering by part** (not in the workbook).

### Remap Style by what each paragraph is (1.14.0, a window since 1.15.0)

**The window (1.15.0).** Remap Style opens a window; nothing is typed.

1. **Style to change:** a list of the styles the document's text is in. It starts on the style at the cursor.
2. **The rows:** each kind of paragraph in that style, with its count and how many carry formatting by hand: not in a list; bulleted or numbered by hand, with its level; in a table; typed bullet.
3. **Becomes:** beside each row, a list of the paragraph styles the document uses. Tick **All Word styles** to add Word's built-in styles it has not used yet (Body Text, for one). A bulleted or numbered row starts on the list style the Global Standards give its level; pick another for this run, or pick one where the workbook has none. A row left on "(leave alone)" is not changed.
4. **Formatting applied by hand:** keep all, keep emphasis only, or clear.
5. **Apply to:** the whole document, or the selected paragraphs only. The window stays open, so you can select text in the document and apply again.
6. **Apply.** Each Apply is one Ctrl+Z. **Report...** makes the full report and changes nothing.

Picks are for the run and are stored nowhere.

Remap Style reads every paragraph in the style you name before it changes anything, so one style can go to several.

| The paragraph is | It becomes |
|---|---|
| Not in a list | The style you give (leave it blank to leave these alone) |
| Bulleted or numbered by hand | The **Style** of the approved row for that kind and **Level** on the Global Standards' **List numbering** sheet. The list style's own bullet replaces the one applied by hand. With no such row, or a style the document lacks, it is left alone: a bullet is never lost |
| In a table | Left alone, unless you apply that one row |
| Led by a typed bullet | Left alone, and listed |

- **Report first.** One numbered row per style and kind: how many paragraphs, how many carry formatting by hand, what they become, and why any are left alone.
- **Apply** every row, one row by its number, or only the selected paragraphs. Each run is one Ctrl+Z, so a document can be changed a step at a time.
- **Formatting by hand**, chosen per run: keep all of it; keep emphasis only (bold, italic, underline, strikethrough, superscript, subscript, color, highlight) while fonts, sizes, indents and spacing take the style's; or clear it.
- **List Paragraph.** Word usually moves text to List Paragraph when someone bullets it from Normal. Changing Normal alone, the report says how many places are in List Paragraph; name it as a second style in the same run.
- Several styles can be named in one run. A character style changes as before: its text takes the other style.
- The tools' own Style map sheet is no longer read. The list styles come from the Global Standards workbook; the style for plain paragraphs is asked for and not stored.

### Chapter numbers in page numbers (1.13.0)

One action sets every section's page numbers to include the chapter number (1-3) or to leave it out (3).

| Tool | What it does | Changes the document? |
|---|---|---|
| Page Numbers: Chapter Numbers On or Off | Asks: include the chapter number, or leave it out. Reports every section's page numbers now and what would change, then sets every section in a copy | No, report; the fix works on a copy |

- **One source.** The chapter's heading level and the separator come from the Global Standards workbook: the Captions sheet's Chapter level and Separator on the approved body row, the same values that number captions (Table 1-2). The on or off choice is made at run time and is not stored anywhere.
- With no workbook chosen, or captions without a chapter level, the tool asks for the heading level and the separator.
- **Three settings only.** On each section: include the chapter number, the chapter's heading level, the separator. Restart, start at and number format stay. What the footers hold is not touched.
- **It refuses with the reason** when the chapter heading level is not numbered in the document. Word makes a page's chapter number from a numbered heading.
- **Proof.** The copy's page count and the page each section starts on are read before and after.
- The app's templates write the same Word setting through the template's Numbering pick. Run this tool to match a document to that pick; it does not change the pick.

### Page numbers restart at each chapter, or continue (1.13.0)

| Tool | What it does | Changes the document? |
|---|---|---|
| Page Numbers: Restart at Each Chapter or Continue | Asks: start at 1 at each chapter, or continue through the whole document. Reports every section's numbering now and what would change, then sets the sections in a copy | No, report; the fix works on a copy |

- **Start at 1 at each chapter.** A chapter starts where a section's first line with text is the chapter heading (the Captions sheet's Chapter level, else Heading 1). That section restarts at 1. Every other section continues, so a landscape page inside a chapter keeps the chapter's count. A first section before any chapter (a cover, front matter) is left alone.
- **Continue.** Every section after the first continues from the one before. The first section is left alone: it sets where the numbers start.
- **Two settings only.** Restart and start at. Number format, chapter numbers and what the footers hold are not touched.
- With no section opening with the chapter heading, the restart is refused with the reason.
- The same proof: page count and each section's first page, before and after.

### Choosing a workbook kept in OneDrive or SharePoint (1.11.0)

The menu's **Global Standards...** button (bottom row) runs Choose Standards Workbook. The link strip above it is for an effort's data folder and does not take a workbook.

Word's file picker names such a file by its web address. Choose Standards Workbook now finds the same file in the synced folder on this PC and uses that. It asks nothing of the network. A workbook that is not synced to the PC is refused with the reason, and the workbook already in use stays. A workbook that cannot be read is refused with Excel's reason.

### One workbook with the Stationkeeper app (1.9.0, Global Standards in 1.9.1)

Point the tools at the app's **Global Standards.xlsx** (`ChooseStandardsWorkbook`; it was Company Standards.xlsx before app 4.31.0) and the table tools read it directly. No second workbook.

With an effort linked (1.9.1), the looks and the table types still come from the Global Standards workbook: the app keeps an effort's own set of standards and does not publish it to the macros yet. Every table report says so on its first lines: these are the Global Standards and the effort may differ. With no Table looks sheet in the workbook, the built-in roles are used (alignment and bold only), and a table cannot be marked assigned.

| App sheet | What the tools take from it |
|---|---|
| Table looks | One look per row role: caption, header, header2, subhead, body, subtotal, total, note. Font, size, bold, italic, color, alignment, vertical alignment, spacing, indents, fill, every edge, the Tag style, and Repeat. |
| Table types | Cell margins per side (points), width, the table's place on the page, the roles a type uses, and its overrides. |
| Settings | Named values the other sheets point to with `=name`. |

- Only **Approved** rows count.
- A **Tag style** the document lacks is made from its look, so no style has to be made by hand.
- A table's type is the first type unless its Alt Text says `type=` with the type's Name or its label.
- The app keeps columns you add and never sends them to Word. The tools read these, and nothing else reads them: a **Recognize** column on Table looks (how many header rows, the words that start a total or subtotal row, the most words in a short cell), a **Tab stops** column, and looks named **number**, **short** and **keep**. Without a number or short look, those cells take the body look.
- Every report names the workbook and when it was last changed.
- The tools' own Table roles sheet (1.8.0) and Apply Team Table Type are retired in 1.9.1. One place holds a table's look: the app's workbook.

### Assigned tables (1.9.0)

A row's role is the paragraph style on it: the role's Tag style. That one fact makes a table easy to fix, review and keep right.

- **Assign:** `Assign Table Role (Selected Rows)` sets the rows in the selection to Header, Header 2, Subhead, Body, Subtotal, Total, Note or Caption, and styles them.
- **Fix:** `Table Roles Check` (a copy of the whole document) or `Apply Table Roles (Table at Cursor)` reads each row's role and applies its look.
- **Review:** `Show Table Roles` opens Word's Draft view with the style area, so every row's style name shows at the left. Run it again to go back.
- The first fix of a table suggests its roles from what it can see, then marks the table **assigned** with a hidden bookmark (`_SKA...`). The mark is not in the Alt Text, so accessibility text stays as it was.
- An assigned table is read from its styles and never guessed again: a row set by hand stays set. A row with no role is read as body and named in the report.
- In a table that is not assigned yet, a header, subhead, subtotal, total, note or caption style someone already set wins over a guess.
- **Suggest again** (the last choice in Assign Table Role) forgets a table's roles and reads it afresh.
- A table can be assigned only when the roles have styles: the Global Standards' Tag styles.

### The progress window (1.8.1)

A tool that runs longer than a moment shows a small window: the tool's name, the item it is on (Paragraph 1,200 of 9,400), a bar, the percent done, the time it has run and the time left. The same line goes on Word's status bar.

- The time left appears after four seconds, once there is a pace to go by.
- A set (Delivery Check, Styles and Fonts Check, Pagination Check, Deck Delivery Check, Book Check) names the check it is on and shows the percent of the whole set.
- A step with no count (opening a copy, saving) shows a full amber bar.
- The window has no buttons. Closing it hides it and does not stop the tool.
- Word gets a turn every second and a half, so its window no longer says Not Responding during a long run.

### The reports folder (1.8.1)

A report opens as a new document. Choose a folder once (the menu's **Reports folder...** button, the admin's Settings page, or the tool `Choose Reports Folder`) and every report is also saved there as it is made: `<report> - <document> - <date time>.docx`. The setting is per PC and is shared by the menu and the tools. Choose again and answer No to stop.

### Delete Unused Styles (1.8.1)

Clears the Styles pane down to what the document uses. It reports first, then asks. Unused custom styles are deleted, pass after pass, so a style goes once the style based on it is gone. Word does not delete built-in styles, so an unused built-in style that shows in the pane is hidden until it is used. Kept: styles in use, styles they depend on, the team's styles (Allowed styles list, StylePrefix), Normal and Heading 1 to 9. One Ctrl+Z undoes it.

### Table roles (1.8.0)

`Table Roles Check` reads every table and gives each row a role and each cell a kind, then offers to style every table in a copy. `Apply Table Roles (Table at Cursor)` does one table in place (one Ctrl+Z undoes it).

| Role | Recognized by | Built-in format |
|---|---|---|
| Caption row | Row 1 holds a caption number field (SEQ) | Left as it is. Text and field codes are never changed. |
| Header row | The row after the caption row, or row 1 | Centered both ways, bold. Repeats on each page (with the caption row: Word repeats heading rows only from row 1 down). |
| Divider row | One merged cell across the table | Left, centered vertically, bold |
| Total row | First cell starts with Total, Subtotal or Grand Total | Bold; its number cells stay right aligned |
| Number cell | Digits with $ , . % ( ) + -, or a dash | Right, centered vertically |
| Short text cell | 1 to 3 words | Centered both ways |
| Long text cell | Everything else | Top left |

- The built-in formats above apply when the workbook has no Table looks sheet. With the app's workbook, each role takes its look from Table looks. **Tab stops** (inches, left tabs) and a **figure caption** look are read from the same sheet (see One workbook above).
- A table's Alt Text (Table Properties) can hold `skip` (leave the table alone) or `type=Name` (use that table type).
- The report's Look at column names anything the tool was unsure of: a row with text in its first cell only (when the Divider role recognizes `first cell only`), a caption row with no number field, a table inside a table.

| Tool | What it does | Changes the document? |
|---|---|---|
| Safe Update All Fields | Updates every field in every part (body, headers, footers, notes, text boxes), then contents and figure tables. Fields that pull in outside files (LINK, INCLUDETEXT, INCLUDEPICTURE) are kept as they are. Track changes is paused for the update and put back. | Yes, one Ctrl+Z |
| Find Broken References | Cross-references to bookmarks that are gone, fields showing Error!, internal links that go nowhere, included files that are missing | No, report |
| Field Code Report | Every field: type, code, result, where, page, locked | No, report |
| Pre-Delivery Leftovers | Comments, tracked changes, hidden text, highlighting, placeholders (TBD, XXX, [insert...), author and company saved in the file | No, report |
| Delivery Check | Category Before delivery (1.7.0). Runs 15 document checks in one go and makes one report: a summary table (Check, Result, Findings), then each check's part. The 15: Leftovers, Broken references, Tables of contents, Styles, Fonts, Direct formatting, Tables, Captions, Keep with next, Breaks and spacing, Page limits, Graphics, Typed lists, Acronyms, Typography. Banned Phrases Review is not part of it (it has its own review flow) | No, report |
| Style Inspector | The paragraph and character style at the cursor and anything formatted by hand on top; several paragraphs make a report | No |
| Styles and Fonts Check | One report with two parts. Styles: styles in use that are not on the team list, required styles that are missing (without a list: custom styles). Fonts: every font used, flagging ones not on the Fonts sheet | No, report |
| Direct Formatting Report | Paragraphs whose font or spacing was set by hand. Bold, italic, superscript and subscript count as emphasis (setting `KeepOnClear`). In tables only font and size are checked, since table styles set the rest | No, report |
| Clear Direct Formatting (Selection) | Clears hand formatting in the selection; keeps emphasis and list numbering | Yes, one Ctrl+Z |
| Table Standards Audit | Each table against the team table types: style, repeating header row, rows breaking across pages, width, floating or nested tables, cell styles, a caption above or below | No, report |
| Section Map | Every section: first page, orientation, paper, margins, columns, how it starts, headers and footers linked or own, page number restarts | No, report |
| Banned Phrases Review | The team's banned phrases (Banned phrases sheet) in the selection or the whole body: replacements as tracked changes, a comment naming the rule on each | Yes, one Ctrl+Z |
| Caption Audit | Tables and figures without captions, typed caption numbers, wrong labels, captions on the wrong side (settings TableCaptionPosition, FigureCaptionPosition), repeated captions | No, report |
| Import Team Styles | Brings the style definitions in from the team template (setting StyleSource, else the attached template) and lists which changed | Yes, one Ctrl+Z |
| Remap Style | Changes a style to another by what each paragraph is (see below) | Yes, one Ctrl+Z |
| Add Rows to Table (1.18.0) | Click in a table. It says how many rows the table has and asks how many more you need, then adds them at the end or below the cursor's row. The new rows are empty and take the look of the row they follow | Yes, one Ctrl+Z |
| Style Cross-References (1.19.0) | Gives every cross-reference (a REF, PAGEREF or NOTEREF field) the team's cross-reference character style and sets each field to keep it when fields update (`\* CHARFORMAT` in place of `\* MERGEFORMAT`). The style's name comes from the Global Standards: a Settings row named `cross_reference_style`. Without one, you pick from the document's character styles by number. Reports first. Tables of contents and figures are left alone | Yes, one Ctrl+Z |
| Stationkeeper Admins (1.20.0) | Not a tool. The three Capture tools change the Global Standards, so they are for Stationkeeper Admins. The list is a sheet of the Global Standards workbook named **Stationkeeper Admins**, with one column, **Name**. A person on it sees the Capture tools; anyone else gets the same menu without them, and a Capture tool run another way says "You are not listed as a Stationkeeper Admin" and writes nothing. A name matches the Windows user name or Word's user name (File > Options > General), ignoring case. With no such sheet, the Capture tools show for everyone and the menu's Global Standards line says "no admin list yet". It is a shared way of working, not a lock | Nothing to undo |
| Check This Workbook (1.22.0) | Asks for a standards workbook and says what is wrong with it in plain words. It reads what the workbook says it is: the line **This workbook is** on its **Read me** tab (Global Standards, Effort Standards or Program Standards), with **It belongs to** beside it. Then each tab: a tab the tools know, an older tab of the macros' own workbook, or a tab of your own (kept, never read). For a list it names a first row without column names, a missing Name column, rows with no Name, a Name used twice, a Status that is not Approved, Draft or Retired, and merged cells, each with its tab and row. An older workbook from the app (an About tab, no Read me) is said to be one and is still read. The workbook is only read. The same tool is to be in the Stationkeeper app under the same name | No, report |
| New Workbook (1.25.0) | Asks for the kind (Global Standards, Effort Standards, Program Standards), the name and a folder, and makes a blank workbook there: **Global Standards.xlsx**, or the name followed by **Effort Standards.xlsx** or **Program Standards.xlsx**. The Read me tab is first, with **This workbook is** in A1 and the kind in B1 (a drop-down), **It belongs to** in A2 and the name in B2, then the text both tools use. After it come Settings, Styles, Fonts allowed, Table looks, Table types, Heading numbering, List numbering, Captions, Page setups, Page numbering, Templates, Theme, Phrases, Placeholders, Ignore, Style map, Quick Parts and Pages; a Global Standards workbook also has Stationkeeper Admins, an effort's or a program's has Page limits and a Volume column after Name. Every tab has all its columns and no rows. The ID column is out of sight, the less used columns are folded (the plus signs above the columns open them), and switches and agreed choices carry a drop-down. It never writes over a file. No format number is stated yet. The app's own tabs (Fields, Authors, Compliance words, Piece types, Document types, Suggestion rules, Columns) are not made: the app adds what it needs. |
| Add a Read Me Tab (1.25.0) | Asks for a workbook and adds a Read me tab as its first tab, with the same two lines and text, so the workbook says what it is. The kind is taken from the file's name (Global Standards, Effort Standards, Effort Formatting, Program Standards) or asked for. Nothing else in the workbook is changed, and its About tab stays. A workbook that already says what it is is left alone; a Read me tab with other words in its first two rows is left alone and the message says what to type. The workbook must be closed in Excel. |
| Work PC Proof 1: This PC and Word's Styles (1.21.0) | One of four proofs under the category **Work PC proofs**. They find out what Word, Excel and a synced folder allow on a PC before the workbook layout is built on it, and they leave the kit once answered. Each report holds setting names and counts only: no style name, no value, no file name. Proof 1 needs nothing: it checks Excel started from Word, settings and files kept on the PC, a Stationkeeper window, synced folders and Word's spelling of capitals. Then, in a scratch document that is not saved, it sets and reads back each of the 49 style settings, sets a follower's value equal to its base and sees whether it follows the base again, and hands a style over as XML. It never sets a style's link: on the build PC that closed Word. If Word closes during a proof, `%TEMP%\Stationkeeper proof trace.txt` names the step | No, report |
| Work PC Proof 2: This Document's Styles (1.21.0) | Open a document made from the team template first. Reads every paragraph and character style the file holds: which settings Word hands over, which differ from the base style, which the style's own XML holds. Then makes the styles again in a scratch document and counts, per setting, what came back different. The document is not changed | No, report |
| Work PC Proof 3: A Workbook (1.21.0) | Asks for a workbook. Times how long it takes to read whole, then saves one row to a copy in the temporary folder through the tools' own writer and counts, tab by tab, whether the columns, hidden columns, folded columns and drop-downs are still there. The workbook chosen is only read; the copy is deleted | No; it works on a copy it deletes |
| Work PC Proof 4: A Shared Workbook (1.21.0) | Asks for a synced folder. The first run makes `Stationkeeper proof scratch.xlsx` there. Run again while a colleague has that file open on another PC or in the browser: it writes one row and reports whether this PC saw the file as held, whether the row landed, and how many scratch workbooks are in the folder (more than one is a conflict copy). Run once more a few minutes later | Only its own scratch workbook |
| Pagination Check | One report with two parts. Keep with next: headings and captions above tables that can be stranded at a page bottom, long kept chains. Breaks and spacing: manual page and column breaks, empty paragraphs used as spacing, line breaks ending paragraphs, body text starting a new page | No, report |
| Typed Lists Finder | Paragraphs starting with a typed bullet or number instead of a real list | No, report |
| TOC Check | Headings missing from the contents, entries whose heading changed, page numbers out of date | No, report |
| Acronym Quick Check (text only) | Acronyms in the typed text: uses, first use in running text (not tables or headings), spelled out there as "Words (ACR)" or not, team meaning from the Acronyms sheet or the Stationkeeper app's export (setting AcronymList). Pictures are not scanned and the report says so; the app does the full check | No, report |
| Typography Sweep | Double spaces, space before punctuation, straight quotes, double or spaced hyphens, three dots, spaces at paragraph ends | No, report |
| Load Team Dictionaries | Adds the dictionaries on the Dictionaries sheet to Word on this PC | Word settings only |
| Page Limit Counter | Pages per top-level part and in total against the Page limits sheet and PageLimitTotal | No, report |
| Graphics Audit | Pictures, charts and shapes: inline or floating, size, too wide, scaled up, no alt text, missing linked file | No, report |
| Metadata Clean Copy | Saves "<name> - clean.docx" beside the document without author, company, template path, variables and personal details | No, saves a copy |
| PDF with Bookmarks | Saves a tagged PDF with heading bookmarks, without file details | No, saves a PDF |
| Remove Unused Custom Styles | Deletes custom styles not applied anywhere, not on the team list and not depended on | Yes, one Ctrl+Z |
| Unlink Fields (Selection) / Lock or Unlock Fields (Selection) | Unlink Fields turns the fields in the selection to plain text (contents and indexes are never unlinked). Lock or Unlock Fields locks the fields in the selection; when every field there is already locked, it unlocks them | Yes, one Ctrl+Z |
| Convert to Team List | Selected typed bullets and numbers become the team list styles (BulletStyle, NumberStyle); the typed markers go | Yes, one Ctrl+Z |
| Fit All Tables | Tables wider than the margins fit the page width (or only those in the selection) | Yes, one Ctrl+Z |
| Go to Next Finding | Selects the next finding of the last report in the document it checked | View only |
| Install Tool Shortcuts | Ctrl+Alt+Shift+N, U, C, R and L for the most used tools, leaving keys already in use alone | Template |
| Insert Acronym List | An Acronym and Definition table at the cursor, A to Z, from the team meanings (only those used, unless AcronymListUsedOnly is No) | Yes, one Ctrl+Z |
| Reset Paragraph to Style | The paragraph at the cursor or the selected ones back to their style, emphasis included | Yes, one Ctrl+Z |

**PowerPoint (0.8.0).** These check the deck open in PowerPoint. They never start PowerPoint or open a file on their own, and their reports open in Word.

**A check offers its fix (1.7.0).** After the report opens, a message shows the result and asks "Fix a copy now? The report stays open." **Yes** runs the fix on a copy saved beside the deck ("<name> - fixed <date> <time>.pptx"); the open deck is never changed. The offer appears only when the check found something its fix can put right. The fixes are no longer separate menu entries: each is reached from its check.

| Tool | What it does | Changes the deck? |
|---|---|---|
| Deck Leftovers | Speaker notes, comments, hidden slides, placeholder text | No, report |
| Deck Banned Phrases | Banned phrases in slides and notes, with replacements. When it finds any, it offers to fix a copy: replaces banned phrases that have a replacement (a capital at the start of a sentence is kept), and adds a comment on each changed slide listing what changed. Phrases with no replacement are flagged in the comment | No, report; the fix goes to a copy |
| Deck Acronym Check (text only) | Acronyms, first slide, spelled out or not, team meaning. When it finds any, it offers to fix a copy: spells out each team acronym at its first use in body text ("Request for Proposal (RFP)"), with a comment on each changed slide. Titles stay short; acronyms used only in titles are listed | No, report; the fix goes to a copy |
| Deck Font Audit | Fonts used, flagged against the Fonts sheet. When it finds any, it offers to fix a copy: text in fonts not on the Fonts sheet goes back to the theme's heading or body font (and its East Asian and complex script slots) | No, report; the fix goes to a copy |
| Deck Small Text | Text under DeckMinFontSize (10 pt) | No, report |
| Deck Alt Text Audit | Graphics without alt text (Decorative ones pass) | No, report |
| Deck Off-Slide Objects | Shapes partly or wholly off the slide | No, report |
| Deck Titles Check | Slides without a title, slides sharing one | No, report |
| Deck Empty Placeholders | Empty placeholders left on slides | No, report |
| Deck Delivery Check | (1.7.0) Runs all 20 deck checks in one go and makes one report: Leftovers, Banned phrases, Acronyms, Fonts, Small text, Alt text, Off-slide shapes, Titles, Empty placeholders, Template, Overlaps, Slide numbers, Marking, Text overflow, Layout drift, Pictures, Colors, Language, Links, Table style | No, report |
| Deck Delivery Copy | "<name> - delivery.pptx" without notes, comments, hidden slides or personal details | No, saves a copy |

**PowerPoint fixes (1.4.0).** The team template is the `DeckTemplate` setting (a .potx or .pptx on the Settings sheet). Without it, the master most slides use stands in. Fixes never touch the open deck: they save "<name> - fixed <date> <time>.pptx" next to it and open that, with a report of every change.

| Tool | What it does | Changes the deck? |
|---|---|---|
| Deck Media List | A window listing every picture, icon, video, audio clip, chart, SmartArt and embedded or linked object, by slide, with its name and alt text. Filter by kind, "No alt text" or "Default names"; find; rename (names stay unique on a slide); suggest a name from the alt text; name every "Picture 3" from its alt text at once; go to it in PowerPoint | Names only, in the open deck (save it in PowerPoint) |
| Deck Template Check | Every master and layout: the template's, pasted-in copies ("1_Title and Content"), foreign masters from other decks, unused ones, template layouts gone missing. When it finds any, it offers to fix a copy, with three answers. **Yes** moves each slide on a copy or foreign layout onto the template layout with the same name (or the same placeholders), then deletes the layouts and masters no slide uses. **No** only deletes layouts no slide uses that are not in the template, and masters no slide uses; template layouts stay. **Cancel** changes nothing | No, report; the fix goes to a copy |
| Deck Overlap Check | Slide shapes sitting on the template's slide number, footer, date, title area, markings (CUI, Proprietary...) or master graphics such as a logo. High for number, footer and markings. Also saves a copy with each overlap outlined in red | No, report and a marked copy |
| Deck Slide Number Check | Slides with no number, layouts with no number placeholder, numbers typed as text, numbers moved or restyled, numbering that does not start at 1. When it finds any, it offers to fix a copy: puts the number placeholder back on masters and layouts, turns numbers on (off on title slides), removes typed numbers, puts each number where the layout has it in its font | No, report; the fix goes to a copy |

Settings for the slide numbers: `DeckFirstSlideNumber` (1) and `DeckNumberTitleSlide` (No). Hidden slides count but need no number.

**PowerPoint tidy-up (1.5.0).** Same rule: checks change nothing, fixes save a copy. New settings: `DeckFooterText`, `DeckMarking`, `DeckLanguage` (en-US), `DeckTableStyle`, `DeckExtraColors`. A linked Stationkeeper effort can set them too. `DeliveryCheckSkip` (1.7.0, blank by default) lists checks the delivery checks leave out (see Delivery checks).

| Tool | What it does | Changes the deck? |
|---|---|---|
| Deck Marking Check | Slides without the team footer (`DeckFooterText`) or the marking (`DeckMarking`, such as CUI) in their text, footer, layout or master. Hidden slides count. When it finds any, it offers to fix a copy: puts the footer placeholder back on masters and layouts and shows the team footer on every slide | No, report; the fix goes to a copy |
| Deck Text Overflow Check | Text running past the bottom of its box or off the slide; tables running off the slide | No, report |
| Deck Layout Drift Check | Titles and other placeholders moved, resized or restyled away from their layout. When it finds any, it offers to fix a copy: moves and resizes them back; titles back to the layout's font and size | No, report; the fix goes to a copy |
| Deck Picture Check | Stretched pictures, pictures enlarged past 150 percent, cropped pictures whose cut-off parts still travel in the file. Measured on a scratch copy. When it finds stretched pictures, it offers to fix a copy: gives them their own shape back, fitted inside the box they had. Only stretched pictures are fixed | No, report; the fix goes to a copy |
| Deck Color Check | Fill, line and text colors that are not theme colors (or `DeckExtraColors`), with the nearest theme color | No, report |
| Deck Language Check | Text marked with a proofing language other than `DeckLanguage`. When it finds any, it offers to fix a copy: marks all text (slides, notes, masters, layouts, SmartArt) and the deck default with the team language | No, report; the fix goes to a copy |
| Deck Links Check | Linked pictures, objects and media, hyperlinks and embedded objects: missing files, links into one computer or someone's own folders, embedded files carrying their data | No, report |
| Deck Table Style Check | Tables not in the team style (`DeckTableStyle`) and tables with no header row. When it finds any, it offers to fix a copy: applies the team table style to every table | No, report; the fix goes to a copy |

**Excel (0.9.0).** These check the workbook open in Excel (the active one, or the only one open). They never start Excel or open a file on their own, and their reports open in Word. On very large sheets they read the first 20,000 text cells of each sheet.

| Tool | What it does | Changes the workbook? |
|---|---|---|
| Book Check | (1.7.0) One report with nine parts: formula errors (cells showing #DIV/0!, #N/A, #REF!, #VALUE! and the rest, with the formula), broken names (defined names that point at #REF!, at nothing, or at another file), external links (other workbooks this one links to, and the cells that use them), hidden content (hidden and very hidden sheets, hidden rows and columns, hidden names), leftovers (notes, comments and placeholder text such as TBD, XXX, [insert), banned phrases (in cell text, with replacements), acronyms (text only: first cell, spelled out or not, team meaning), fonts (flagged against the Fonts sheet) and merged cells (they break sorting, filtering and screen readers) | No, report |
| Book Delivery Copy | "<name> - delivery.xlsx" without notes, comments, links to other workbooks (kept as values) or personal details. Hidden sheets stay; Book Check lists them | No, saves a copy |

How they behave:
- **Reports** open as a new Word document with one table and the standards used at the top. They change nothing in your document, not even its Undo list (the report is written as a web page and opened, so Word edits nothing while making it).
- **Fixes** are one step on Undo, named "Stationkeeper: ...".
- **Co-authored files:** on a SharePoint or OneDrive file with AutoSave on, the fixes do not run (to protect co-authors' edits); the reports still do.
- **Test every macro** in the admin runs the report tools quietly and closes their reports; the setup tools wait for a person.

## Delivery checks (1.7.0)

Three tools run every check in one go, so one run before delivery covers the whole file:

- **Delivery Check** (Word): 15 document checks.
- **Deck Delivery Check** (PowerPoint): all 20 deck checks.
- **Book Check** (Excel): nine workbook checks.

Each makes one report: a summary table (Check, Result, Findings) at the top, then one part per check. They change nothing. Banned Phrases Review is not part of Delivery Check, because it has its own review flow.

**Leave checks out:** `DeliveryCheckSkip` on the standards workbook's Settings sheet (blank by default) lists checks the three delivery checks leave out, by their name in the report, separated by semicolons, for example "Colors; Typography; Merged cells". The summary table shows a left-out check as "Left out (DeliveryCheckSkip)". New standards workbooks come with the row; add it by hand to an older one if you want it.

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
powershell -ExecutionPolicy Bypass -File test.ps1    # 246 checks on a sample template; changes no settings
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

- PowerPoint, next: a deck-wide fix-everything run (template, fonts, numbers, footer, language, table style in one copy), notes checks, and the effort's own deck template from the linked Stationkeeper effort.
- Shared findings with the Stationkeeper app; both directions (a hand-filled setup workbook the app imports).

- Library mode: add or remove whole macro modules from a master Macro Library template, and export one macro family as its own small template.
- Load Nicholas's ProposalEngine macros as `SKM_` modules with their catalog filled.
- Stationkeeper app: show "Macros in this template" in the template library.
