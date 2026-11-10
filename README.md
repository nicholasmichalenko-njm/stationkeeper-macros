# Stationkeeper Macros

A menu and a bookshelf for the macros in a Word template.

- **The menu** (`StationkeeperMacros`): search as you type, sort by category, A to Z, most used or recent, filter by what a macro affects, read what it does, and run it. It can stay open beside the document.
- **The admin** (`StationkeeperMacrosAdmin`): see every macro in the template, then give each one a title, description, category, keywords, what each one affects and whether it changes the document, and choose which ones show in the menu. Set and order your own categories, export the catalog to Excel and back, and check for macros that were added, moved or removed.

It works with whatever macros are already in the template: one module or many. It never edits your code. Part of the Stationkeeper suite and separate from the Stationkeeper app.

## What is in a release

| File | What it is |
|---|---|
| `Stationkeeper Macros Kit.dotm` | The kit, plus an installer that copies the menu and the admin into any template |
| `Import files/` | The same three pieces as files (`SK_Macros.bas`, `frmSKMenu.frm/.frx`, `frmSKAdmin.frm/.frx`) for importing in the VBA editor |
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
5. Faster for many macros: **Bulk tools > Export the list to Excel (CSV)**, fill it in Excel, save as **CSV UTF-8**, then **Import the list from Excel (CSV)**.
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

## Testing macros

- **Test column** in the admin's list: Passed, Failed or Not tested (on this PC). **Show > Test failed / Not tested** lists them.
- **Test this macro** (Macros tab) runs it on a scratch document made from the template, or on a fresh copy of a sample document you pick (**Bulk tools > Choose a test document**). Your real files and the sample are never changed.
- **Test every macro** (Bulk tools) runs them one after another. Macros whose code saves, closes or deletes files, or ends Word, are skipped (test them one at a time).
- If a macro hits an error it doesn't handle, Word shows its own error dialog: click **End**. That macro is then marked **Failed**. The same happens when a macro from the menu stops with an error.
- **Check the code health** (Bulk tools) runs nothing. It finds missing reference libraries (the usual cause of "works on my PC, not hers") and notes risky lines in each macro: saves, closes or deletes files; a fixed file path; Word alerts or screen updating switched off and never back on; On Error Resume Next (a test can't see errors the macro hides); a Stop line. To check the code compiles, use Debug > Compile in the VBA editor.
- Limit: a test shows the macro runs without an error. It doesn't check the result is right.

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
powershell -ExecutionPolicy Bypass -File test.ps1    # 51 checks on a sample template; changes no settings
```

`test.ps1` builds a sample template with all of its macros in one module (plus one in a second module), installs the kit, and checks scan, catalog save and reload, menu search, admin, CSV round trip, the shortcut and button, use from a document based on the template, macros that move between modules, updating the kit, and the no-access path. See `TESTING.md` for the checks run by hand.

## Layout

```
src/            SK_Macros.bas (engine + entry macros), SK_Install.bas (kit only),
                frmSKMenu.code.vba, frmSKAdmin.code.vba (forms build their controls in code)
test/           sample macros and the self-test module
tools/          wordjob.ps1 (runs one hidden Word and always closes it), watch.ps1 (reads stuck dialogs)
build.ps1, test.ps1
```

## Next

- Library mode: add or remove whole macro modules from a master Macro Library template, and export one macro family as its own small template.
- Load Nicholas's ProposalEngine macros as `SKM_` modules with their catalog filled.
- Stationkeeper app: show "Macros in this template" in the template library.
