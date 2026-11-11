# Tests the kit on a sample template laid out like a colleague's: all macros in one module (plus one more).
# Runs each step in its own hidden Word. Changes no Word or Windows settings.
# Word started by automation can always read VBA projects, so the access-off case is tested by hand (see TESTING.md).
# Usage: powershell -ExecutionPolicy Bypass -File test.ps1 [-Root <repo folder>]
param([string]$Root = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
. (Join-Path $Root 'tools\wordjob.ps1')

[string]$work = Join-Path $Root 'work'
[string]$kit = Join-Path $Root 'build\Stationkeeper Macros Kit.dotm'
[string]$sample = Join-Path $work 'Team Sample Template.dotm'
$results = New-Object System.Collections.Generic.List[string]
$fails = 0
function Check([string]$name, [bool]$ok, [string]$detail) {
    $script:results.Add(("{0}  {1}  {2}" -f ($(if ($ok) { 'PASS' } else { 'FAIL' })), $name, $detail))
    if (-not $ok) { $script:fails++ }
    Write-Output ("{0} {1} :: {2}" -f ($(if ($ok) { 'PASS' } else { 'FAIL' })), $name, $detail)
}
function OpenDoc($word, [string]$path, [bool]$readOnly = $false) {
    $a = $path; $b = $false; $c = $readOnly
    $word.Documents.Open([ref]$a, [ref]$b, [ref]$c)
}
function NewFrom($word, [string]$template) { $a = $template; $word.Documents.Add([ref]$a) }
# Word.Run right after a file opens can miss the file's macros for a moment; retry briefly.
function RunW($word, [string]$name, $arg) {
    for ($i = 1; $i -le 6; $i++) {
        try {
            if ($PSBoundParameters.ContainsKey('arg')) { $v = [string]$arg; return $word.Run($name, [ref]$v) }
            return $word.Run($name)
        }
        catch {
            if ($i -eq 6 -or $_.Exception.Message -notmatch '0x80020003') { throw }
            Start-Sleep -Milliseconds 700
        }
    }
}
function TR($word, [string]$macro) { RunW $word "SK_Test.$macro" }
function TR1($word, [string]$macro, [string]$a) { RunW $word "SK_Test.$macro" $a }

Remove-Item (Join-Path $work 'Stationkeeper Macros backups') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $sample -Force -ErrorAction SilentlyContinue
Get-ChildItem $work -Filter '~$*' -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
$excelBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$setKey = 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Macros\Settings'
$prevBackup = (Get-ItemProperty $setKey -ErrorAction SilentlyContinue).backupDir
foreach ($k in @('Team Sample Template.dotm')) { Remove-Item "HKCU:\Software\VB and VBA Program Settings\Stationkeeper Macros\$k" -Recurse -ErrorAction SilentlyContinue }

# Guard: a dialog left open by Word (a VBA error, a security question) would stall a step for ever.
# Every 15 seconds look for one; if the same Word still shows it next time, note it and end that Word.
Remove-Item (Join-Path $work 'dialog.txt') -ErrorAction SilentlyContinue
$guard = Start-Job -ScriptBlock {
    param($r)
    $seen = ''
    while ($true) {
        Start-Sleep 15
        $pidNow = Get-Content (Join-Path $r 'work\word.pid') -ErrorAction SilentlyContinue
        $out = (& (Join-Path $r 'tools\watch.ps1') -Root $r 2>&1 | Out-String)
        if ($out -match 'WINDOW:') {
            if ($seen -eq "$pidNow") {
                Add-Content (Join-Path $r 'work\dialog.txt') $out
                & (Join-Path $r 'tools\watch.ps1') -Root $r -Kill | Out-Null
                $seen = ''
            } else { $seen = "$pidNow" }
        } else { $seen = '' }
    }
} -ArgumentList $Root

try {
    # 1. Sample template: macros in Module1 and Module2, plus the test harness. No kit yet.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = $word.Documents.Add()
        $d.VBProject.Name = 'TeamSample'
        foreach ($m in 'Module1', 'Module2', 'SK_Test') { [void]$d.VBProject.VBComponents.Import((Join-Path $Root "test\$m.bas")) }
        $d.Content.Text = 'Sample template body.'
        $fmt = 15; $p = $sample; $d.SaveAs2([ref]$p, [ref]$fmt)
        $d.Close([ref]0)
    }
    Check 'sample template made' (Test-Path $sample) $sample

    # 2. Install with Word's Organizer (no VBA editor involved).
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $k = OpenDoc $word $kit $true
        $msg = RunW $word 'SK_Install.SK_InstallInto' $sample
        Check 'install via Organizer' ($msg -like 'Stationkeeper Macros is in*') ($msg -replace "`r", ' / ')
        $k.Close([ref]0)
    }

    # 4. Read the template's code: modules, comments, filtering.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        [string]$bk = Join-Path $work 'Stationkeeper Macros backups'
        $set = TR1 $word 'T_SetBackup' $bk
        Check 'backup folder chosen and kept' ($set -eq $bk) $set
        $scan = TR $word 'T_Scan'
        $ok = ($scan -like 'the VBA project | 8 | 2 |*') -and ($scan -match 'Module1.FitTablesToMargins') -and ($scan -notmatch 'HelperNotAMacro|NeedsArgs|ReturnsValue|AutoOpen|StationkeeperMacros|T_Scan')
        Check 'scan finds 8 macros in 2 modules, skips helpers/args/functions/Auto/kit' $ok $scan
        Check 'comment above Sub read' ($scan -match 'RepeatHeaderRows\{Repeats the header row') ''
        Check 'comment inside Sub read' ($scan -match 'ShowHiddenText\{Turns hidden text display on') ''
        Check 'separator lines skipped' ($scan -match 'UpdateAllFields\{Updates every field and table of contents\.\}') ''
        Check 'friendly titles' ((TR $word 'T_Friendly') -eq 'Find Broken Refs|Insert Caption Table|TOC Fix|Update TOC') (TR $word 'T_Friendly')
        $sum = TR $word 'T_ModelSummary'
        Check 'new template: all macros listed, none described' ($sum -match 'entries=8' -and $sum -match 'Not described') $sum
        $save = TR $word 'T_EditAndSave'
        Check 'catalog saves into the template' ($save -like 'Saved to Team Sample Template.dotm*') ($save -replace "`r", ' / ')
        $cat = TR $word 'T_CatalogPresent'
        Check 'catalog part present' ($cat -match 'urn:stationkeeper:macros:1') $cat
        $d.Close([ref]0)
    }

    # 5. Reopen: catalog persisted, menu search, admin, CSV round trip, shortcut + button.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $sum = TR $word 'T_ModelSummary'
        Check 'catalog survives close and reopen' (($sum -match 'In the menu') -and ($sum -match 'hist=1') -and ($sum -match 'cats=2')) $sum
        $all = TR1 $word 'T_MenuSearch' ''
        Check 'menu lists only macros put in the menu' ($all -like '3:*') $all
        $q = TR1 $word 'T_MenuSearch' 'alp'
        Check 'menu search by keyword prefix' (($q -like '3:*') -and ($q -match 'in keywords')) $q
        $q2 = TR1 $word 'T_MenuSearch' 'angle'
        Check 'menu search in description' ($q2 -match 'in description') $q2
        $q3 = TR1 $word 'T_MenuSearch' 'zzz'
        Check 'menu search with no match' ($q3 -eq '0:') $q3
        $adm = TR $word 'T_Admin'
        Check 'admin loads and edits' (($adm -match '^8 rows') -and ($adm -match 'pending')) $adm
        Check 'admin list has In menu and Description columns' (($adm -match 'Clear Direct Formatting \| Not set \| Missing') -and ($adm -match '\| Shown \| Written')) ''
        Check 'admin tabs and labels' (($adm -match 'Macros\|Categories\|Settings\|Bulk tools\|History\|') -and ($adm -match 'A=What this macro affects') -and ($adm -match 'Ch=Does this change the document\?') -and ($adm -match 'also box=none')) ''
        Check 'Save disabled with nothing to save' ($adm -match 'gate0: pending=0 unticked=False ticked=False') ''
        Check 'Save needs Ready to save ticked' ($adm -match 'gate1: pending=\d+ unticked=False ticked=True') ''
        $csv = TR $word 'T_Csv'
        Check 'CSV export and import round trip (quotes, commas, line breaks)' (($csv -match '8 row\(s\) read') -and ($csv -match '1 macro\(s\) changed: descriptions 1') -and ($csv -match 'line one\\nline two, with comma')) $csv
        $sc = TR $word 'T_Shortcut'
        Check 'shortcut and Add-ins button' (($sc -match 'Alt\+Shift\+M now opens') -and ($sc -match 'now has a Stationkeeper Macros button')) $sc
        $d.Save()
        $d.Close([ref]0)
    }

    # 6. A document based on the template (gate G3): read the catalog, run a macro, save from there.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        $cat = TR $word 'T_CatalogPresent'
        Check 'G3 catalog read from a document based on the template' ($cat -match 'urn:stationkeeper:macros:1') $cat
        $m = TR1 $word 'T_MenuSearch' 'fields'
        Check 'G3 menu works from that document' ($m -match '^\d+:') $m
        $run = TR1 $word 'T_Run' 'Module2.WriteTestMarker'
        Check 'run a macro by Module.Macro' (($run -match '^True') -and ($doc.Content.Text -like 'SKTEST*')) "$run ; doc starts '$($doc.Content.Text.Substring(0,12))'"
        $save = TR $word 'T_EditAndSave'
        Check 'save the catalog from a document based on the template' ($save -like 'Saved to*') ($save -replace "`r", ' / ')
        $doc.Close([ref]0)
    }
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $sum = TR $word 'T_ModelSummary'
        Check 'that save reached the template file' ($sum -match 'hist=2') $sum
        $d.Close([ref]0)
    }

    # 7. Macro moves to another module: entry follows by name.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $cm1 = $d.VBProject.VBComponents.Item('Module1').CodeModule
        $start = $cm1.ProcStartLine('RepeatHeaderRows', 0)
        $count = $cm1.ProcCountLines('RepeatHeaderRows', 0)
        $code = $cm1.Lines($start, $count)
        $cm1.DeleteLines($start, $count)
        $d.VBProject.VBComponents.Item('Module2').CodeModule.AddFromString($code)
        $d.Save()
        $sum = TR $word 'T_ModelSummary'
        Check 'macro moved to Module2 keeps its description' ($sum -match 'Module2\.RepeatHeaderRows\[In the menu/Tables/Module1\.RepeatHeaderRows\]') $sum
        $d.Close([ref]0)
    }


    # 7b. 0.2.0: health check, test runs, crash markers, one-step undo, needs check, test all.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        [void]$d.VBProject.VBComponents.Import([string](Join-Path $Root 'test\Module3.bas'))
        $d.Save()
        $h = TR $word 'T_Health'
        Write-Output "HEALTH: $h"
        Check 'health points to Debug > Compile' ($h -match 'Debug > Compile') ''
        Check 'no missing libraries' ($h -match 'Every library the code uses is present') ''
        Check 'health: saves, closes or deletes files' ($h -match 'Module3.SaveCopyToC\{Saves, closes or deletes files') ''
        Check 'health: hides its own errors' ($h -match 'Module3.QuietErrors\{Hides its own errors') ''
        Check 'health: alerts off and never back on' ($h -match 'Module3.AlertsOff\{Turns Word alerts off') ''
        Check 'health: fixed file path' ($h -match 'Module3.FixedPath\{Uses a fixed file path') ''
        Check 'health: clean macros have no notes' ($h -notmatch 'Module1\.|Module2\.|TwoEdits') ''
        $t1 = TR1 $word 'T_TestOne' 'Module2.WriteTestMarker'
        Check 'test run passes on a scratch document and closes it' (($t1 -match '^Passed') -and ($t1 -match 'state=Passed') -and ($t1 -match 'docs=1')) $t1
        $t2 = TR1 $word 'T_TestOne' 'Module1.Nope'
        Check 'test run of a missing macro fails' (($t2 -match '^Failed') -and ($t2 -match 'state=Failed') -and ($t2 -match 'docs=1')) $t2
        $mk = TR $word 'T_Marker'
        Check 'a run that never finished shows Failed' ($mk -match '^Failed .*stopped with an error during a test') $mk
        $u = TR $word 'T_Undo'
        Check 'one Ctrl+Z undoes a whole macro run' ($u -match 'after run \[BBB AAA base\] after one undo \[base\]') $u
        $n = TR $word 'T_Needs'
        Check 'needs check: table and selection' ($n -match 'table:False\(Put the cursor in a table first.\) sel:False\(Select some text first.\) table after:True sel after:True whole:True') $n
        $ta = TR $word 'T_TestAll'
        Check 'test all runs the safe macros and skips the file-changing one' (($ta -match 'Passed: 12') -and ($ta -match 'Failed: 0') -and ($ta -match 'ends Word\): 1') -and ($ta -match 'docs=1')) $ta
        $d.Close([ref]0)
    }

    # 7c. 0.3.0: update from lists. Excel export, primary list edited in Excel, a Word table, JSON.
    [string]$primaryX = Join-Path $work 'Macro primary list.xlsx'
    [string]$wordList = Join-Path $work 'Macro list.docx'
    [string]$jsonList = Join-Path $work 'macro list.json'
    Remove-Item $primaryX, $wordList, $jsonList -Force -ErrorAction SilentlyContinue
    $script:exported = ''
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $x = TR $word 'T_ExportXlsx'
        Check 'export the list to Excel and read it back' (($x -match 'rows=13') -and ($x -match '\.xlsx')) $x
        $script:exported = ($x -split ' \| ')[0]
        [void](TR1 $word 'T_MakeDocxList' $wordList)
        $d.Close([ref]0)
    }
    # Edit the exported list in Excel, the way a person keeping the primary list would.
    $xl = New-Object -ComObject Excel.Application
    try {
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open([string]$script:exported)
        $ws = $wb.Worksheets.Item('Macros')
        $last = $ws.UsedRange.Rows.Count
        for ($r = 2; $r -le $last; $r++) {
            $mac = [string]$ws.Cells.Item($r, 1).Text
            if ($mac -eq 'Module1.ShowHiddenText') { $ws.Cells.Item($r, 4).Value2 = 'From the primary list'; $ws.Cells.Item($r, 2).Value2 = ''; $ws.Cells.Item($r, 3).Value2 = 'View' }
            if ($mac -eq 'Module1.ClearDirectFormatting') { $ws.Cells.Item($r, 9).Value2 = 'Yes' }
        }
        $ws.Cells.Item($last + 1, 1).Value2 = 'Module9.NotHere'
        $ws.Cells.Item($last + 1, 4).Value2 = 'Not in this template'
        $wb.SaveAs($primaryX, 51)
        $wb.Close($false)
    }
    finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $json = '{"macros": [' + "`n" +
        '  {"Macro": "Module1.TOCFix", "Description": "Updates every table of contents.\nSecond line with \"quotes\" and caf\u00e9.", "Category": "Fields", "Show in menu": true},' + "`n" +
        '  {"Macro": "Gone.Macro", "Title": "Not here"},' + "`n" +
        '  {"macro": "UpdateAllFields", "keywords": "fields, toc, f9", "workson": "Whole document", "note": null, "count": 3}' + "`n" + ']}'
    [IO.File]::WriteAllText($jsonList, $json, (New-Object System.Text.UTF8Encoding($true)))
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $ux = TR1 $word 'T_UpdateFrom' $primaryX
        Write-Output "EXCEL: $ux"
        Check 'Excel list: preview counts the changes' (($ux -match 'PREVIEW 14 row\(s\) read') -and ($ux -match '2 macro\(s\) will change') -and ($ux -match 'skipped: Module9\.NotHere')) ''
        Check 'Excel list: description and category applied, blank title left alone' ($ux -match 'ShowHiddenText: title=Show Hidden Text\|cat=View\|desc=From the primary list') ''
        Check 'Excel list: Show in menu Yes applied' ($ux -match 'ClearDirectFormatting: [^#]*\|shown=1') ''
        $uw = TR1 $word 'T_UpdateFrom' $wordList
        Write-Output "WORD: $uw"
        Check 'Word table list: applied, unknown macro skipped' (($uw -match 'FitTablesToMargins: [^#]*cat=Tables\|desc=Sets every table to the full text width\.') -and ($uw -match 'skipped: NoSuch\.Macro')) ''
        $uj = TR1 $word 'T_UpdateFrom' $jsonList
        Write-Output "JSON: $uj"
        Check 'JSON list: escapes, unicode, true, bare macro name' (($uj -match 'TOCFix: [^#]*cat=Fields\|desc=Updates every table of contents\.\\nSecond line with "quotes" and caf.\.') -and ($uj -match 'TOCFix: [^#]*shown=1') -and ($uj -match 'UpdateAllFields: [^#]*kw=fields, toc, f9\|works=Whole document') -and ($uj -match 'skipped: Gone\.Macro')) ''
        $pl = TR1 $word 'T_PrimaryLabel' $primaryX
        Check 'primary list shown in the admin and its button turned on' (($pl -match 'Macro primary list\.xlsx') -and ($pl -match 'enabled=True')) $pl
        $d.Close([ref]0)
    }

    # 8. Update the kit in place (Organizer).
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $k = OpenDoc $word $kit $true
        $msg = RunW $word 'SK_Install.SK_InstallInto' $sample
        Check 'install again (update) keeps the catalog' ($msg -like 'Stationkeeper Macros is in*') ($msg -replace "`r", ' / ')
        $k.Close([ref]0)
    }
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        $sum = TR $word 'T_ModelSummary'
        Check 'catalog intact after update' ($sum -match 'hist=2') $sum
        $na = TR $word 'T_NoAccess'
        Write-Output "NO ACCESS: $na"
        Check 'no access: catalog trusted, nothing marked missing' (($na -match 'by the saved catalog') -and ($na -notmatch 'Not found')) ''
        Check 'no access: menu still lists the 3 macros' ($na -match 'menu: 3:') ''
        Check 'no access: menu search works' ($na -match 'search: [1-9]') ''
        Check 'no access: admin opens' ($na -match 'admin: 8 rows') ''
        Check 'no access: reads .bas files, one module or many' ($na -match 'bas: 8 \|') ''
        Check 'no access: comments read from .bas' (($na -match 'Module1.RepeatHeaderRows\{Repeats the header row') -and ($na -match 'ShowHiddenText\{Turns hidden text')) ''
        Check 'no access: message explains the two ways forward' (($na -match 'Trust access to the VBA project object model') -and ($na -match 'Bulk tools > Read macros from .bas files')) ''
        $d.Close([ref]0)
    }


    # 9. 0.4.0: Stationkeeper Tools. Install into the sample, catalog arrives described, standards
    # workbook, then every tool on a document with one of each defect.
    [string]$stdX = Join-Path $work 'Team standards.xlsx'
    Remove-Item $stdX -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $k = OpenDoc $word $kit $true
        $msg = RunW $word 'SK_Install.SK_InstallToolsInto' $sample
        Check 'tools install via Organizer' ($msg -like 'Stationkeeper Tools are in*(6 modules)*') ($msg -replace "`r", ' / ')
        $k.Close([ref]0)
    }
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        [void]$d.VBProject.VBComponents.Import([string](Join-Path $Root 'test\SKW_Test.bas'))
        $d.Save()
        $names = @($d.VBProject.VBComponents | ForEach-Object Name) -join ','
        Check 'tool modules are in the template' (($names -match 'SKW_Lib') -and ($names -match 'SKW_Tables') -and ($names -match 'SKW_Review')) $names
        $tc = RunW $word 'SKW_Test.TW_ToolCount'
        Write-Output "TOOLS: $tc"
        Check 'catalog: 12 tools arrive described, categorised and in the menu' (($tc -like '12 |*') -and ($tc -match 'SKW_Fields\.SafeUpdateAllFields\[[^/]*/Fields/Whole document/1\]') -and ($tc -match 'SKW_Tables\.ApplyTeamTableType\[[^/]*/Tables/Table at the cursor/1\]') -and ($tc -notmatch '/0\]|//')) ''
        $ms = TR1 $word 'T_MenuSearch' 'broken'
        Check 'menu search finds a tool' ($ms -match '^[1-9]') $ms
        $sum = TR $word 'T_ModelSummary'
        Check 'own macros keep their catalog next to the tools' ($sum -match 'hist=2') ''
        $mk = RunW $word 'SKW_Test.TW_MakeStandards' $stdX
        Write-Output "STANDARDS: $mk"
        Check 'standards workbook made and read' ((Test-Path $stdX) -and ($mk -match 'Allowed styles: 25, Table types: 2, Placeholders: 10') -and ($mk -match 'Standards: Team standards\.xlsx')) ''
        $d.Close([ref]0)
    }
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        $b = RunW $word 'SKW_Test.TW_BuildDefects'
        Check 'defect document built' ($b -match 'tables=3') $b
        $u0 = RunW $word 'SKW_Test.TW_UndoList'
        [void](RunW $word 'SKW_Test.TW_AllReports')
        $u1 = RunW $word 'SKW_Test.TW_UndoList'
        Check 'report tools leave the Undo list as it was' ($u0 -eq $u1) "before: $u0 ## after: $u1"
        $r = RunW $word 'SKW_Test.TW_Broken'
        Write-Output "BROKEN: $r"
        Check 'broken references: gone bookmark, missing file, link to nowhere; good ref not listed' (($r -like '3 broken*') -and ($r -match 'bookmark Gone') -and ($r -match 'missing\.docx') -and ($r -match 'NoSuchPlace') -and ($r -notmatch 'bookmark Intro')) ''
        $r = RunW $word 'SKW_Test.TW_FieldReport'
        Check 'field code report lists REF, SEQ, TOC, INCLUDETEXT, PAGE' (($r -match 'REF \d') -and ($r -match 'SEQ \d') -and ($r -match 'TOC \d') -and ($r -match 'INCLUDETEXT \d') -and ($r -match 'PAGE \d')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Leftovers'
        Write-Output "LEFTOVERS: $($r.Substring(0, [Math]::Min(400, $r.Length)))"
        Check 'leftovers: comment, tracked change, hidden text, highlight, two placeholders' ($r -match '^Comments 1, tracked changes [1-9]\d*, hidden text 1, highlights 1, placeholders 2\.') ''
        $r = RunW $word 'SKW_Test.TW_Inspector'
        Check 'style inspector shows the hand-set font and size' (($r -match 'Paragraph style: Normal') -and ($r -match 'font Arial \(style: ') -and ($r -match 'size 14')) $r
        $r = RunW $word 'SKW_Test.TW_NonAllowed'
        Check 'non-allowed styles: Fancy Body listed, Heading 1 fine' (($r -like '1 style(s) not allowed, 0 required*') -and ($r -match 'Paragraph style Fancy Body used 1') -and ($r -notmatch 'style Heading 1 used')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $prev = RunW $word 'SKW_Test.TW_NoStandards'
        $r = RunW $word 'SKW_Test.TW_NonAllowed'
        Check 'without a workbook: custom styles listed, built-in ones allowed' (($r -match 'Custom style \| Paragraph style Fancy Body') -and ($r -match 'built-in Word styles count as allowed')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        [void](RunW $word 'SKW_Test.TW_UseStandards' $prev)
        $r = RunW $word 'SKW_Test.TW_Direct'
        Write-Output "DIRECT: $($r.Substring(0, [Math]::Min(600, $r.Length)))"
        Check 'direct formatting: Arial and extra space found, bold and subscript treated as emphasis' (($r -match 'font Arial') -and ($r -match 'space after 30') -and ($r -notmatch 'Good paragraph')) ''
        $r = RunW $word 'SKW_Test.TW_Clear'
        Write-Output "CLEAR: $r"
        Check 'clear (selection): back to the style, bold kept, one Ctrl+Z restores' (($r -match 'Kept 1 piece') -and ($r -notmatch 'font=Arial size') -and ($r -match 'kept=-1') -and ($r -notmatch ' after=30 ##') -and ($r -match 'undo: font=Arial after=30')) ''
        $r = RunW $word 'SKW_Test.TW_Coauthor'
        Check 'co-authored file: changes blocked, reports still run' (($r -match 'clear: This document is on SharePoint or OneDrive with AutoSave on') -and ($r -match 'report: 3 broken')) $r
        $r = RunW $word 'SKW_Test.TW_Tables'
        Write-Output "TABLES: $($r.Substring(0, [Math]::Min(900, $r.Length)))"
        Check 'table audit: good table clean, bad and floating tables flagged' (($r -like '3 table(s) checked, 2 with findings*') -and ($r -match 'Header row does not repeat') -and ($r -match 'Rows can break') -and ($r -match 'Wider than the margins') -and ($r -match 'Floating') -and ($r -match 'Style Table Normal is not a team table type') -and ($r -notmatch '\| 1 \| \d+ \|')) ''
        $r = RunW $word 'SKW_Test.TW_ApplyType'
        Write-Output "APPLY: $($r.Substring(0, [Math]::Min(300, $r.Length))) ... $($r.Substring([Math]::Max(0, $r.Length - 200)))"
        $r = $r + ' ## ' + (RunW $word 'SKW_Test.TW_TableState') + ' ## audit: ' + (RunW $word 'SKW_Test.TW_Tables')
        $ul = RunW $word 'SKW_Test.TW_UndoList'
        $u = 1; [void]$doc.Undo([ref]$u)
        $st = (RunW $word 'SKW_Test.TW_TableState') + " (undo list was: $ul)"
        Check 'apply table type: style, header repeats, rows whole, full width; one Ctrl+Z restores' (($r -match 'style=Table Grid head=-1 break=0 widthType=2 width=100') -and ($st -like 'style=Table Normal head=0 break=-1 widthType=3 width=* (undo list was: * | Stationkeeper: Apply table type |*')) "after one undo: $st"
        Check 'after applying, only the caption is left on that table' (($r -match '\| 2 \| \d+ \| Caption') -and ($r -notmatch '\| 2 \| \d+ \| (Header row|Rows break|Width|Table style)')) ''
        $r = RunW $word 'SKW_Test.TW_SafeUpdate'
        Write-Output "UPDATE: $r"
        Check 'safe update: included file kept, broken ref now shows, tracking restored' (($r -match 'kept as they were') -and ($r -match 'include same=True locked=False brokenref=Error!') -and ($r -match 'track=True') -and ($r -match 'Track changes was paused')) ''
        $r = RunW $word 'SKW_Test.TW_MenuRun'
        Check 'a report tool runs from the menu engine' ($r -match 'ran=True newdocs=1 state=Passed') $r
        $r = RunW $word 'SKW_Test.TW_TestAll'
        Write-Output "TESTALL: $r"
        Check 'test every macro with the tools in: none fail, reports closed, quiet flag cleared' (($r -match 'Failed: 0') -and ($r -match 'docs before=(\d+) after=\1') -and ($r -match 'quiet flag=$')) ''
        $doc.Close([ref]0)
    }
}
catch {
    Check 'test run' $false ($_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}
finally {
    Stop-Job $guard -ErrorAction SilentlyContinue; Remove-Job $guard -Force -ErrorAction SilentlyContinue
    if (Test-Path (Join-Path $work 'dialog.txt')) { Check 'no dialog stopped a step' $false ((Get-Content (Join-Path $work 'dialog.txt') -Raw) -replace '\s+', ' ') }
    # End only Excel started by automation during this run (never one a person opened).
    Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue |
        Where-Object { ($excelBefore -notcontains $_.ProcessId) -and ($_.CommandLine -match '/automation|-Embedding') } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-ItemProperty 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Tools\Standards' -Name 'Team Sample Template.dotm' -ErrorAction SilentlyContinue
    Remove-ItemProperty 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Tools\Run' -Name 'quiet' -ErrorAction SilentlyContinue
    if ($null -ne $prevBackup) { Set-ItemProperty $setKey -Name backupDir -Value $prevBackup }
    else { Remove-ItemProperty $setKey -Name backupDir -ErrorAction SilentlyContinue }
    Write-Output ''
    Write-Output ("RESULT: {0} checks, {1} failed." -f $results.Count, $fails)
    $results | Set-Content (Join-Path $work 'test-results.txt')
}
