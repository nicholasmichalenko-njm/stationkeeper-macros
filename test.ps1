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
$pptBefore = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
$setKey = 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Macros\Settings'
$prevBackup = (Get-ItemProperty $setKey -ErrorAction SilentlyContinue).backupDir
$linkKey = 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Tools\Link'
$prevLink = (Get-ItemProperty $linkKey -ErrorAction SilentlyContinue).folder
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
        Check 'tools install via Organizer' ($msg -like 'Stationkeeper Tools are in*(15 modules)*') ($msg -replace "`r", ' / ')
        $k.Close([ref]0)
    }
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        [void]$d.VBProject.VBComponents.Import([string](Join-Path $Root 'test\SKW_Test.bas'))
        $d.Save()
        $names = @($d.VBProject.VBComponents | ForEach-Object Name) -join ','
        Check 'tool modules are in the template' (($names -match 'SKW_Lib') -and ($names -match 'SKW_Tables') -and ($names -match 'SKW_Review') -and ($names -match 'SKW_Layout') -and ($names -match 'SKW_Captions') -and ($names -match 'SKW_Lists')) $names
        $tc = RunW $word 'SKW_Test.TW_ToolCount'
        Write-Output "TOOLS: $tc"
        Check 'catalog: 63 tools arrive described, categorised and in the menu' (($tc -like '63 |*') -and ($tc -match 'SKW_Fields\.SafeUpdateAllFields\[[^/]*/Fields/Whole document/1\]') -and ($tc -match 'SKW_Tables\.ApplyTeamTableType\[[^/]*/Tables/Table at the cursor/1\]') -and ($tc -notmatch '/0\]|//')) ''
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

    # 10. 0.5.0: batch 2 on a fresh defect document with the batch 2 defects added.
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        [void](RunW $word 'SKW_Test.TW_BuildDefects')
        $b = RunW $word 'SKW_Test.TW_BuildDefects2'
        Check 'batch 2 defects added' ($b -match 'sections=2 pictures=2') $b
        $u0 = RunW $word 'SKW_Test.TW_UndoList'
        [void](RunW $word 'SKW_Test.TW_AllReports')
        $u1 = RunW $word 'SKW_Test.TW_UndoList'
        Check 'all 18 report tools leave the Undo list as it was' ($u0 -eq $u1) "before: $u0 ## after: $u1"
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_SectionMap'
        Write-Output "SECTIONS: $($r.Substring(0, [Math]::Min(900, $r.Length)))"
        Check 'section map: two sections, the second landscape with its own footer' (($r -like '2 section(s) mapped*') -and ($r -match 'Landscape') -and ($r -match 'same as previous / own')) ''
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_CaptionAudit'
        Write-Output "CAPTIONS: $($r.Substring(0, [Math]::Min(900, $r.Length)))"
        Check 'caption audit: two tables without captions, one typed number, good table and figure fine' ($r -like 'No caption 2, wrong side 0, wrong label 0, typed numbers 1, repeated 0*') ''
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_KeepWithNextAudit'
        Write-Output "KEEP: $($r.Substring(0, [Math]::Min(600, $r.Length)))"
        Check 'keep with next: the loose heading is listed' (($r -like 'Headings 1, captions [01], long chains 0*') -and ($r -match 'Loose heading')) ''
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_BreakHygiene'
        Write-Output "BREAKS: $($r.Substring(0, [Math]::Min(600, $r.Length)))"
        Check 'break hygiene: one page break (not the section break), one run of empty paragraphs, one line break' ($r -like 'Page breaks 1, column breaks 0, empty runs 1, line breaks 1, page break before 0*') ''
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_TypedListsFinder'
        Check 'typed lists: bullet, number and letter found; the real list is not' (($r -like '3 typed list*') -and ($r -match 'Typed bullet') -and ($r -match 'Typed number') -and ($r -match 'Typed letter') -and ($r -notmatch 'Real list item')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_TOCCheck'
        Write-Output "TOC: $($r.Substring(0, [Math]::Min(600, $r.Length)))"
        Check 'TOC check: the two headings added later are missing' (($r -like 'Missing 2, no heading 0, old page numbers 0*') -and ($r -match 'Heading added later') -and ($r -match 'Loose heading')) ''
        $r = RunW $word 'SKW_Test.TW_Toggle'
        Check 'toggle field codes on and back off' ($r -match '^Field codes shown.* / Field results shown\. / now=False') $r
        $r = RunW $word 'SKW_Test.TW_Banned'
        Write-Output "BANNED: $r"
        Check 'banned phrases: two replaced as tracked changes, one flagged, a comment each, tracking put back' (($r -match 'Changed 2 as tracked changes and flagged 1') -and ($r -match 'comments\+3') -and ($r -match 'track=False')) ''
        $u = 1; [void]$doc.Undo([ref]$u)
        $r2 = RunW $word 'SKW_Test.TW_Report' 'SKW_PreDeliveryLeftovers'
        Check 'banned phrases: one Ctrl+Z takes the changes and comments back out' ($r2 -match '^Comments 1, tracked changes 1,') ($r2.Substring(0, [Math]::Min(120, $r2.Length)))
        $r = RunW $word 'SKW_Test.TW_Remap'
        $c1 = RunW $word 'SKW_Test.TW_StyleCount' 'Fancy Body'
        [void]$doc.Undo([ref]$u)
        $c2 = RunW $word 'SKW_Test.TW_StyleCount' 'Fancy Body'
        Check 'remap style: Fancy Body to Normal, a missing style skipped, one Ctrl+Z restores' (($r -match 'Remapped 1 place') -and ($r -match 'No Such Style is not in this document') -and ($c1 -eq '0') -and ($c2 -eq '1')) "$r ## after=$c1 undo=$c2"
        $s0 = RunW $word 'SKW_Test.TW_SetNormalSize' '20'
        $r = RunW $word 'SKW_Test.TW_Import'
        $s1 = RunW $word 'SKW_Test.TW_NormalSize'
        [void]$doc.Undo([ref]$u)
        $s2 = RunW $word 'SKW_Test.TW_NormalSize'
        Write-Output "IMPORT: $r ## size set=$s0 after import=$s1 after one undo=$s2"
        Check 'import team styles: Normal comes back from the template and is listed' (($r -match 'changed: [^.]*Normal') -and ($s1 -ne '20')) "$r ## after=$s1"
        $script:importUndo = $s2
        $doc.Close([ref]0)
    }

    # 11. 0.6.0: batch 3 on a saved defect document (clean copy and PDF need a file).
    [string]$d6 = Join-Path $work 'Defects 0.6.docx'
    [string]$c6 = Join-Path $work 'Defects 0.6 - clean.docx'
    [string]$p6 = Join-Path $work 'Defects 0.6.pdf'
    Remove-Item $d6, $c6, $p6 -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        [void](RunW $word 'SKW_Test.TW_BuildDefects')
        [void](RunW $word 'SKW_Test.TW_BuildDefects2')
        $b = RunW $word 'SKW_Test.TW_BuildDefects3' $d6
        Check 'batch 3 defects added and saved' (Test-Path $d6) $b
        $u0 = RunW $word 'SKW_Test.TW_UndoList'
        [void](RunW $word 'SKW_Test.TW_AllReports')
        $u1 = RunW $word 'SKW_Test.TW_UndoList'
        Check 'all 18 report tools leave the Undo list as it was (batch 3 document)' ($u0 -eq $u1) "before: $u0 ## after: $u1"
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_AcronymQuickCheck'
        Write-Output "ACRONYMS: $($r.Substring(0, [Math]::Min(1500, $r.Length)))"
        Check 'acronym quick check: labelled text only, RFP spelled out and on the list, SOW spelled out only later' (($r -match 'text only') -and ($r -match 'pictures and images are not scanned') -and ($r -match '\| RFP \| 1 \| \d+ \| Yes \| Yes \| Request for Proposal \|') -and ($r -match '\| SOW \| 2 \| \d+ \| No \(spelled out later\) \| No \|')) ''
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_TypographySweep'
        Check 'typography sweep: double space, space before punctuation, straight quotes, double hyphen, three dots, space at paragraph end' (($r -match 'Double space [1-9]') -and ($r -match 'Space before punctuation [1-9]') -and ($r -match 'Straight quote [1-9]') -and ($r -match 'Double hyphen [1-9]') -and ($r -match 'Three dots [1-9]') -and ($r -match 'Space at paragraph end [1-9]')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Spelling'
        Check 'check selection spelling: the two misspellings with suggestions' (($r -match '^Spelling [2-9]') -and ($r -match 'recieve \| receive')) ($r.Substring(0, [Math]::Min(400, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Dictionaries'
        Check 'team dictionaries: says the list is empty' ($r -match 'lists no dictionaries') $r
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_FontNotAllowed'
        Check 'fonts: Arial and Courier New not allowed' (($r -match 'Not allowed \| Arial') -and ($r -match 'Not allowed \| Courier New')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_PageLimitCounter'
        Check 'page limit counter: parts by top-level heading and the whole document' (($r -match 'page\(s\), 0 over the limit') -and ($r -match '\| Introduction \| 1 \|') -and ($r -match '\| Heading added later \|') -and ($r -match '\| Whole document \| 1 \|')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Report' 'SKW_GraphicsAudit'
        Check 'graphics audit: two pictures, one without alt text' ($r -like '2 graphic(s): no alt text 1,*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
        $r = RunW $word 'SKW_Test.TW_Clean' $c6
        Check 'metadata clean copy: author and company gone, the open document untouched' ((Test-Path $c6) -and ($r -match 'props=\[\|\]') -and ($r -match 'active=Defects 0\.6\.docx') -and ($r -match 'comment')) $r
        $r = RunW $word 'SKW_Test.TW_Pdf' $p6
        Check 'PDF with bookmarks saved' ((Test-Path $p6) -and ($r -match 'bytes=\d{4,}')) $r
        $r = RunW $word 'SKW_Test.TW_RemoveUnused'
        $e1 = RunW $word 'SKW_Test.TW_StyleExists' 'Unused Custom'
        $u = 1; [void]$doc.Undo([ref]$u)
        $e2 = RunW $word 'SKW_Test.TW_StyleExists' 'Unused Custom'
        $fb = RunW $word 'SKW_Test.TW_StyleExists' 'Fancy Body'
        Check 'remove unused custom styles: both unused ones go, Fancy Body stays, one Ctrl+Z brings them back' (($r -match 'Deleted 2 unused') -and ($r -match 'Unused Custom') -and ($r -match 'Unused Char') -and ($r -notmatch 'Fancy Body') -and ($e1 -eq 'False') -and ($e2 -eq 'True') -and ($fb -eq 'True')) "$r ## after=$e1 undo=$e2"
        $doc.Close([ref]0)
    }

    # 12. 0.7.0: utilities.
    [string]$d7 = Join-Path $work 'Defects 0.7.docx'
    Remove-Item $d7 -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        [void](RunW $word 'SKW_Test.TW_BuildDefects')
        [void](RunW $word 'SKW_Test.TW_BuildDefects2')
        [void](RunW $word 'SKW_Test.TW_BuildDefects3' $d7)
        $b = RunW $word 'SKW_Test.TW_BuildLists'
        $r = RunW $word 'SKW_Test.TW_Numbering'
        Write-Output "NUMBERING: $b ## $r"
        Check 'continue numbering makes B1 3; restart at B2 makes it 1' (($r -match 'continue: .* B1=3\. B2=4\.') -and ($r -match 'restart: .* B1=3\. B2=1\.')) "$b ## $r"
        $r = RunW $word 'SKW_Test.TW_Fields' 'lock'
        $lk = RunW $word 'SKW_Test.TW_FieldLocks'
        Check 'lock fields in the selection' (($r -match '^2 field\(s\) locked.*before=2 after=2') -and ($lk -eq 'LL')) "$r ## $lk"
        $r = RunW $word 'SKW_Test.TW_Fields' 'unlock'
        $lk = RunW $word 'SKW_Test.TW_FieldLocks'
        Check 'unlock fields in the selection' (($r -match '^2 field\(s\) unlocked') -and ($lk -eq 'uu')) "$r ## $lk"
        $r = RunW $word 'SKW_Test.TW_Fields' 'unlink'
        Check 'unlink fields: the two fields become plain text' ($r -match '^2 field\(s\) turned into plain text.*before=2 after=0') $r
        $u = 1; [void]$doc.Undo([ref]$u)
        $r2 = RunW $word 'SKW_Test.TW_Fields' 'lock'
        Check 'unlink: one Ctrl+Z brings the fields back' ($r2 -match 'before=2') $r2
        $r = RunW $word 'SKW_Test.TW_Convert'
        Write-Output "CONVERT: $r"
        Check 'convert to team list: markers gone, bullet to List Bullet, number and letter to List Number' (($r -match 'Converted 1 bullet\(s\) to List Bullet and 2 numbered') -and ($r -match 'List Bullet:Typed bullet one:') -and ($r -match 'List Number:Typed number one:') -and ($r -match 'List Number:Typed letter one:')) ''
        $w0 = RunW $word 'SKW_Test.TW_TableWidth' '2'
        $r = RunW $word 'SKW_Test.TW_Fit'
        $r = $r + ' ## width2=' + (RunW $word 'SKW_Test.TW_TableWidth' '2')
        [void]$doc.Undo([ref]$u)
        $w2 = RunW $word 'SKW_Test.TW_TableWidth' '2'
        Check 'fit all tables: the 9-inch table fits the 6.5-inch text width; one Ctrl+Z restores it' (($w0 -eq '9.00') -and ($r -match 'Fitted 1 of 3') -and ($r -match 'width2=6\.[45]\d') -and ($w2 -eq '9.00')) "before=$w0 $r undo=$w2"
        $r = RunW $word 'SKW_Test.TW_NextFinding'
        Write-Output "NEXT: $r"
        Check 'go to next finding walks the last report in the document' (($r -match '1: Finding 1 of \d+') -and ($r -match '2: Finding 2 of \d+') -and ($r -match 'sel=Bad head')) ''
        $r = RunW $word 'SKW_Test.TW_InsertAcronyms'
        $t1 = RunW $word 'SKW_Test.TW_TableCount'
        [void]$doc.Undo([ref]$u)
        $t2 = RunW $word 'SKW_Test.TW_TableCount'
        Check 'insert acronym list: one table with RFP and its meaning; one Ctrl+Z takes it out' (($r -match 'Inserted 1 acronym') -and ($r -match 'tables\+1') -and ($r -match 'RFP') -and ($r -match 'Request for Proposal') -and ([int]$t2 -eq ([int]$t1 - 1))) "$r ## undo tables $t1 to $t2"
        $r = RunW $word 'SKW_Test.TW_Reset'
        [void]$doc.Undo([ref]$u)
        $f2 = RunW $word 'SKW_Test.TW_FontName' 'typed in Arial'
        Check 'reset paragraph to style: Arial and bold gone; one Ctrl+Z restores' (($r -match '1 paragraph\(s\) back to their style') -and ($r -notmatch 'font=Arial') -and ($r -match 'kept=0') -and ($f2 -eq 'Arial')) "$r ## undo font=$f2"
        $r = RunW $word 'SKW_Test.TW_Shortcuts'
        Check 'tool shortcuts set in the template' (($r -match 'Ctrl\+Alt\+Shift\+N: next finding') -and ($r -match 'N=.*GoToNextFinding')) $r
        $doc.Close([ref]0)
    }


    # 13. 0.8.0: PowerPoint tools on a deck the test makes (handed to the tools directly).
    [string]$deck = Join-Path $work 'Deck 0.8.pptx'
    [string]$deckCopy = Join-Path $work 'Deck 0.8 - delivery.pptx'
    Remove-Item $deck, $deckCopy -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        try {
            $m = RunW $word 'SKW_Test.TW_MakeDeck' $deck
            Check 'test deck made' ((Test-Path $deck) -and ($m -match 'slides=4')) $m
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckTextLeftovers'
            Check 'deck leftovers: notes, hidden slide, placeholder' ($r -like 'Notes 1, comments 0, hidden slides 1, placeholders [1-9]*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckBannedPhrases'
            Check 'deck banned phrases: utilize on slide 1' (($r -match '^[1-9]\d* banned') -and ($r -match '\| 1 \| Slide \| utilize \| use \|')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckAcronymCheck'
            Check 'deck acronyms: RFP spelled out on slide 1 and on the team list' (($r -match 'text only') -and ($r -match '\| RFP \| 2 \| 1 \| slide 1 \| Yes \| Request for Proposal \|')) ($r.Substring(0, [Math]::Min(500, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckFontAudit'
            Check 'deck fonts: Arial not allowed' ($r -match 'Not allowed \| Arial') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckSmallText'
            Check 'deck small text: the 8 pt footnote' (($r -like '1 piece(s) of text under 10 pt*') -and ($r -match 'Tiny Arial footnote')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckAltTextAudit'
            Check 'deck alt text: the picture without it' ($r -like '1 graphic(s), 1 without alt text*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckOffSlideObjects'
            Check 'deck off-slide objects: the parked box' (($r -like '1 shape(s) off the slide*') -and ($r -match 'Wholly off the slide')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckTitlesCheck'
            Check 'deck titles: one slide without, one sharing' ($r -like '1 slide(s) without a title, 1 sharing one*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Deck' 'SKP_DeckEmptyPlaceholders'
            Check 'deck empty placeholders: the empty body' ($r -like '1 empty placeholder*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_DeckCopy' $deckCopy
            Check 'deck delivery copy: hidden slide and notes gone, open deck unchanged' ((Test-Path $deckCopy) -and ($r -match '1 hidden slide') -and ($r -match 'notes on 1 slide') -and ($r -match 'copySlides=3 openSlides=4')) $r
        }
        finally { [void](RunW $word 'SKW_Test.TW_CloseDeck') }
        $doc.Close([ref]0)
    }

    # 14. 0.9.0: Excel tools on a workbook the test makes (handed to the tools directly).
    [string]$book = Join-Path $work 'Book 0.9.xlsx'
    [string]$bookCopy = Join-Path $work 'Book 0.9 - delivery.xlsx'
    Remove-Item $book, $bookCopy -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        try {
            $m = RunW $word 'SKW_Test.TW_MakeBook' $book
            Check 'test workbook made' ((Test-Path $book) -and ($m -match 'sheets=3')) $m
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookFormulaErrors'
            Check 'book formula errors: #DIV/0! and #N/A' (($r -like '2 cell(s) with an error*') -and ($r -match '#DIV/0!') -and ($r -match '#N/A')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookBrokenNames'
            Check 'book broken names: BadName' (($r -like '1 broken name*') -and ($r -match 'BadName')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookExternalLinks'
            Check 'book external links: none' ($r -like '0 linked workbook(s), 0 cell(s)*') ($r.Substring(0, [Math]::Min(200, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookHiddenContent'
            Check 'book hidden content: the hidden sheet and the hidden row' ($r -like 'Hidden sheets 1, sheets with hidden rows or columns 1,*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookLeftovers'
            Check 'book leftovers: the note and the TBD cell' ($r -like 'Notes and comments 1, placeholders 1*') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookBannedPhrases'
            Check 'book banned phrases: utilize in A2' (($r -like '1 banned phrase*') -and ($r -match '\| Pricing \| A2 \| utilize \| use \|')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookAcronymCheck'
            Check 'book acronyms: RFP spelled out and on the team list' (($r -match 'text only') -and ($r -match '\| RFP \| 1 \| Pricing!A2 \| Yes \| Yes \| Request for Proposal \|')) ($r.Substring(0, [Math]::Min(400, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookFontAudit'
            Check 'book fonts: Arial not allowed' ($r -match 'Not allowed \| Arial') ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_Book' 'SKX_BookMergedCells'
            Check 'book merged cells: A10:B10' (($r -like '1 merged area*') -and ($r -match 'A10:B10')) ($r.Substring(0, [Math]::Min(300, $r.Length)))
            $r = RunW $word 'SKW_Test.TW_BookCopy' $bookCopy
            Check 'book delivery copy: note gone in the copy, still in the open workbook' ((Test-Path $bookCopy) -and ($r -match 'copyComments=0 openComments=1')) $r
        }
        finally { [void](RunW $word 'SKW_Test.TW_CloseBook') }
        $doc.Close([ref]0)
    }
    # 15. 0.10.0: three kinds of macro, filters, badges, the two questions, the link strip and its states.
    [string]$effort = Join-Path $work 'Effort A'
    Remove-Item $effort -Recurse -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $d = OpenDoc $word $sample
        try {
            $su = TR $word 'T_SkSetup'
            Check 'mark built, needs rules, needs template, not related; save' (($su -match 'built=Module1\.ShowHiddenText') -and ($su -match 'needs=Module2\.WriteTestMarker') -and ($su -match 'needsT=Module1\.UpdateAllFields') -and ($su -match '\| Saved to')) $su
            $k = TR $word 'T_SkKinds'
            Check 'kinds kept in the catalog; tools prefilled (built or not related)' (($k -match 'ShowHiddenText=Built for Stationkeeper;') -and ($k -match 'WriteTestMarker=Needs Stationkeeper data: rules;') -and ($k -match 'StyleInspector=Built for Stationkeeper;') -and ($k -match 'ToggleFieldCodes=Not related;') -and ($k -match 'BookFontAudit=Built for Stationkeeper;') -and ($k -match 'DeckDeliveryCopy=Not related;') -and ($k -match 'xml has stationkeeper=True')) $k

            $l = TR1 $word 'T_SkLink' ''
            Check 'link: not linked' ($l -like 'Not linked*') $l
            [void](TR1 $word 'T_SkAskPref' '1')
            $mn = TR $word 'T_SkMenu'
            Write-Output "MENU0: $mn"
            Check 'menu link strip: not linked, Link a folder, legend' (($mn -match 'LINK Stationkeeper data: not linked / Not linked\.') -and ($mn -match 'button=Link a folder\.\.\. unlink=False') -and ($mn -match 'legend=<o> SK  built for Stationkeeper, works alone +<f> SK  needs Stationkeeper data')) ''
            Check 'filter Needs effort data: the two, with the filled badge' ($mn -match 'NEEDS EFFORT DATA 2: \[[^\]]+ SK\+\] \[[^\]]+ SK\+\] \|\|') ''
            Check 'filter Stationkeeper tools: tools with outlined badge, own macros left out' (($mn -match 'STATIONKEEPER TOOLS \d+:[^|]*\[Style Inspector SK-\]') -and ($mn -match 'STATIONKEEPER TOOLS \d+:[^|]*\[Toggle Field Codes\]') -and ($mn -notmatch 'STATIONKEEPER TOOLS [^|]*Fit Tables To Margins')) ''
            Check 'filter My macros: own macros only' (($mn -match 'MY MACROS \d+:[^|]*\[Fit Tables To Margins\]') -and ($mn -notmatch 'MY MACROS [^|]*(SK-|SK\+|Style Inspector)')) ''
            Check 'filter Ready to run now: needs-data macros left out when not linked' (($mn -match 'READY TO RUN NOW \d+:') -and ($mn -notmatch 'READY TO RUN NOW [^|]*SK\+')) ''
            Check 'built, not linked: the gentle question' ($mn -match 'GATE built: ask \| Built for Stationkeeper \| This is built for Stationkeeper but is still functional\. Would you like to run it\?[^|]*\| buttons: Run it, Link a folder\.\.\., Cancel \| Don''t ask again') ''
            Check 'needs data, not linked: says what and where, offers the picker' ($mn -match 'GATE needs: needs \| Needs Stationkeeper data \| This macro is used with the Stationkeeper application\. Please point to the Stationkeeper data folder for your effort\.[^|]*It needs the effort rules\.[^|]*Stationkeeper Data\\rules\\rules\.json[^|]*Right now: Not linked\.[^|]*\| buttons: Point to the folder\.\.\., Cancel \|\|') ''
            Check 'not related: runs with no question' ($mn -match 'GATE none: run') ''
            $t1 = TR1 $word 'T_SkTestOne' 'Module2.WriteTestMarker'
            Check 'test run skips a needs-data macro when not linked' ($t1 -like 'Skipped: it needs Stationkeeper data. Not linked*') $t1
            [void](TR1 $word 'T_SkAskPref' '0')
            $mn = TR $word 'T_SkMenu'
            Check 'Don''t ask again: built runs with no question' ($mn -match 'GATE built: run') ''
            [void](TR1 $word 'T_SkAskPref' '1')

            $l = TR1 $word 'T_SkLink' (Join-Path $work 'No Such Effort')
            Check 'link: folder missing' ($l -like 'Folder missing*') $l
            $l = TR1 $word 'T_SkLink' $work
            Check 'link: no data yet' ($l -like 'No data yet*') $l
            $md = TR1 $word 'T_SkMakeData' "$effort|current"
            $l = TR1 $word 'T_SkLink' $effort
            Check 'link: current, rules present, effort named' (($l -like 'Current | parts=rules | rev=15 | effort=Effort A*') -and ($l -match 'by tester')) "$l ;; $md"
            $l2 = TR1 $word 'T_SkLink' (Join-Path $effort 'Stationkeeper Data')
            Check 'link: picking the Stationkeeper Data folder itself works too' ($l2 -like 'Current*') $l2
            [void](TR1 $word 'T_SkLink' $effort)
            $mn = TR $word 'T_SkMenu'
            Write-Output "MENU1: $mn"
            Check 'menu link strip: current, Change, Unlink' (($mn -match 'LINK Stationkeeper data: Current  \|  ') -and ($mn -match 'Effort A, revision 15, saved') -and ($mn -match 'button=Change\.\.\. unlink=True')) ''
            Check 'linked: built and needs-rules run with no question' (($mn -match 'GATE built: run') -and ($mn -match 'GATE needs: run')) ''
            Check 'linked, part missing: needs-template says the template is missing' ($mn -match 'GATE needsT: needs \| Needs Stationkeeper data \| [^|]*Right now: This effort''s Stationkeeper Data has no effort template yet\.') ''
            Check 'Ready to run now: needs-rules in, needs-template out' (($mn -match 'READY TO RUN NOW [^|]*\[Write Test Marker SK\+\]') -and ($mn -notmatch 'READY TO RUN NOW [^|]*\[Update All Fields')) ''
            [void](TR1 $word 'T_SkMakeData' "$effort|incomplete")
            $l = TR1 $word 'T_SkLink' $effort
            Check 'link: incomplete when rules.json is a revision behind' (($l -like 'Incomplete*') -and ($l -match 'rules/rules\.json \(revision 14\)')) $l
            [void](TR1 $word 'T_SkMakeData' "$effort|stale")
            $l = TR1 $word 'T_SkLink' $effort
            Check 'link: stale after 72 hours' (($l -like 'Stale*') -and ($l -match 'saved 5 days ago')) $l
            [void](TR1 $word 'T_SkMakeData' "$effort|bad")
            $l = TR1 $word 'T_SkLink' $effort
            Check 'link: a manifest that is not Stationkeeper''s' ($l -like 'Unreadable*') $l
        }
        finally { [void](TR1 $word 'T_SkLink' '') }
        $d.Close([ref]0)
    }

    # 16. 1.0.0: the tools read the linked effort's rules first, keep a local copy, and fall back to it.
    [string]$effB = Join-Path $work 'Effort B'
    Remove-Item $effB, "$effB moved" -Recurse -Force -ErrorAction SilentlyContinue
    Invoke-WordJob -Root $Root -Job {
        param($word)
        $doc = NewFrom $word $sample
        try {
            [void](TR $word 'T_SkClearCache')
            [void](TR1 $word 'T_SkLink' '')
            $r = RunW $word 'SKW_Test.TW_Rules'
            Check 'not linked: the tools read the standards workbook as before' (($r -match '^note=Standards: Team standards\.xlsx \|') -and ($r -match 'caption=Caption \|') -and ($r -match 'banned=utilize,in order to,and/or \|') -and ($r -match 'placeholders=10 \|')) $r
            [void](TR1 $word 'T_SkPublish' "$effB|3|3|Effort Caption")
            [void](TR1 $word 'T_SkLink' $effB)
            $r = RunW $word 'SKW_Test.TW_Rules'
            Write-Output "RULES3: $r"
            Check 'linked: the effort''s rules come first, stamped on reports' ($r -match '^note=Rules: effort Effort B rev 3, then Team standards\.xlsx \|') $r
            Check 'effort settings win one by one; unset ones come from the workbook' (($r -match 'caption=Effort Caption \|') -and ($r -match 'table=Table \|') -and ($r -match 'firstrow=Yes \|')) ''
            Check 'an effort list replaces the sheet, even an empty one' (($r -match 'banned=leverage \|') -and ($r -match 'placeholders=0 \|')) ''
            Check 'a list the effort does not set stays from the workbook' ($r -match 'acronyms=1 \|') ''
            Check 'the effort''s template and dictionary come from the local copy' (($r -match 'source=[^|]*\\Effort cache\\r3\\templates\\effort\.dotx \|') -and ($r -match 'dic=[^|]*\\Effort cache\\r3\\dictionaries\\effort\.dic$')) ''
            $c = RunW $word 'SKW_Test.TW_Cache'
            Check 'local copy: one revision folder' ($c -eq 'folders: r3 | root=r3') $c
            $b = RunW $word 'SKW_Test.TW_EffortBanned'
            Write-Output "BANNED: $($b.Substring(0, [Math]::Min(500, $b.Length)))"
            Check 'tools follow the effort: its banned phrase caught, the workbook''s not; reports stamped' (($b -match 'against 1 banned phrase') -and ($b -match 'comment=[^#]*EF-1') -and ($b -notmatch 'PL-0') -and ($b -match 'report=[^#]*Rules: effort Effort B rev 3')) ''

            [void](TR1 $word 'T_SkPublish' "$effB|4|3|Caption Four")
            $r = RunW $word 'SKW_Test.TW_Rules'
            Check 'mid-save: the last complete copy is used, and the report says so' (($r -match '^note=Rules: effort Effort B rev 3 \(last complete copy: incomplete\)') -and ($r -match 'caption=Effort Caption \|')) $r
            $lf = TR $word 'T_SkLinkFiles'
            Check 'menu: incomplete, but needs-data macros can use the copy' (($lf -like 'Incomplete | cached=True | parts=rules,template,dictionary | rules file=*\Effort cache\r3\rules\rules.json |*') -and ($lf -match 'last complete copy \(revision 3\)')) $lf
            [void](TR1 $word 'T_SkPublish' "$effB|5|5|Effort Caption 5")
            $r = RunW $word 'SKW_Test.TW_Rules'
            Check 'a new revision is picked up on the next run, no reopening' (($r -match '^note=Rules: effort Effort B rev 5, then') -and ($r -match 'caption=Effort Caption 5 \|')) $r
            $c = RunW $word 'SKW_Test.TW_Cache'
            Check 'local copy: the old revision folder is removed' ($c -eq 'folders: r5 | root=r5') $c
            [void](RunW $word 'SKW_Test.TW_RenameFolder' "$effB|$effB moved")
            $r = RunW $word 'SKW_Test.TW_Rules'
            $lf = TR $word 'T_SkLinkFiles'
            [void](RunW $word 'SKW_Test.TW_RenameFolder' "$effB moved|$effB")
            Check 'out of reach (offline, sync off): the last complete copy is used' (($r -match '^note=Rules: effort Effort B rev 5 \(last complete copy: folder missing\)') -and ($r -match 'caption=Effort Caption 5 \|') -and ($lf -like 'Folder missing | cached=True*')) "$r ;; $lf"
            [void](TR1 $word 'T_SkLink' '')
            $c = RunW $word 'SKW_Test.TW_Cache'
            $r = RunW $word 'SKW_Test.TW_Rules'
            Check 'Unlink removes the local copy; the tools go back to the workbook' (($c -eq 'no cache') -and ($r -match '^note=Standards: Team standards\.xlsx \|') -and ($r -match 'caption=Caption \|')) "$c ;; $r"
        }
        finally { [void](TR1 $word 'T_SkLink' '') }
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
    Get-CimInstance Win32_Process -Filter "Name='POWERPNT.EXE'" -ErrorAction SilentlyContinue |
        Where-Object { ($pptBefore -notcontains $_.ProcessId) -and ($_.CommandLine -match '/automation|-Embedding') } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue |
        Where-Object { ($excelBefore -notcontains $_.ProcessId) -and ($_.CommandLine -match '/automation|-Embedding') } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-ItemProperty 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Tools\Standards' -Name 'Team Sample Template.dotm' -ErrorAction SilentlyContinue
    Remove-ItemProperty 'HKCU:\Software\VB and VBA Program Settings\Stationkeeper Tools\Run' -Name 'quiet' -ErrorAction SilentlyContinue
    if ($null -ne $prevLink) { if (-not (Test-Path $linkKey)) { [void](New-Item $linkKey -Force) }; Set-ItemProperty $linkKey -Name folder -Value $prevLink }
    if ($null -ne $prevBackup) { Set-ItemProperty $setKey -Name backupDir -Value $prevBackup }
    else { Remove-ItemProperty $setKey -Name backupDir -ErrorAction SilentlyContinue }
    Write-Output ''
    Write-Output ("RESULT: {0} checks, {1} failed." -f $results.Count, $fails)
    $results | Set-Content (Join-Path $work 'test-results.txt')
}
