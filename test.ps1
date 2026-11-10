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
foreach ($k in @('Team Sample Template.dotm')) { Remove-Item "HKCU:\Software\VB and VBA Program Settings\Stationkeeper Macros\$k" -Recurse -ErrorAction SilentlyContinue }

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
        Check 'Save disabled with nothing to save' ($adm -match 'gate0: pending=0 unticked=False ticked=False') ''
        Check 'Save needs Ready to save ticked' ($adm -match 'gate1: pending=\d+ unticked=False ticked=True') ''
        $csv = TR $word 'T_Csv'
        Check 'CSV export and import round trip (quotes, commas, line breaks)' (($csv -like 'Updated 8 macro(s)*') -and ($csv -match 'line one\\nline two, with comma')) $csv
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
        Check 'no access: Check tab explains the two ways forward' ($na -match 'Trust access to the VBA project object model') ''
        $d.Close([ref]0)
    }

}
catch {
    Check 'test run' $false ($_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}
finally {
    Write-Output ''
    Write-Output ("RESULT: {0} checks, {1} failed." -f $results.Count, $fails)
    $results | Set-Content (Join-Path $work 'test-results.txt')
}
