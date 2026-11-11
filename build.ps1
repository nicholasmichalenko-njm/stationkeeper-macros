# Builds "Stationkeeper Macros Kit.dotm" and the import files from src\.
# Usage: powershell -ExecutionPolicy Bypass -File build.ps1 [-Root <repo folder>]
# Paths go to Word as [string]: a cmdlet's output reaches COM as a wrapped object and Word stalls on it.
# Needs Word, and "Trust access to the VBA project object model" on for the build only.
param([string]$Root = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
. (Join-Path $Root 'tools\wordjob.ps1')

[string]$build = Join-Path $Root 'build'
[string]$imp = Join-Path $build 'Import files'
New-Item -ItemType Directory -Force $build, $imp | Out-Null

Invoke-WordJob -Root $Root -Job {
    param($word)
    Write-Host 'step: new document'
    $doc = $word.Documents.Add()
    $vbp = $doc.VBProject
    $vbp.Name = 'StationkeeperMacrosKit'
    [void]$vbp.VBComponents.Import((Join-Path $Root 'src\SK_Macros.bas'))
    [void]$vbp.VBComponents.Import((Join-Path $Root 'src\SK_Install.bas'))
    foreach ($t in Get-ChildItem (Join-Path $Root 'src\tools') -Filter '*.bas' | Sort-Object Name) {
        [void]$vbp.VBComponents.Import([string]$t.FullName)
    }
    Write-Host 'step: modules imported'
    foreach ($f in 'frmSKMenu', 'frmSKAdmin') {
        $c = $vbp.VBComponents.Add(3)
        $c.Name = $f
        $cm = $c.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString((Get-Content (Join-Path $Root "src\$f.code.vba") -Raw))
    }
    Write-Host 'step: forms added'
    # Compile. A compile error opens a dialog that would block forever, so a watchdog
    # reads the dialog, ends this Word after 25 seconds, and the build fails with the message.
    $wd = Start-Job -ScriptBlock { param($r) Start-Sleep 25; & (Join-Path $r 'tools\watch.ps1') -Root $r -Kill } -ArgumentList $Root
    try {
        $word.VBE.CommandBars.FindControl(1, 578).Execute()
        Stop-Job $wd; Remove-Job $wd -Force
        Write-Host 'step: compiled'
    }
    catch {
        Write-Host 'COMPILE FAILED:'
        Receive-Job $wd -Wait | ForEach-Object { Write-Host "  $_" }
        throw 'The VBA code does not compile. See the dialog text above.'
    }
    [string]$kit = Join-Path $build 'Stationkeeper Macros Kit.dotm'
    if (Test-Path $kit) { Remove-Item $kit -Force }
    Write-Host 'step: saving'
    $doc.SaveAs2($kit, 15)
    Write-Host 'step: saved'
    Get-ChildItem $imp | Remove-Item -Recurse -Force
    $vbp.VBComponents.Item('SK_Macros').Export((Join-Path $imp 'SK_Macros.bas'))
    $vbp.VBComponents.Item('frmSKMenu').Export((Join-Path $imp 'frmSKMenu.frm'))
    $vbp.VBComponents.Item('frmSKAdmin').Export((Join-Path $imp 'frmSKAdmin.frm'))
    [string]$timp = Join-Path $imp 'Stationkeeper Tools'
    New-Item -ItemType Directory -Force $timp | Out-Null
    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -like 'SK[WPX]_*') { $c.Export([string](Join-Path $timp "$($c.Name).bas")) }
    }
    $doc.Close(0)
    Write-Output "Built $kit"
    Get-ChildItem $build -Recurse -File | ForEach-Object { "  {0}  {1:N0} bytes" -f $_.FullName.Substring($Root.Length + 1), $_.Length }
}
