# Reads any dialog the build's Word is showing (compile errors, prompts), then optionally ends that Word.
param([string]$Root = (Split-Path $PSScriptRoot), [switch]$Kill)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$pidFile = Join-Path $Root 'work\word.pid'
if (-not (Test-Path $pidFile)) { 'No build Word running.'; return }
$pids = (Get-Content $pidFile) -split ',' | Where-Object { $_ } | ForEach-Object { [int]$_ }
$A = [System.Windows.Automation.AutomationElement]
foreach ($p in $pids) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($A::ProcessIdProperty, $p)
    $wins = $A::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
    foreach ($w in $wins) {
        "WINDOW: " + $w.Current.Name + " [" + $w.Current.ClassName + "]"
        $all = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($e in $all) { if ($e.Current.Name) { "   " + $e.Current.ControlType.ProgrammaticName + ": " + $e.Current.Name } }
    }
    if ($Kill) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue; "Ended Word $p" }
}
