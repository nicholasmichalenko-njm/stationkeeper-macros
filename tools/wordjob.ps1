# Runs a script block against a Word instance this script starts, and only that one.
# Word stays hidden, alerts off, and is always quit at the end, even on error.
function Invoke-WordJob {
    param([string]$Root, [scriptblock]$Job, [object[]]$JobArgs = @())
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $word = New-Object -ComObject Word.Application
    Start-Sleep -Milliseconds 400
    $mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
    $pidFile = Join-Path $Root 'work\word.pid'
    New-Item -ItemType Directory -Force (Split-Path $pidFile) | Out-Null
    Set-Content $pidFile ($mine -join ',')
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        & $Job $word @JobArgs
    }
    finally {
        try { foreach ($d in @($word.Documents)) { $d.Close([ref]0) } } catch {}
        try { $word.Quit([ref]0) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500
        foreach ($p in $mine) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { Stop-Process -Id $p -Force } }
        Remove-Item $pidFile -ErrorAction SilentlyContinue
    }
}

function Set-VbomAccess([int]$Value) {
    $k = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
    New-Item -Path $k -Force -ErrorAction SilentlyContinue | Out-Null
    Set-ItemProperty -Path $k -Name AccessVBOM -Value $Value -Type DWord
}

function Get-VbomAccess {
    (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Word\Security' -ErrorAction SilentlyContinue).AccessVBOM
}
