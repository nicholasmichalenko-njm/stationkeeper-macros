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
    $forms = @('frmSKMenu', 'frmSKAdmin', 'frmSKBrand') | ForEach-Object { Join-Path $Root "src\$_.code.vba" }
    $forms += @(Get-ChildItem (Join-Path $Root 'src\tools') -Filter 'frm*.code.vba' | Sort-Object Name | ForEach-Object { $_.FullName })
    # The tools carry their own copy of the Stationkeeper look, so each half installs and removes alone.
    $forms += @('frmSKWBrand=' + (Join-Path $Root 'src\frmSKBrand.code.vba'))
    foreach ($ff in $forms) {
        $f = (Split-Path $ff -Leaf) -replace '\.code\.vba$', ''
        if ($ff -like '*=*') { $f = $ff.Split('=')[0]; $ff = $ff.Substring($f.Length + 1) }
        $c = $vbp.VBComponents.Add(3)
        $c.Name = $f
        $cm = $c.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString((Get-Content ([string]$ff) -Raw))
    }
    Write-Host 'step: forms added'
    # The satellite pictures go into the look's forms (kept in the .frx, so no file is needed on a PC).
    [string]$ico = Join-Path $Root 'src\brand\sk_icon.ico'
    [string]$logo = Join-Path $Root 'src\brand\sk_logo.bmp'
    $boot = $vbp.VBComponents.Add(1)
    $boot.Name = 'SKBuild'
    $boot.CodeModule.AddFromString("Public Function SKBuild_Brand(ByVal arg As String) As String`r`n    Dim d As Object, c As Object, a As Variant`r`n    On Error GoTo Fail`r`n    a = Split(arg, ""|"")`r`n    Set d = ThisDocument.VBProject.VBComponents(a(0)).Designer`r`n    Set c = d.Controls.Add(""Forms.Image.1"", ""imgLogo"")`r`n    c.Left = 0: c.Top = 0: c.Width = 37.5: c.Height = 37.5: c.PictureSizeMode = 1`r`n    Set c.Picture = LoadPicture(a(2))`r`n    Set c = d.Controls.Add(""Forms.Image.1"", ""imgIcon"")`r`n    c.Left = 44: c.Top = 0: c.Width = 24: c.Height = 24`r`n    Set c.Picture = LoadPicture(a(1))`r`n    Exit Function`r`nFail:`r`n    SKBuild_Brand = ""Error "" & Err.Number & "": "" & Err.Description`r`nEnd Function`r`n")
    foreach ($bf in @('frmSKBrand', 'frmSKWBrand')) { $v = "$bf|$ico|$logo"; $e = $word.Run('SKBuild_Brand', [ref]$v); if ($e) { throw "Could not put the brand pictures in ${bf}: $e" } }
    $vbp.VBComponents.Remove($boot)
    Write-Host 'step: brand pictures in'
    # Compile. A compile error opens a dialog that would block forever, so a watchdog
    # reads the dialog, ends this Word after 25 seconds, and the build fails with the message.
    $wd = Start-Job -ScriptBlock { param($r) Start-Sleep 25; & (Join-Path $r 'tools\watch.ps1') -Root $r -Kill } -ArgumentList $Root
    try {
        $word.VBE.ActiveVBProject = $vbp   # compile the kit, not whichever project the editor last showed
        $word.VBE.CommandBars.FindControl(1, 578).Execute()
        Stop-Job $wd; Remove-Job $wd -Force
        # With Word hidden, a compile error shows no dialog: the command just stays available.
        # Debug > Compile is greyed out only after a clean compile, so that is the check.
        if ($word.VBE.CommandBars.FindControl(1, 578).Enabled) {
            $where = ''
            try {
                $pane = $word.VBE.ActiveCodePane
                $sl = 0; $sc = 0; $el = 0; $ec = 0
                $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                $where = "$($pane.CodeModule.Parent.Name), line ${sl}: $($pane.CodeModule.Lines($sl, 1).Trim())"
            } catch { $where = '(the editor did not say where)' }
            Write-Host "COMPILE FAILED: $where"
            throw "The VBA code does not compile: $where"
        }
        Write-Host 'step: compiled'
    }
    catch {
        if ($_.Exception.Message -like 'The VBA code does not compile:*') { throw }
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
    $vbp.VBComponents.Item('frmSKBrand').Export((Join-Path $imp 'frmSKBrand.frm'))
    [string]$timp = Join-Path $imp 'Stationkeeper Tools'
    New-Item -ItemType Directory -Force $timp | Out-Null
    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -like 'SK[WPX]_*') { $c.Export([string](Join-Path $timp "$($c.Name).bas")) }
        if ($c.Name -like 'frmSK[WPX]*') { $c.Export([string](Join-Path $timp "$($c.Name).frm")) }
    }
    $doc.Close(0)
    Write-Output "Built $kit"

    # The Stationkeeper tab: a separate small add-in with a ribbon part, so it can be left out.
    Write-Host 'step: ribbon add-in'
    $rd = $word.Documents.Add()
    $rd.VBProject.Name = 'StationkeeperRibbon'
    [void]$rd.VBProject.VBComponents.Import((Join-Path $Root 'src\ribbon\SKR_Ribbon.bas'))
    $probe = $word.Run('SKR_Probe')
    if ($probe -notlike 'Stationkeeper Ribbon *') { throw "The ribbon add-in's code does not run: $probe" }
    [string]$rib = Join-Path $build 'Stationkeeper Ribbon.dotm'
    if (Test-Path $rib) { Remove-Item $rib -Force }
    $rd.SaveAs2($rib, 15)
    $rd.Close(0)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $z = [IO.Compression.ZipFile]::Open($rib, 'Update')
    try {
        function ZipText([string]$name) { $r = New-Object IO.StreamReader($z.GetEntry($name).Open()); $t = $r.ReadToEnd(); $r.Close(); $t }
        function ZipPut([string]$name, [byte[]]$bytes) { $e = $z.GetEntry($name); if ($e) { $e.Delete() }; $st = $z.CreateEntry($name).Open(); $st.Write($bytes, 0, $bytes.Length); $st.Close() }
        $ct = ZipText '[Content_Types].xml'
        if ($ct -notmatch 'Extension="png"') { $ct = ([regex]'<Default ').Replace($ct, '<Default Extension="png" ContentType="image/png"/><Default ', 1) }
        ZipPut '[Content_Types].xml' $utf8.GetBytes($ct)
        $rels = (ZipText '_rels/.rels') -replace '</Relationships>', '<Relationship Id="skCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
        ZipPut '_rels/.rels' $utf8.GetBytes($rels)
        ZipPut 'customUI/customUI14.xml' ([IO.File]::ReadAllBytes((Join-Path $Root 'src\ribbon\customUI14.xml')))
        ZipPut 'customUI/_rels/customUI14.xml.rels' $utf8.GetBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="skIcon" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/skIcon.png"/></Relationships>')
        ZipPut 'customUI/images/skIcon.png' ([IO.File]::ReadAllBytes((Join-Path $Root 'src\brand\sk_ribbon.png')))
    }
    finally { $z.Dispose() }
    Write-Output "Built $rib"
    Get-ChildItem $build -Recurse -File | ForEach-Object { "  {0}  {1:N0} bytes" -f $_.FullName.Substring($Root.Length + 1), $_.Length }
}
