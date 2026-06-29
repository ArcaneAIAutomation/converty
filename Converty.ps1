# ============================================================================
#  Converty  --  Microsoft Publisher (.pub) converter
#
#  Convert .pub files to PDF and/or PowerPoint, scan a folder tree and report
#  exactly where the .pub files live, and (electively) remove the originals.
#
#  Run via Converty.cmd  (launches Windows PowerShell 5.1 in STA mode).
#  Publisher / PowerPoint automation is delegated to cscript VBScript engines
#  (engine_publisher.vbs / engine_powerpoint.vbs) because Publisher's COM
#  server deadlocks when driven directly from PowerShell.
# ============================================================================

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- paths -----------------------------------------------------------------
$Script:AppDir       = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:EnginePub    = Join-Path $AppDir 'engine_publisher.vbs'
$Script:EnginePpt    = Join-Path $AppDir 'engine_powerpoint.vbs'
$Script:WorkRoot     = Join-Path $env:TEMP ('Converty_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $WorkRoot | Out-Null

# ---- shared state ----------------------------------------------------------
$Script:Root        = $null
$Script:Files       = @()          # list of [pscustomobject] file records
$Script:ItemByPath  = @{}          # pub path -> ListViewItem
$Script:Converted   = @{}          # pub path -> $true when at least one output succeeded
$Script:Busy        = $false

# ============================================================================
#  WinRT PDF -> PNG rasterizer (built into Windows; used for PowerPoint output)
# ============================================================================
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]              | Out-Null
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]                | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$Script:WrtExt = [System.WindowsRuntimeSystemExtensions].GetMethods()

function Await-Op($op, $resultType) {
    $m = ($Script:WrtExt | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0].MakeGenericMethod($resultType)
    $t = $m.Invoke($null, @($op)); $t.Wait(-1) | Out-Null; $t.Result
}
function Await-Act($act) {
    $m = ($Script:WrtExt | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and -not $_.IsGenericMethod -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
    $t = $m.Invoke($null, @($act)); $t.Wait(-1) | Out-Null
}

# Renders every page of $pdfPath to page-1.png .. page-N.png inside $outDir at
# ~$dpi.  Returns @{ Count; WidthPts; HeightPts } (page size of the first page).
function Convert-PdfToPngs {
    param([string]$pdfPath, [string]$outDir, [int]$dpi = 150)
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $file = Await-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)) ([Windows.Storage.StorageFile])
    $pdf  = Await-Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
    $n = [int]$pdf.PageCount
    $wPts = 0.0; $hPts = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $page = $pdf.GetPage([uint32]$i)
        if ($i -eq 0) { $wPts = [double]$page.Size.Width * 0.75; $hPts = [double]$page.Size.Height * 0.75 }
        $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $opts.DestinationWidth = [uint32][math]::Round([double]$page.Size.Width * $dpi / 96.0)
        $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        Await-Act ($page.RenderToStreamAsync($stream, $opts))
        $size   = [uint32]$stream.Size
        $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
        Await-Op ($reader.LoadAsync($size)) ([uint32]) | Out-Null
        $bytes  = New-Object byte[] $size
        $reader.ReadBytes($bytes)
        [System.IO.File]::WriteAllBytes((Join-Path $outDir ("page-{0}.png" -f ($i + 1))), $bytes)
        $reader.Dispose(); $stream.Dispose(); $page.Dispose()
    }
    [pscustomobject]@{ Count = $n; WidthPts = $wPts; HeightPts = $hPts }
}

# ============================================================================
#  Default-printer safety (optional). A slow/offline default *network* printer
#  can make Office's PDF export hang on printer-metrics init. When enabled we
#  temporarily switch the default to "Microsoft Print to PDF" and restore it.
# ============================================================================
$Script:PrinterRegKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows'

function Push-LocalPdfPrinter {
    $dev = (Get-ItemProperty $Script:PrinterRegKey -Name Device -ErrorAction SilentlyContinue).Device
    $havePdf = Get-CimInstance Win32_Printer -Filter "Name='Microsoft Print to PDF'" -ErrorAction SilentlyContinue
    if (-not $dev -or -not $havePdf) { return $null }
    $origMode = (Get-ItemProperty $Script:PrinterRegKey -Name LegacyDefaultPrinterMode -ErrorAction SilentlyContinue).LegacyDefaultPrinterMode
    $origName = ($dev -split ',')[0]
    try {
        Set-ItemProperty $Script:PrinterRegKey -Name LegacyDefaultPrinterMode -Value 1 -Type DWord
        (New-Object -ComObject WScript.Network).SetDefaultPrinter('Microsoft Print to PDF')
    } catch { return $null }
    [pscustomobject]@{ Name = $origName; Mode = $origMode }
}
function Pop-LocalPdfPrinter($state) {
    if (-not $state) { return }
    try { (New-Object -ComObject WScript.Network).SetDefaultPrinter($state.Name) } catch {}
    try {
        if ($null -ne $state.Mode) { Set-ItemProperty $Script:PrinterRegKey -Name LegacyDefaultPrinterMode -Value $state.Mode -Type DWord }
        else { Remove-ItemProperty $Script:PrinterRegKey -Name LegacyDefaultPrinterMode -ErrorAction SilentlyContinue }
    } catch {}
}

# ============================================================================
#  cscript batch runner with stall watchdog
#  $Jobs : array of objects; $ToLine builds a manifest line per job;
#  $KeyOf returns the job's key (the .pub/.pptx path) for result tracking.
#  Returns a hashtable key -> @{ Ok; Info; Error }.  $OnResult fires per file.
# ============================================================================
function Invoke-EngineBatch {
    param(
        [string]$EnginePath,
        [object[]]$Jobs,
        [scriptblock]$ToLine,
        [scriptblock]$KeyOf,
        [scriptblock]$OnResult,
        [scriptblock]$OnProgress,
        [int]$PerFileTimeoutSec = 240
    )
    $results = @{}
    $pending = New-Object System.Collections.Generic.List[object]
    foreach ($j in $Jobs) { $pending.Add($j) }

    while ($pending.Count -gt 0) {
        $manifest = Join-Path $WorkRoot ('manifest_' + [guid]::NewGuid().ToString('N') + '.txt')
        $logFile  = Join-Path $WorkRoot ('log_'      + [guid]::NewGuid().ToString('N') + '.txt')
        $lines = foreach ($j in $pending) { & $ToLine $j }
        [System.IO.File]::WriteAllLines($manifest, [string[]]$lines, [System.Text.UnicodeEncoding]::new($false, $true))
        [System.IO.File]::WriteAllText($logFile, '', [System.Text.UnicodeEncoding]::new($false, $true))

        $proc = Start-Process -FilePath 'cscript.exe' `
            -ArgumentList @('//B', '//Nologo', $EnginePath, $manifest, $logFile) `
            -WindowStyle Hidden -PassThru

        $seen = 0; $doneFiles = 0; $fatal = $null; $lastProgress = Get-Date; $stalled = $false
        while (-not $proc.HasExited) {
            Start-Sleep -Milliseconds 350
            [System.Windows.Forms.Application]::DoEvents()
            $raw = $null
            try { $raw = @(Get-Content -LiteralPath $logFile -Encoding Unicode -ErrorAction Stop) } catch { $raw = $null }
            if ($raw) {
                for ($k = $seen; $k -lt $raw.Count; $k++) {
                    $ln = $raw[$k]
                    if ($ln.StartsWith('OK|'))        { $p = $ln.Split('|'); $key=$p[1]; $results[$key]=@{Ok=$true;Info=$p[$p.Count-1];Error=$null}; if($OnResult){& $OnResult $key $true $p[$p.Count-1]} }
                    elseif ($ln.StartsWith('ERR|'))   { $p = $ln.Split('|'); $key=$p[1]; $results[$key]=@{Ok=$false;Info=$null;Error=$p[$p.Count-1]}; if($OnResult){& $OnResult $key $false $p[$p.Count-1]} }
                    elseif ($ln -eq 'DONE_FILE')      { $doneFiles++; $lastProgress = Get-Date; if($OnProgress){& $OnProgress} }
                    elseif ($ln.StartsWith('FATAL|')) { $fatal = $ln.Substring(6) }
                }
                $seen = $raw.Count
            }
            if ($fatal) { break }
            if (((Get-Date) - $lastProgress).TotalSeconds -gt $PerFileTimeoutSec) { $stalled = $true; break }
        }

        if ($stalled) {
            # The job being processed (index = $doneFiles) hung. Kill the engine +
            # Office, record the hang, and resume on whatever is left.
            try { Get-Process MSPUB, POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
            try { $proc | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
            $idx = [math]::Min($doneFiles, $pending.Count - 1)
            $stalledKey = & $KeyOf $pending[$idx]
            if (-not $results.ContainsKey($stalledKey)) {
                $results[$stalledKey] = @{ Ok=$false; Info=$null; Error='Timed out (document hung) - skipped' }
                if ($OnResult) { & $OnResult $stalledKey $false 'Timed out (document hung) - skipped' }
            }
        }
        else {
            try {
                $raw = @(Get-Content -LiteralPath $logFile -Encoding Unicode -ErrorAction SilentlyContinue)
                for ($k = $seen; $k -lt $raw.Count; $k++) {
                    $ln = $raw[$k]
                    if ($ln.StartsWith('OK|'))      { $p=$ln.Split('|'); $results[$p[1]]=@{Ok=$true;Info=$p[$p.Count-1];Error=$null}; if($OnResult){& $OnResult $p[1] $true $p[$p.Count-1]} }
                    elseif ($ln.StartsWith('ERR|')) { $p=$ln.Split('|'); $results[$p[1]]=@{Ok=$false;Info=$null;Error=$p[$p.Count-1]}; if($OnResult){& $OnResult $p[1] $false $p[$p.Count-1]} }
                    elseif ($ln.StartsWith('FATAL|')){ $fatal = $ln.Substring(6) }
                }
            } catch {}
        }

        if ($fatal) {
            foreach ($j in $pending) {
                $key = & $KeyOf $j
                if (-not $results.ContainsKey($key)) {
                    $results[$key] = @{ Ok=$false; Info=$null; Error=$fatal }
                    if ($OnResult) { & $OnResult $key $false $fatal }
                }
            }
            break
        }

        $still = New-Object System.Collections.Generic.List[object]
        foreach ($j in $pending) { if (-not $results.ContainsKey((& $KeyOf $j))) { $still.Add($j) } }
        if ($still.Count -eq $pending.Count) {
            $key = & $KeyOf $pending[0]
            $results[$key] = @{ Ok=$false; Info=$null; Error='No progress from engine - skipped' }
            if ($OnResult) { & $OnResult $key $false 'No progress from engine - skipped' }
            $still.RemoveAt(0)
        }
        $pending = $still
    }
    return $results
}

# ============================================================================
#  Scan + report
# ============================================================================
function Get-PubFiles {
    param([string]$root, [bool]$recurse)
    $opt = @{ Path = $root; Filter = '*.pub'; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($recurse) { $opt.Recurse = $true }
    Get-ChildItem @opt | Where-Object { $_.Extension -ieq '.pub' } |
        Sort-Object FullName |
        ForEach-Object {
            [pscustomobject]@{
                FullName  = $_.FullName
                Directory = $_.DirectoryName
                Name      = $_.Name
                Length    = $_.Length
                LastWrite = $_.LastWriteTime
                RelDir    = ($_.DirectoryName.Substring([math]::Min($root.Length, $_.DirectoryName.Length))).TrimStart('\')
                Status    = 'Found'
            }
        }
}

function Format-Size([long]$b) {
    if ($b -ge 1GB) { '{0:N2} GB' -f ($b/1GB) }
    elseif ($b -ge 1MB) { '{0:N2} MB' -f ($b/1MB) }
    elseif ($b -ge 1KB) { '{0:N1} KB' -f ($b/1KB) }
    else { "$b B" }
}
function HtmlEnc([string]$s) { if ($null -eq $s) { '' } else { $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;') } }

function Build-Report {
    param([string]$root, [object[]]$files)
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlOut = Join-Path $root "Converty_Report_$stamp.html"
    $csvOut  = Join-Path $root "Converty_Report_$stamp.csv"
    $txtOut  = Join-Path $root "Converty_Report_$stamp.txt"
    $totalBytes = ($files | Measure-Object Length -Sum).Sum
    if (-not $totalBytes) { $totalBytes = 0 }
    $folders = $files | Group-Object Directory | Sort-Object Name

    $files | Select-Object @{n='FileName';e={$_.Name}}, @{n='Folder';e={$_.Directory}},
        @{n='RelativeFolder';e={$_.RelDir}}, @{n='SizeBytes';e={$_.Length}},
        @{n='LastModified';e={$_.LastWrite}}, @{n='FullPath';e={$_.FullName}} |
        Export-Csv -LiteralPath $csvOut -NoTypeInformation -Encoding UTF8

    $tb = New-Object System.Text.StringBuilder
    [void]$tb.AppendLine("Converty - Publisher File Report")
    [void]$tb.AppendLine("Generated : $(Get-Date)")
    [void]$tb.AppendLine("Root      : $root")
    [void]$tb.AppendLine("Files     : $($files.Count)")
    [void]$tb.AppendLine("Total size: $(Format-Size $totalBytes)")
    [void]$tb.AppendLine("Folders   : $($folders.Count)")
    [void]$tb.AppendLine(("=" * 70))
    foreach ($g in $folders) {
        $rel = ($g.Name.Substring([math]::Min($root.Length,$g.Name.Length))).TrimStart('\')
        if ([string]::IsNullOrEmpty($rel)) { $rel = '.\ (root)' }
        [void]$tb.AppendLine("")
        [void]$tb.AppendLine("[$rel]   ($($g.Count) file(s))")
        foreach ($f in ($g.Group | Sort-Object Name)) {
            [void]$tb.AppendLine(("    - {0,-45} {1,12}   {2}" -f $f.Name, (Format-Size $f.Length), $f.LastWrite))
        }
    }
    [System.IO.File]::WriteAllText($txtOut, $tb.ToString(), [System.Text.Encoding]::UTF8)

    $h = New-Object System.Text.StringBuilder
    [void]$h.Append(@"
<!DOCTYPE html><html><head><meta charset="utf-8"><title>Converty Report</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2937;background:#f8fafc}
 h1{margin:0 0 4px}.sub{color:#64748b;margin-bottom:18px}
 .cards{display:flex;gap:14px;margin-bottom:22px;flex-wrap:wrap}
 .card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 18px;min-width:120px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
 .card .n{font-size:26px;font-weight:700;color:#0f766e}.card .l{color:#64748b;font-size:12px;text-transform:uppercase;letter-spacing:.04em}
 h2{margin-top:28px;border-bottom:2px solid #e2e8f0;padding-bottom:6px}
 .folder{margin:14px 0 6px;font-weight:600;color:#0f172a}
 .folder .cnt{color:#64748b;font-weight:400;font-size:12px;margin-left:8px}
 table{border-collapse:collapse;width:100%;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
 th,td{text-align:left;padding:8px 10px;font-size:13px;border-bottom:1px solid #f1f5f9}
 th{background:#f1f5f9;color:#334155}
 tr:hover td{background:#f8fafc}
 .ml{margin-left:22px}.path{color:#94a3b8;font-size:12px}
</style></head><body>
<h1>Converty &mdash; Publisher File Report</h1>
<div class="sub">Generated $(Get-Date) &nbsp;&bull;&nbsp; Root: <b>$(HtmlEnc $root)</b></div>
<div class="cards">
 <div class="card"><div class="n">$($files.Count)</div><div class="l">Publisher files</div></div>
 <div class="card"><div class="n">$($folders.Count)</div><div class="l">Folders</div></div>
 <div class="card"><div class="n">$(Format-Size $totalBytes)</div><div class="l">Total size</div></div>
</div>
<h2>Folder structure</h2>
"@)
    foreach ($g in $folders) {
        $rel = ($g.Name.Substring([math]::Min($root.Length,$g.Name.Length))).TrimStart('\')
        if ([string]::IsNullOrEmpty($rel)) { $rel = '(root folder)' }
        [void]$h.Append("<div class='folder'>&#128193; $(HtmlEnc $rel)<span class='cnt'>$($g.Count) file(s)</span></div>")
        [void]$h.Append("<table class='ml'><tr><th>File</th><th>Size</th><th>Last modified</th></tr>")
        foreach ($f in ($g.Group | Sort-Object Name)) {
            [void]$h.Append("<tr><td>$(HtmlEnc $f.Name)</td><td>$(Format-Size $f.Length)</td><td>$($f.LastWrite)</td></tr>")
        }
        [void]$h.Append("</table>")
    }
    [void]$h.Append("<h2>All files</h2><table><tr><th>#</th><th>File</th><th>Folder</th><th>Size</th><th>Modified</th></tr>")
    $i = 0
    foreach ($f in ($files | Sort-Object FullName)) {
        $i++
        [void]$h.Append("<tr><td>$i</td><td>$(HtmlEnc $f.Name)<div class='path'>$(HtmlEnc $f.FullName)</div></td><td>$(HtmlEnc $f.RelDir)</td><td>$(Format-Size $f.Length)</td><td>$($f.LastWrite)</td></tr>")
    }
    [void]$h.Append("</table></body></html>")
    [System.IO.File]::WriteAllText($htmlOut, $h.ToString(), [System.Text.Encoding]::UTF8)

    [pscustomobject]@{ Html = $htmlOut; Csv = $csvOut; Txt = $txtOut }
}

# ============================================================================
#  GUI
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Converty  -  Publisher Converter'
$form.Size = New-Object System.Drawing.Size(860, 724)
$form.MinimumSize = New-Object System.Drawing.Size(720, 624)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)

$teal = [System.Drawing.Color]::FromArgb(15, 118, 110)
$red  = [System.Drawing.Color]::FromArgb(190, 40, 40)

$hdr = New-Object System.Windows.Forms.Label
$hdr.Text = 'Converty'
$hdr.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16, [System.Drawing.FontStyle]::Bold)
$hdr.ForeColor = $teal
$hdr.Location = New-Object System.Drawing.Point(18, 12)
$hdr.AutoSize = $true
$form.Controls.Add($hdr)

$hdr2 = New-Object System.Windows.Forms.Label
$hdr2.Text = 'Convert Microsoft Publisher (.pub) files to PDF or PowerPoint'
$hdr2.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$hdr2.Location = New-Object System.Drawing.Point(20, 44)
$hdr2.AutoSize = $true
$form.Controls.Add($hdr2)

$lblF = New-Object System.Windows.Forms.Label
$lblF.Text = 'Folder:'; $lblF.Location = New-Object System.Drawing.Point(20, 78); $lblF.AutoSize = $true
$form.Controls.Add($lblF)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(70, 75)
$txtFolder.Size = New-Object System.Drawing.Size(610, 24)
$txtFolder.ReadOnly = $true
$txtFolder.Anchor = 'Top,Left,Right'
$txtFolder.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'; $btnBrowse.Location = New-Object System.Drawing.Point(690, 74)
$btnBrowse.Size = New-Object System.Drawing.Size(140, 26); $btnBrowse.Anchor = 'Top,Right'
$form.Controls.Add($btnBrowse)

$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text = 'Include subfolders'; $chkRecurse.Checked = $true
$chkRecurse.Location = New-Object System.Drawing.Point(70, 104); $chkRecurse.AutoSize = $true
$form.Controls.Add($chkRecurse)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = 'Output format'; $grp.Location = New-Object System.Drawing.Point(20, 132)
$grp.Size = New-Object System.Drawing.Size(360, 64)
$form.Controls.Add($grp)
$rbPdf  = New-Object System.Windows.Forms.RadioButton
$rbPdf.Text = 'PDF'; $rbPdf.Checked = $true; $rbPdf.Location = New-Object System.Drawing.Point(16, 26); $rbPdf.AutoSize = $true
$rbPpt  = New-Object System.Windows.Forms.RadioButton
$rbPpt.Text = 'PowerPoint'; $rbPpt.Location = New-Object System.Drawing.Point(100, 26); $rbPpt.AutoSize = $true
$rbBoth = New-Object System.Windows.Forms.RadioButton
$rbBoth.Text = 'Both'; $rbBoth.Location = New-Object System.Drawing.Point(220, 26); $rbBoth.AutoSize = $true
$grp.Controls.AddRange(@($rbPdf, $rbPpt, $rbBoth))

$chkOverwrite = New-Object System.Windows.Forms.CheckBox
$chkOverwrite.Text = 'Overwrite existing output files'; $chkOverwrite.Checked = $true
$chkOverwrite.Location = New-Object System.Drawing.Point(400, 140); $chkOverwrite.AutoSize = $true
$form.Controls.Add($chkOverwrite)

$chkForcePdf = New-Object System.Windows.Forms.CheckBox
$chkForcePdf.Text = 'Force local PDF printer (use if conversions stall)'
$chkForcePdf.Location = New-Object System.Drawing.Point(400, 166); $chkForcePdf.AutoSize = $true
$form.Controls.Add($chkForcePdf)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'Scan && Build Report'; $btnScan.Location = New-Object System.Drawing.Point(20, 204)
$btnScan.Size = New-Object System.Drawing.Size(170, 32)
$btnScan.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($btnScan)

$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Text = 'Convert'; $btnConvert.Location = New-Object System.Drawing.Point(200, 204)
$btnConvert.Size = New-Object System.Drawing.Size(150, 32)
$btnConvert.BackColor = $teal; $btnConvert.ForeColor = [System.Drawing.Color]::White
$btnConvert.FlatStyle = 'Flat'; $btnConvert.Enabled = $false
$form.Controls.Add($btnConvert)

$lv = New-Object System.Windows.Forms.ListView
$lv.Location = New-Object System.Drawing.Point(20, 248)
$lv.Size = New-Object System.Drawing.Size(810, 270)
$lv.Anchor = 'Top,Bottom,Left,Right'
$lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true
$lv.BackColor = [System.Drawing.Color]::White
[void]$lv.Columns.Add('File', 230)
[void]$lv.Columns.Add('Folder', 300)
[void]$lv.Columns.Add('Size', 90)
[void]$lv.Columns.Add('Status', 170)
$form.Controls.Add($lv)

$pb = New-Object System.Windows.Forms.ProgressBar
$pb.Location = New-Object System.Drawing.Point(20, 526); $pb.Size = New-Object System.Drawing.Size(810, 16)
$pb.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($pb)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Select a folder to begin.'; $lblStatus.Location = New-Object System.Drawing.Point(20, 548)
$lblStatus.Size = New-Object System.Drawing.Size(810, 20); $lblStatus.Anchor = 'Bottom,Left,Right'
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
$form.Controls.Add($lblStatus)

$chkOnlyConverted = New-Object System.Windows.Forms.CheckBox
$chkOnlyConverted.Text = 'Only delete files that converted successfully'; $chkOnlyConverted.Checked = $true
$chkOnlyConverted.Location = New-Object System.Drawing.Point(20, 576); $chkOnlyConverted.AutoSize = $true
$chkOnlyConverted.Anchor = 'Bottom,Left'
$form.Controls.Add($chkOnlyConverted)

$chkPermanent = New-Object System.Windows.Forms.CheckBox
$chkPermanent.Text = 'Permanently delete (skip Recycle Bin)'
$chkPermanent.Location = New-Object System.Drawing.Point(20, 598); $chkPermanent.AutoSize = $true
$chkPermanent.Anchor = 'Bottom,Left'
$form.Controls.Add($chkPermanent)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = 'Delete Original .pub Files...'
$btnDelete.Location = New-Object System.Drawing.Point(560, 584)
$btnDelete.Size = New-Object System.Drawing.Size(270, 34); $btnDelete.Anchor = 'Bottom,Right'
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(214, 216, 220); $btnDelete.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
$btnDelete.FlatStyle = 'Flat'; $btnDelete.Enabled = $false
$form.Controls.Add($btnDelete)

# --- footer credit note (clickable GitHub + website links) ---
$footer = New-Object System.Windows.Forms.LinkLabel
$footer.AutoSize = $false
$footer.TextAlign = 'MiddleCenter'
$footer.Location = New-Object System.Drawing.Point(20, 658)
$footer.Size = New-Object System.Drawing.Size(810, 20)
$footer.Anchor = 'Bottom,Left,Right'
$footer.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$footer.ForeColor = [System.Drawing.Color]::FromArgb(120, 130, 145)
$footer.LinkColor = $teal
$footer.ActiveLinkColor = $red
$footer.Text = 'Created by Morgan Coetzee     |     github.com/arcaneAIAutomation     |     arcane.group'
$footer.LinkArea = New-Object System.Windows.Forms.LinkArea(0, 0)
$footer.Links.Clear()
$ghText  = 'github.com/arcaneAIAutomation'
$webText = 'arcane.group'
[void]$footer.Links.Add($footer.Text.IndexOf($ghText),  $ghText.Length,  'https://github.com/arcaneAIAutomation')
[void]$footer.Links.Add($footer.Text.IndexOf($webText), $webText.Length, 'https://arcane.group')
$footer.Add_LinkClicked({ param($s, $e) try { Start-Process ([string]$e.Link.LinkData) } catch {} })
$form.Controls.Add($footer)

# ---- UI helpers ----
function Set-Status([string]$s)  { $lblStatus.Text = $s; [System.Windows.Forms.Application]::DoEvents() }
function Show-Info([string]$m) {
    if ($env:CONVERTY_HEADLESS) { Write-Output "[info] $m" }
    else { [void][System.Windows.Forms.MessageBox]::Show($m, 'Converty') }
}
function Set-Busy([bool]$b) {
    $Script:Busy = $b
    $en = -not $b
    $btnBrowse.Enabled=$en; $btnScan.Enabled=$en; $chkRecurse.Enabled=$en
    $btnConvert.Enabled = ($en -and $Script:Files.Count -gt 0)
    # Deletion is only ever offered AFTER a conversion has actually produced
    # output this session - never by default, only if the user then elects to.
    $delOn = ($en -and $Script:Converted.Count -gt 0)
    $btnDelete.Enabled = $delOn
    if ($delOn) { $btnDelete.BackColor = $red; $btnDelete.ForeColor = [System.Drawing.Color]::White }
    else { $btnDelete.BackColor = [System.Drawing.Color]::FromArgb(214, 216, 220); $btnDelete.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140) }
    $form.Cursor = if ($b) { 'WaitCursor' } else { 'Default' }
}
function Update-Item([string]$pub, [string]$status, [System.Drawing.Color]$color) {
    if ($Script:ItemByPath.ContainsKey($pub)) {
        $it = $Script:ItemByPath[$pub]
        $it.SubItems[3].Text = $status
        $it.ForeColor = $color
        $lv.Refresh()
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ============================================================================
#  Core operations (called by the buttons AND by the self-test hook)
# ============================================================================
function Invoke-Scan {
    if (-not $Script:Root) { Show-Info 'Please choose a folder first.'; return }
    Set-Busy $true
    try {
        Set-Status 'Scanning for .pub files...'
        $Script:Files = @(Get-PubFiles -root $Script:Root -recurse $chkRecurse.Checked)
        $lv.Items.Clear(); $Script:ItemByPath=@{}; $Script:Converted=@{}
        foreach ($f in $Script:Files) {
            $it = New-Object System.Windows.Forms.ListViewItem($f.Name)
            [void]$it.SubItems.Add($f.RelDir); [void]$it.SubItems.Add((Format-Size $f.Length)); [void]$it.SubItems.Add('Found')
            [void]$lv.Items.Add($it); $Script:ItemByPath[$f.FullName] = $it
        }
        if ($Script:Files.Count -eq 0) {
            Set-Status 'No .pub files found in the selected folder.'
            Show-Info 'No Publisher (.pub) files were found here.'
        } else {
            Set-Status "Building report for $($Script:Files.Count) file(s)..."
            $r = Build-Report -root $Script:Root -files $Script:Files
            Set-Status "Found $($Script:Files.Count) file(s). Report saved next to your folder."
            if (-not $env:CONVERTY_HEADLESS) { try { Start-Process $r.Html } catch {} }
        }
    } catch {
        Show-Info "Scan failed:`n$($_.Exception.Message)"
    } finally { Set-Busy $false }
}

function Invoke-ConvertFiles {
    if ($Script:Files.Count -eq 0) { return }
    $wantPdf = $rbPdf.Checked -or $rbBoth.Checked
    $wantPpt = $rbPpt.Checked -or $rbBoth.Checked
    Set-Busy $true
    $printerState = $null
    try {
        if ($chkForcePdf.Checked) { $printerState = Push-LocalPdfPrinter }

        foreach ($f in $Script:Files) { Update-Item $f.FullName 'Queued' ([System.Drawing.Color]::FromArgb(100,116,139)) }
        $pb.Value = 0; $pb.Maximum = [math]::Max(1, $Script:Files.Count)

        # ---- Phase 1: Publisher -> PDF (final pdf if wanted, else temp) ----
        $pubJobs = @()
        $pdfFor = @{}
        foreach ($f in $Script:Files) {
            if ($wantPdf) { $pdf = [System.IO.Path]::ChangeExtension($f.FullName, '.pdf') }
            else { $pdf = Join-Path $WorkRoot (([guid]::NewGuid().ToString('N')) + '.pdf') }
            $pdfFor[$f.FullName] = $pdf
            if ($wantPdf -and -not $chkOverwrite.Checked -and (Test-Path $pdf)) {
                # Output already present - count it as converted (so it can later be
                # deleted) only if every requested format already exists.
                $pptxExisting = [System.IO.Path]::ChangeExtension($f.FullName, '.pptx')
                if (-not $wantPpt -or (Test-Path $pptxExisting)) { $Script:Converted[$f.FullName] = $true }
                Update-Item $f.FullName 'Skipped (output exists)' ([System.Drawing.Color]::FromArgb(120,120,120))
                continue
            }
            $pubJobs += [pscustomobject]@{ Pub=$f.FullName; Pdf=$pdf }
        }

        Set-Status "Converting $($pubJobs.Count) file(s) with Publisher..."
        $pubResults = @{}
        if ($pubJobs.Count -gt 0) {
            $pubResults = Invoke-EngineBatch -EnginePath $EnginePub -Jobs $pubJobs `
                -ToLine { param($j) "$($j.Pub)|$($j.Pdf)" } `
                -KeyOf  { param($j) $j.Pub } `
                -OnProgress { $pb.Value = [math]::Min($pb.Maximum, $pb.Value + 1) } `
                -OnResult {
                    param($key, $ok, $info)
                    if ($ok) {
                        if ($wantPdf) { Update-Item $key 'PDF done' ([System.Drawing.Color]::FromArgb(15,118,110)); $Script:Converted[$key]=$true }
                        else { Update-Item $key 'Rendered' ([System.Drawing.Color]::FromArgb(15,118,110)) }
                    } else { Update-Item $key "Failed: $info" $red }
                }
        }

        # ---- Phase 2 + 3: PowerPoint (rasterise each PDF, then assemble deck) ----
        if ($wantPpt) {
            $pptJobs = @()
            foreach ($f in $Script:Files) {
                $r = $pubResults[$f.FullName]
                if (-not $r -or -not $r.Ok) { continue }
                $pptx = [System.IO.Path]::ChangeExtension($f.FullName, '.pptx')
                if (-not $chkOverwrite.Checked -and (Test-Path $pptx)) { $Script:Converted[$f.FullName] = $true; Update-Item $f.FullName 'Skipped (PPTX exists)' ([System.Drawing.Color]::FromArgb(120,120,120)); continue }
                Update-Item $f.FullName 'Rendering slides...' ([System.Drawing.Color]::FromArgb(100,116,139))
                try {
                    $pngDir = Join-Path $WorkRoot ([guid]::NewGuid().ToString('N'))
                    $info = Convert-PdfToPngs -pdfPath $pdfFor[$f.FullName] -outDir $pngDir
                    $pptJobs += [pscustomobject]@{ Pub=$f.FullName; Pptx=$pptx; Dir=$pngDir; W=$info.WidthPts; H=$info.HeightPts; N=$info.Count }
                } catch {
                    Update-Item $f.FullName "Slide render failed" $red
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            if ($pptJobs.Count -gt 0) {
                Set-Status "Building $($pptJobs.Count) PowerPoint deck(s)..."
                $pptKeyByPptx = @{}; foreach ($j in $pptJobs) { $pptKeyByPptx[$j.Pptx] = $j.Pub }
                Invoke-EngineBatch -EnginePath $EnginePpt -Jobs $pptJobs `
                    -ToLine { param($j) "$($j.Pptx)|$($j.Dir)|$([math]::Round($j.W,2))|$([math]::Round($j.H,2))|$($j.N)" } `
                    -KeyOf  { param($j) $j.Pptx } `
                    -OnResult {
                        param($key, $ok, $info)
                        $pub = $pptKeyByPptx[$key]
                        if ($ok) {
                            if ($rbBoth.Checked) { Update-Item $pub 'PDF + PPTX done' ([System.Drawing.Color]::FromArgb(15,118,110)) }
                            else { Update-Item $pub 'PPTX done' ([System.Drawing.Color]::FromArgb(15,118,110)) }
                            $Script:Converted[$pub] = $true
                        } else { Update-Item $pub "PPTX failed: $info" $red }
                    } | Out-Null
            }
        }

        $okCount = ($Script:Converted.Keys).Count
        $pb.Value = $pb.Maximum
        Set-Status "Done. $okCount of $($Script:Files.Count) file(s) converted. Outputs are next to each .pub file."
    } catch {
        Show-Info "Conversion error:`n$($_.Exception.Message)"
    } finally {
        Pop-LocalPdfPrinter $printerState
        try { Get-ChildItem $WorkRoot -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        Set-Busy $false
    }
}

function Invoke-DeleteOriginals {
    if ($Script:Files.Count -eq 0) { return }
    $targets = if ($chkOnlyConverted.Checked) {
        @($Script:Files | Where-Object { $Script:Converted.ContainsKey($_.FullName) })
    } else { @($Script:Files) }

    if ($targets.Count -eq 0) {
        Show-Info 'No files match the deletion criteria. (Tip: convert first, or untick "Only delete files that converted successfully".)'
        return
    }
    $how = if ($chkPermanent.Checked) { 'PERMANENTLY DELETED' } else { 'moved to the Recycle Bin' }
    if (-not $env:CONVERTY_HEADLESS) {
        $msg = "You are about to delete $($targets.Count) original Publisher (.pub) file(s).`n`nThey will be $how.`n`nThis cannot be easily undone. Continue?"
        if ([System.Windows.Forms.MessageBox]::Show($msg,'Confirm delete',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning) -ne 'Yes') { return }
    }

    Set-Busy $true
    $deleted = 0; $failed = 0
    $recycle = [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    foreach ($f in $targets) {
        try {
            if ($chkPermanent.Checked) {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName, 'OnlyErrorDialogs', [Microsoft.VisualBasic.FileIO.RecycleOption]::DeletePermanently)
            } else {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName, 'OnlyErrorDialogs', $recycle)
            }
            Update-Item $f.FullName 'Deleted' ([System.Drawing.Color]::FromArgb(120,120,120))
            $deleted++
        } catch { Update-Item $f.FullName 'Delete failed' $red; $failed++ }
        [System.Windows.Forms.Application]::DoEvents()
    }
    Set-Busy $false
    Set-Status "Deleted $deleted file(s)$(if($failed){"; $failed failed"})."
    Show-Info "Deleted $deleted file(s).$(if($failed){"`n$failed could not be deleted."})"
}

# ============================================================================
#  Wire up buttons
# ============================================================================
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Choose the folder that contains your Publisher files'
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq 'OK') {
        $Script:Root = $dlg.SelectedPath
        $txtFolder.Text = $dlg.SelectedPath
        $lv.Items.Clear(); $Script:Files=@(); $Script:ItemByPath=@{}; $Script:Converted=@{}
        $btnConvert.Enabled = $false; $btnDelete.Enabled = $false
        Set-Status 'Folder selected. Click "Scan && Build Report".'
    }
})
$btnScan.Add_Click({ Invoke-Scan })
$btnConvert.Add_Click({ Invoke-ConvertFiles })
$btnDelete.Add_Click({ Invoke-DeleteOriginals })

$form.Add_FormClosed({
    try { Get-ChildItem $WorkRoot -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    try { Remove-Item $WorkRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
})

# --- env-gated self-test hook (inert in normal use; used for automated testing) ---
if ($env:CONVERTY_HEADLESS -and $env:CONVERTY_FOLDER) {
    $Script:Root = $env:CONVERTY_FOLDER
    $txtFolder.Text = $Script:Root
    $chkRecurse.Checked = $true
    switch ($env:CONVERTY_FORMAT) {
        'both' { $rbPdf.Checked=$false; $rbPpt.Checked=$false; $rbBoth.Checked=$true }
        'pptx' { $rbPdf.Checked=$false; $rbBoth.Checked=$false; $rbPpt.Checked=$true }
        default { $rbPdf.Checked=$true }
    }
    Invoke-Scan
    if ($env:CONVERTY_CONVERT -eq '1') { Invoke-ConvertFiles }
    if ($env:CONVERTY_DELETE  -eq '1') { Invoke-DeleteOriginals }
    Write-Output ("HEADLESS files={0} converted={1}" -f $Script:Files.Count, $Script:Converted.Keys.Count)
    foreach ($it in $lv.Items) { Write-Output ("  {0,-26} {1}" -f $it.Text, $it.SubItems[3].Text) }
    return
}

[void]$form.ShowDialog()
