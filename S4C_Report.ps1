# S4C REPORT SCRIPT - makes the daily work report from the work log the tools keep on this PC.
# S4C-FILE-VERSION: 1.1
# The log is %LOCALAPPDATA%\S4C_Tools\WorkLog\worklog.tsv (one line per file a tool saved).
$ErrorActionPreference = 'Stop'
$ci = [Globalization.CultureInfo]::InvariantCulture
$base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
$logPath = Join-Path (Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog') 'worklog.tsv'

Write-Host 'Daily work report' -ForegroundColor Cyan
Write-Host ''
if ($env:S4C_REPORT_DATE) { $ans = $env:S4C_REPORT_DATE } else { $ans = Read-Host 'Press Enter for today, or type a date (for example 20-11-2026)' }
$ans = ([string]$ans).Trim()
$day = (Get-Date).Date
if ($ans) {
    $ok = $false
    foreach ($fmt in 'dd-MM-yyyy', 'd-M-yyyy', 'dd/MM/yyyy', 'd/M/yyyy', 'yyyy-MM-dd', 'dd.MM.yyyy', 'dd-MM-yy', 'dd/MM/yy') {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($ans, $fmt, $ci, [Globalization.DateTimeStyles]::None, [ref]$d)) { $day = $d.Date; $ok = $true; break }
    }
    if (-not $ok) { Write-Host ('Could not read the date "' + $ans + '". Type it like 20-11-2026.') -ForegroundColor Yellow; exit 0 }
}
$key = $day.ToString('yyyy-MM-dd', $ci)

$rows = @()
if (Test-Path -LiteralPath $logPath) {
    foreach ($l in [IO.File]::ReadAllLines($logPath, [Text.Encoding]::UTF8)) {
        $c = $l.Split("`t")
        if ($c.Count -lt 8 -or $c[0] -ne $key) { continue }
        $rows += [pscustomobject]@{ Time = $c[1]; User = $c[2]; Tool = $c[3]; In = $c[4]; Out = $c[5]; Count = $c[6]; Unit = $c[7] }
    }
}

# one line per chapter file: the same chapter run through several tools counts once
# ("CH18.pptx", "CH18 Remediated.pptx" and "CH18 Remediated AltText.pptx" are the same chapter)
$suffixRx = '(?i)(\s+(Remediated|AltText|Links|Fixed)|_BEFORE_IMBKfix\d*)+$'
$groups = [ordered]@{}
foreach ($r in $rows) {
    $leaf = [IO.Path]::GetFileName($r.In)
    $stem = [IO.Path]::GetFileNameWithoutExtension($leaf) -replace $suffixRx, ''
    $k = ($stem + [IO.Path]::GetExtension($leaf)).ToLowerInvariant()
    if (-not $groups.Contains($k)) { $groups[$k] = [pscustomobject]@{ Name = $leaf; Count = ''; Unit = $r.Unit; Tools = New-Object System.Collections.Generic.List[string]; First = $r.Time; Last = $r.Time } }
    $g = $groups[$k]
    if ($leaf.Length -lt $g.Name.Length) { $g.Name = $leaf }
    if ($r.Count -ne '') { $g.Count = $r.Count }
    if (-not $g.Tools.Contains($r.Tool)) { $g.Tools.Add($r.Tool) }
    if ([string]::CompareOrdinal($r.Time, $g.First) -lt 0) { $g.First = $r.Time }
    if ([string]::CompareOrdinal($r.Time, $g.Last) -gt 0) { $g.Last = $r.Time }
}
$items = @($groups.Values)

$user = if ($env:S4C_REPORT_NAME) { $env:S4C_REPORT_NAME } elseif ($rows.Count) { $rows[0].User } else { $env:USERNAME }
$nameW = 9; foreach ($g in $items) { if ($g.Name.Length -gt $nameW) { $nameW = $g.Name.Length } }
$toolW = 10; foreach ($g in $items) { $t = $g.Tools -join ', '; if ($t.Length -gt $toolW) { $toolW = $t.Length } }
$width = [Math]::Max(72, 5 + $nameW + 3 + 7 + 3 + $toolW + 3 + 11 + 2)
$eq = '=' * $width; $dash = '-' * $width

$L = New-Object System.Collections.Generic.List[string]
$L.Add($eq)
$L.Add('  S4CARLISLE - DAILY WORK REPORT')
$L.Add($eq)
$L.Add('  Name      : ' + $user + '   (Windows user: ' + $env:USERNAME + ', PC: ' + $env:COMPUTERNAME + ')')
$L.Add('  Date      : ' + $day.ToString('dddd, dd MMMM yyyy', $ci))
$L.Add('  Generated : ' + (Get-Date).ToString('dd-MM-yyyy HH:mm', $ci))
$L.Add($dash)
if ($items.Count -eq 0) {
    $L.Add('  No files were saved by the tools on this day.')
} else {
    $hasDoc = @($items | Where-Object { $_.Unit -eq 'pages' }).Count -gt 0
    $cw = if ($hasDoc) { 13 } else { 7 }
    $L.Add('  ' + '#'.PadRight(3) + 'File name'.PadRight($nameW + 3) + $(if ($hasDoc) { 'Slides/Pages' } else { 'Slides' }).PadLeft($cw) + '   ' + 'Tools used'.PadRight($toolW + 3) + 'Time')
    $L.Add($dash)
    $i = 0
    foreach ($g in $items) {
        $i++
        $cnt = if ($g.Count -ne '') { $g.Count } else { '-' }
        if ($g.Unit -eq 'pages' -and $cnt -ne '-') { $cnt = $cnt + ' p' }
        $time = if ($g.First -eq $g.Last) { $g.First } else { $g.First + '-' + $g.Last }
        $L.Add('  ' + ([string]$i).PadRight(3) + $g.Name.PadRight($nameW + 3) + $cnt.PadLeft($cw) + '   ' + ($g.Tools -join ', ').PadRight($toolW + 3) + $time)
    }
    $L.Add($dash)
    $ppt = @($items | Where-Object { $_.Unit -eq 'slides' })
    $doc = @($items | Where-Object { $_.Unit -eq 'pages' })
    $slides = 0; foreach ($g in $ppt) { if ($g.Count -match '^\d+$') { $slides += [int]$g.Count } }
    $pages = 0; foreach ($g in $doc) { if ($g.Count -match '^\d+$') { $pages += [int]$g.Count } }
    $L.Add('  PowerPoint files : ' + ([string]$ppt.Count).PadRight(6) + 'Total slides : ' + $slides)
    if ($doc.Count) {
        $L.Add('  Word files       : ' + ([string]$doc.Count).PadRight(6) + 'Total pages  : ' + $pages)
        $L.Add('  (Word page counts are as last saved by Word; "p" = pages)')
    }
    $L.Add('  Total files      : ' + $items.Count)
}
$L.Add($eq)

if ($env:S4C_REPORT_DIR) { $dl = $env:S4C_REPORT_DIR }
else {
    $dl = if ($env:USERPROFILE) { Join-Path $env:USERPROFILE 'Downloads' } else { '' }
    if (-not $dl -or -not (Test-Path -LiteralPath $dl)) { $dl = [Environment]::GetFolderPath('Desktop') }
}
$safe = ($user -replace '[\\/:*?"<>|]', '').Trim(); if (-not $safe) { $safe = 'report' }
$file = Join-Path $dl ('Work report ' + $day.ToString('dd-MM-yyyy', $ci) + ' - ' + $safe + '.txt')
[IO.File]::WriteAllLines($file, $L.ToArray(), (New-Object Text.UTF8Encoding($true)))
$L | ForEach-Object { Write-Host $_ }
Write-Host ''
Write-Host ('Saved: ' + $file) -ForegroundColor Green
if (-not $env:S4C_REPORT_DATE) { try { Start-Process notepad.exe -ArgumentList ('"' + $file + '"') } catch {} }
