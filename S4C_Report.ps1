# S4C REPORT SCRIPT - makes the daily work report from the work log the tools keep on this PC.
# S4C-FILE-VERSION: 1.5
# The log is %LOCALAPPDATA%\S4C_Tools\WorkLog\worklog.tsv (one line per file a tool saved).
$ErrorActionPreference = 'Stop'
$ci = [Globalization.CultureInfo]::InvariantCulture
function To12([string]$t) { $d = [datetime]::MinValue; if ([datetime]::TryParseExact($t, 'HH:mm', $ci, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToString('h:mm tt', $ci) } return $t }
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
        $rows += [pscustomobject]@{ Time = $c[1]; User = $c[2]; Tool = $(switch ($c[3]) { 'Word Fixer' { 'Word Tool' } 'SLID Macro' { 'SLID Tool' } 'IMBK Macro' { 'IMBK Tool' } 'Issue Fixer' { 'Issues Fix' } 'Link Fixer' { 'Hyperlink Tool' } default { $c[3] } }); In = $c[4]; Out = $c[5]; Count = $c[6]; Unit = $c[7] }
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
    if ($r.Count -ne '') { $g.Count = $r.Count; $g.Unit = $r.Unit }
    if (-not $g.Tools.Contains($r.Tool)) { $g.Tools.Add($r.Tool) }
    if ([string]::CompareOrdinal($r.Time, $g.First) -lt 0) { $g.First = $r.Time }
    if ([string]::CompareOrdinal($r.Time, $g.Last) -gt 0) { $g.Last = $r.Time }
}
$items = @($groups.Values)

# Title = the part of the file name before the chapter / part / appendix marker (or a leading ISBN).
# Titles the person renamed on the page arrive as JSON in S4C_REPORT_TITLES (key = lower-case pattern).
$chapRx = '(?i)[\s_.\-]*\b(?:ch(?:apter)?|chap|app(?:endix)?|part|unit|sec(?:tion)?|front|back|preface|index|glossary|intro(?:duction)?|ex(?:ercise)?)s?[\s_.\-]*(?:\d+|[a-z]\b)?.*$'
$custom = @{}
if ($env:S4C_REPORT_TITLES) { try { $o = $env:S4C_REPORT_TITLES | ConvertFrom-Json; foreach ($pp in $o.PSObject.Properties) { $custom[$pp.Name.ToLowerInvariant()] = [string]$pp.Value } } catch {} }
function Get-TitleKey([string]$name) {
    $st = ([IO.Path]::GetFileNameWithoutExtension($name)) -replace $suffixRx, ''
    $k = ($st -replace $chapRx, '') -replace '[\s_.\-]+$', ''
    if (-not $k -or $k -eq $st) { $i = $st.IndexOf('_'); $k = if ($i -gt 2) { $st.Substring(0, $i) } else { '' } }
    if (-not $k) { $m = [regex]::Match(($st -replace '[-\s]', ''), '^(97[89]\d{10}|\d{9}[\dXx])'); if ($m.Success) { $k = $m.Groups[1].Value } }
    return $k.Trim()
}
foreach ($g in $items) {
    $k = (Get-TitleKey $g.Name).ToLowerInvariant()
    $t = if ($custom.ContainsKey($k)) { $custom[$k] } else { ((Get-TitleKey $g.Name) -replace '_', ' ') }
    if (-not $t) { $t = '(no title - name the files like "Book Title_CH01")' }
    $g | Add-Member -NotePropertyName Title -NotePropertyValue $t -Force
}
$titles = @($items | ForEach-Object { $_.Title } | Select-Object -Unique)

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
$L.Add('  Generated : ' + (Get-Date).ToString('dd-MM-yyyy h:mm tt', $ci))
$L.Add($dash)
if ($items.Count -eq 0) {
    $L.Add('  No files were saved by the tools on this day.')
} else {
    $hasDoc = @($items | Where-Object { $_.Unit -eq 'pages' }).Count -gt 0
    $cw = if ($hasDoc) { 13 } else { 7 }
    $L.Add('  ' + '#'.PadRight(3) + 'File name'.PadRight($nameW + 3) + $(if ($hasDoc) { 'Slides/Pages' } else { 'Slides' }).PadLeft($cw) + '   ' + 'Tools used'.PadRight($toolW + 3) + 'Time')
    $L.Add($dash)
    $i = 0
    foreach ($t in $titles) {
        $set = @($items | Where-Object { $_.Title -eq $t })
        $ss = 0; foreach ($g in $set) { if ($g.Unit -eq 'slides' -and $g.Count -match '^\d+$') { $ss += [int]$g.Count } }
        $pp2 = 0; foreach ($g in $set) { if ($g.Unit -eq 'pages' -and $g.Count -match '^\d+$') { $pp2 += [int]$g.Count } }
        $sum = [string]$set.Count + ' file' + $(if ($set.Count -ne 1) { 's' }) + $(if ($ss) { ', ' + $ss + ' slides' }) + $(if ($pp2) { ', ' + $pp2 + ' pages' })
        $L.Add('  TITLE: ' + $t + '   (' + $sum + ')')
        foreach ($g in $set) {
            $i++
            $cnt = if ($g.Count -ne '') { $g.Count } else { '-' }
            if ($g.Unit -eq 'pages' -and $cnt -ne '-') { $cnt = $cnt + ' p' }
            $time = if ($g.First -eq $g.Last) { To12 $g.First } else { (To12 $g.First) + ' - ' + (To12 $g.Last) }
            $L.Add('  ' + ([string]$i).PadRight(3) + $g.Name.PadRight($nameW + 3) + $cnt.PadLeft($cw) + '   ' + ($g.Tools -join ', ').PadRight($toolW + 3) + $time)
        }
        $L.Add('')
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
