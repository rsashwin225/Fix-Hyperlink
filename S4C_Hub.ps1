# S4C HUB SCRIPT - the S4C Tools window. Runs a small page on this PC only (http://localhost) and opens it
# S4C-FILE-VERSION: 1.2
# in the browser. From the page you pick a tool, it checks GitHub for that one tool, then opens the tool
# in its own window exactly as before. Close this window to close S4C Tools.
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$exe = (Get-Process -Id $PID).Path
$ci = [Globalization.CultureInfo]::InvariantCulture

$TOOLS = [ordered]@{
    'Word_Fixer.ps1'   = @{ Name = 'Word Fixer';    Rx = "WF_VERSION\s*=\s*'([^']+)'" }
    'SLID_Macro.ps1'   = @{ Name = 'SLID Macro';    Rx = "SLID_VERSION\s*=\s*'([^']+)'" }
    'IMBK_Macro.ps1'   = @{ Name = 'IMBK Macro';    Rx = "IMBK_VERSION\s*=\s*'([^']+)'" }
    'AltText_Tool.ps1' = @{ Name = 'Alt Text Tool'; Rx = "ALT_VERSION\s*=\s*'([^']+)'" }
    'Link_Fixer.ps1'   = @{ Name = 'Link Fixer';    Rx = '(?m)^\$LF_VERSION\s*=\s*(\d+)' }
}
function Get-ToolVersion([string]$file) {
    $p = Join-Path $here $file
    if (-not (Test-Path -LiteralPath $p)) { return '' }
    $m = [regex]::Match([IO.File]::ReadAllText($p), $TOOLS[$file].Rx)
    if ($m.Success) { return $m.Groups[1].Value } else { return '' }
}
function Get-LogRows([string]$day) {
    $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
    $log = Join-Path (Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog') 'worklog.tsv'
    $rows = New-Object System.Collections.Generic.List[object]
    if (Test-Path -LiteralPath $log) {
        foreach ($l in [IO.File]::ReadAllLines($log, [Text.Encoding]::UTF8)) {
            $c = $l.Split("`t")
            if ($c.Count -lt 8 -or $c[0] -ne $day) { continue }
            $rows.Add([ordered]@{ time = $c[1]; user = $c[2]; tool = $c[3]; file = [IO.Path]::GetFileName($c[4]); out = [IO.Path]::GetFileName($c[5]); count = $c[6]; unit = $c[7] })
        }
    }
    return , $rows.ToArray()
}
function Read-Status([string]$text) {
    $st = 'unknown'; $msg = ''; $reason = ''
    foreach ($line in ($text -split "`r?`n")) {
        $t = $line.Trim()
        if ($t -match 'UPDATED: version (.+)$') { $st = 'updated'; $msg = 'Updated ' + $Matches[1] }
        elseif ($t -match 'already have the latest version \((.+)\)') { $st = 'latest'; $msg = 'Newest version (' + $Matches[1] + ')' }
        elseif ($t -match '^Could not reach GitHub') { $st = 'offline'; $msg = 'Not connected - using the copy here' }
        elseif ($t -match 'looks incomplete|has errors') { $st = 'kept'; $msg = 'Connected, kept your copy' }
        elseif ($st -eq 'offline' -and $t -match '^\((.+)\)$') { $reason = $Matches[1] }
    }
    return [ordered]@{ status = $st; message = $msg; reason = $reason }
}

# ---------------------------------------------------------------- jobs: GitHub check, then the tool in its own window
$jobs = @{}
function Start-Job2([string]$file, [bool]$launch) {
    $id = [guid]::NewGuid().ToString('N').Substring(0, 10)
    $out = Join-Path ([IO.Path]::GetTempPath()) ("s4c_upd_$id.txt"); $err = $out + '.err'
    $args1 = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'S4C_Update.ps1') + '" -Tool ' + $file
    $sf = $out + '.status'
    $env:S4C_STATUS_FILE = $sf
    $p = Start-Process -FilePath $exe -ArgumentList $args1 -NoNewWindow -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    Remove-Item Env:S4C_STATUS_FILE -ErrorAction SilentlyContinue
    $jobs[$id] = @{ Id = $id; File = $file; Launch = $launch; State = 'updating'; Upd = $p; Out = $out; Err = $err; Sf = $sf; Status = $null; Tool = $null; Started = (Get-Date) }
    return $id
}
function Step-Job($j) {
    if ($j.State -eq 'updating' -and $j.Upd.HasExited) {
        $txt = ''; try { $txt = [IO.File]::ReadAllText($j.Out) } catch {}
        $j.Status = Read-Status $txt
        if (Test-Path -LiteralPath $j.Sf) {
            $c = [IO.File]::ReadAllText($j.Sf).Split("`t")
            if ($c.Count -ge 2) { $j.Status = [ordered]@{ status = $c[0]; message = $c[1]; reason = $(if ($c.Count -ge 3) { $c[2] } else { '' }) } }
        }
        Remove-Item -LiteralPath $j.Out, $j.Err, $j.Sf -ErrorAction SilentlyContinue
        if ($j.Launch) {
            # the tool window repeats the GitHub result at the top, so both windows say the same thing
            $env:S4C_GH_STATUS = [string]$j.Status.status; $env:S4C_GH_MSG = [string]$j.Status.message; $env:S4C_GH_REASON = [string]$j.Status.reason
            $a = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'S4C_RunTool.ps1') + '" -Tool ' + $j.File
            if ($env:S4C_HUB_TEST) { $j.Tool = Start-Process -FilePath $exe -ArgumentList ($a + ' -Test') -NoNewWindow -PassThru }
            else { $j.Tool = Start-Process -FilePath $exe -ArgumentList $a -PassThru }
            Remove-Item Env:S4C_GH_STATUS, Env:S4C_GH_MSG, Env:S4C_GH_REASON -ErrorAction SilentlyContinue
            $j.State = 'running'
        } else { $j.State = 'done' }
    } elseif ($j.State -eq 'running' -and $j.Tool.HasExited) { $j.State = 'done' }
}

# ---------------------------------------------------------------- the page server
$token = [guid]::NewGuid().ToString('N')
$listener = $null; $port = 0
foreach ($try in 1..40) {
    $port = Get-Random -Minimum 49200 -Maximum 49900
    if ($env:S4C_HUB_PORT) { $port = [int]$env:S4C_HUB_PORT }
    $l = New-Object System.Net.HttpListener
    $l.Prefixes.Add("http://localhost:$port/")
    try { $l.Start(); $listener = $l; break } catch { $l.Close() }
}
if (-not $listener) { Write-Host 'Could not start S4C Tools (no free port).' -ForegroundColor Red; exit 1 }
$url = "http://localhost:$port/"
Write-Host ''
Write-Host '  S4C Tools is open in your browser.' -ForegroundColor Green
Write-Host ('  Address: ' + $url)
Write-Host '  Keep this window open while you work. Close it to close S4C Tools.' -ForegroundColor DarkGray
Write-Host ''
if (-not $env:S4C_HUB_TEST) { Start-Process $url }

# the page files keep themselves up to date from GitHub in the background (the page opens straight away)
$selfStatus = Join-Path ([IO.Path]::GetTempPath()) ('s4c_self_' + $token.Substring(0, 8) + '.txt')
$selfProc = $null
if (Test-Path -LiteralPath (Join-Path $here 'S4C_SelfUpdate.ps1')) {
    $env:S4C_SELF_STATUS = $selfStatus
    try { $selfProc = Start-Process -FilePath $exe -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'S4C_SelfUpdate.ps1') + '"') -NoNewWindow -PassThru } catch {}
    Remove-Item Env:S4C_SELF_STATUS -ErrorAction SilentlyContinue
}

function Send($res, [int]$code, [string]$ctype, $body, [string]$fileName) {
    $bytes = if ($body -is [byte[]]) { $body } else { [Text.Encoding]::UTF8.GetBytes([string]$body) }
    $res.StatusCode = $code; $res.ContentType = $ctype
    $res.Headers.Add('Cache-Control', 'no-store')
    if ($fileName) { $res.Headers.Add('Content-Disposition', 'attachment; filename="' + $fileName + '"') }
    $res.ContentLength64 = $bytes.Length
    $res.OutputStream.Write($bytes, 0, $bytes.Length); $res.OutputStream.Close()
}
function ToJson($o) { return ($o | ConvertTo-Json -Depth 6 -Compress) }

$quit = $false
while ($listener.IsListening -and -not $quit) {
    $ctx = $listener.GetContext(); $req = $ctx.Request; $res = $ctx.Response
    try {
        $path = $req.Url.AbsolutePath
        $q = @{}; foreach ($k in $req.QueryString.AllKeys) { if ($k) { $q[$k] = $req.QueryString[$k] } }
        if ($path -eq '/' -or $path -eq '/index.html') {
            $html = [IO.File]::ReadAllText((Join-Path $here 'S4C_Hub.html'), [Text.Encoding]::UTF8).Replace('__S4C_TOKEN__', $token)
            Send $res 200 'text/html; charset=utf-8' $html $null; continue
        }
        if ($path.StartsWith('/api/') -and $req.Headers['X-S4C'] -ne $token -and $q['t'] -ne $token) { Send $res 403 'text/plain' 'no' $null; continue }
        switch ($path) {
            '/api/state' {
                $day = if ($q['day']) { $q['day'] } else { (Get-Date).ToString('yyyy-MM-dd', $ci) }
                $tl = @(); foreach ($k in $TOOLS.Keys) { $tl += [ordered]@{ file = $k; version = (Get-ToolVersion $k) } }
                $body = [ordered]@{ day = $day; today = (Get-Date).ToString('yyyy-MM-dd', $ci); user = $env:USERNAME; pc = $env:COMPUTERNAME; tools = $tl; rows = (Get-LogRows $day) }
                Send $res 200 'application/json; charset=utf-8' (ToJson $body) $null
            }
            '/api/days' {
                # files per day for the last 60 days (for the streak); a chapter run through several tools counts once
                $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
                $log = Join-Path (Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog') 'worklog.tsv'
                $from = (Get-Date).Date.AddDays(-60).ToString('yyyy-MM-dd', $ci)
                $map = @{}
                if (Test-Path -LiteralPath $log) {
                    foreach ($l in [IO.File]::ReadAllLines($log, [Text.Encoding]::UTF8)) {
                        $c = $l.Split("`t"); if ($c.Count -lt 8 -or [string]::CompareOrdinal($c[0], $from) -lt 0) { continue }
                        $leaf = [IO.Path]::GetFileName($c[4])
                        $k = ([IO.Path]::GetFileNameWithoutExtension($leaf) -replace '(?i)(\s+(Remediated|AltText|Links|Fixed)|_BEFORE_IMBKfix\d*)+$', '').ToLowerInvariant()
                        if (-not $map.ContainsKey($c[0])) { $map[$c[0]] = @{} }
                        $map[$c[0]][$k] = 1
                    }
                }
                $list = @(); foreach ($d in ($map.Keys | Sort-Object)) { $list += [ordered]@{ day = $d; files = $map[$d].Count } }
                Send $res 200 'application/json' (ToJson @{ days = $list }) $null
            }
            '/api/run' {
                $f = $q['tool']; if (-not $TOOLS.Contains($f)) { Send $res 400 'text/plain' 'unknown tool' $null; break }
                $id = Start-Job2 $f ($q['launch'] -ne '0')
                Send $res 200 'application/json' (ToJson @{ id = $id }) $null
            }
            '/api/job' {
                $j = $jobs[$q['id']]
                if (-not $j) { Send $res 404 'text/plain' 'no job' $null; break }
                Step-Job $j
                $body = [ordered]@{ id = $j.Id; tool = $j.File; state = $j.State; update = $j.Status; version = (Get-ToolVersion $j.File) }
                Send $res 200 'application/json' (ToJson $body) $null
            }
            '/api/report' {
                $day = [datetime]::ParseExact($q['day'], 'yyyy-MM-dd', $ci)
                $dir = Join-Path ([IO.Path]::GetTempPath()) ('s4c_rep_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
                [void](New-Item -ItemType Directory -Path $dir -Force)
                $env:S4C_REPORT_DATE = $day.ToString('dd-MM-yyyy', $ci); $env:S4C_REPORT_DIR = $dir
                if ($q['titles']) { $env:S4C_REPORT_TITLES = $q['titles'] } else { Remove-Item Env:S4C_REPORT_TITLES -ErrorAction SilentlyContinue }
                if ($q['name']) { $env:S4C_REPORT_NAME = $q['name'] } else { Remove-Item Env:S4C_REPORT_NAME -ErrorAction SilentlyContinue }
                & $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'S4C_Report.ps1') | Out-Null
                Remove-Item Env:S4C_REPORT_DATE, Env:S4C_REPORT_DIR, Env:S4C_REPORT_NAME, Env:S4C_REPORT_TITLES -ErrorAction SilentlyContinue
                $rf = Get-ChildItem -LiteralPath $dir -Filter '*.txt' | Select-Object -First 1
                if (-not $rf) { Send $res 500 'text/plain' 'The report could not be made.' $null; break }
                $bytes = [IO.File]::ReadAllBytes($rf.FullName)
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
                Send $res 200 'text/plain; charset=utf-8' $bytes $rf.Name
            }
            '/api/hub' {
                $done = (-not $selfProc) -or $selfProc.HasExited
                $st = ''; $upd = @()
                if ($done -and (Test-Path -LiteralPath $selfStatus)) { $lines = [IO.File]::ReadAllText($selfStatus).Split("`n") | Where-Object { $_ }; if ($lines) { $st = $lines[0]; $upd = @($lines | Select-Object -Skip 1) } }
                $hv = [regex]::Match([IO.File]::ReadAllText((Join-Path $here 'S4C_Hub.html')), 'name="s4c-hub-version" content="([0-9.]+)"').Groups[1].Value
                Send $res 200 'application/json' (ToJson ([ordered]@{ done = $done; status = $st; updated = $upd; version = $hv })) $null
            }
            '/api/quit' { Send $res 200 'text/plain' 'bye' $null; $quit = $true }
            default { Send $res 404 'text/plain' 'not found' $null }
        }
    } catch {
        try { Send $res 500 'text/plain' $_.Exception.Message $null } catch {}
    }
}
$listener.Stop()
