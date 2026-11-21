# WORD FIXER UPDATER
# Checks GitHub for a newer Word_Fixer.ps1 (every run) and a newer copy of this updater (once a day),
# and replaces a local copy only if the download is complete, error-free and different.
# After this file is in place, every future change - to the macro AND to this updater - comes from GitHub.
# Never stops the macro from running: if nothing can be reached, the copy already here is used.
#
# Some networks block raw.githubusercontent.com, so several addresses that serve the same GitHub file are tried:
#   1. raw.githubusercontent.com   2. api.github.com   3. cdn.jsdelivr.net (public mirror of GitHub, can be up to 12 h behind)
# The address that worked last time is tried first, so the check stays quick.
$UPD_VERSION = '2.0'
$ErrorActionPreference = 'Stop'

$Owner = 'rsashwin225'; $Repo = 'Fix-Hyperlink'; $Branch = 'main'
$TimeoutMs = 7000      # per try
$BudgetSec = 40        # whole check gives up after this, so the macro always starts

$script:start = Get-Date
$script:state = @{ source = ''; selfCheck = '' }
$stateDir = Join-Path $env:LOCALAPPDATA 'S4C-Tools'
$statePath = Join-Path $stateDir 'word_fixer_update.json'

function Get-Ver([string]$text, [string]$varName) {
    $m = [regex]::Match($text, [regex]::Escape($varName) + "\s*=\s*'([^']+)'")
    if ($m.Success) { return $m.Groups[1].Value } else { return 'unknown' }
}

function Get-Sources([string]$file) {
    return @(
        @{ Name = 'raw.githubusercontent.com'; Accept = $null
           Url = 'https://raw.githubusercontent.com/' + $Owner + '/' + $Repo + '/refs/heads/' + $Branch + '/' + $file },
        @{ Name = 'api.github.com'; Accept = 'application/vnd.github.raw'
           Url = 'https://api.github.com/repos/' + $Owner + '/' + $Repo + '/contents/' + $file + '?ref=' + $Branch },
        @{ Name = 'cdn.jsdelivr.net'; Accept = $null
           Url = 'https://cdn.jsdelivr.net/gh/' + $Owner + '/' + $Repo + '@' + $Branch + '/' + $file }
    )
}

# way 1: through the computer's proxy settings with the Windows sign-in;  way 2: straight out, no proxy
function Invoke-Fetch([hashtable]$src, [int]$way) {
    $rq = [Net.HttpWebRequest]::Create($src.Url)
    $rq.Timeout = $TimeoutMs; $rq.ReadWriteTimeout = $TimeoutMs; $rq.UserAgent = 'S4C-Tools'
    if ($src.Accept) { $rq.Accept = $src.Accept }
    if ($way -eq 1) {
        try { $p = [Net.WebRequest]::GetSystemWebProxy(); $p.Credentials = [Net.CredentialCache]::DefaultCredentials; $rq.Proxy = $p } catch { }
    } else { $rq.Proxy = $null }
    $rs = $rq.GetResponse()
    try {
        $sr = New-Object IO.StreamReader($rs.GetResponseStream(), [Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    } finally { $rs.Close() }
}

# Downloads one file from the first address that answers with the expected content.
function Get-Remote([string]$file, [string]$marker, [int]$minLen) {
    $srcs = Get-Sources $file
    if ($script:state.source) {
        $srcs = @($srcs | Where-Object { $_.Name -eq $script:state.source }) + @($srcs | Where-Object { $_.Name -ne $script:state.source })
    }
    $errs = @()
    foreach ($s in $srcs) {
        foreach ($way in 1, 2) {
            if (((Get-Date) - $script:start).TotalSeconds -gt $BudgetSec) {
                $errs += 'stopped trying (too slow)'; throw ($errs -join ' | ')
            }
            try {
                $t = Invoke-Fetch $s $way
                if ($null -eq $t) { $t = '' }
                $t = $t.TrimStart([char]0xFEFF)
                if ($t.Length -lt $minLen -or $t.IndexOf($marker) -lt 0) { throw 'the answer was not the expected file' }
                $script:state.source = $s.Name
                return $t
            } catch {
                $msg = $_.Exception.Message
                if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
                $errs += ($s.Name + ' way ' + $way + ': ' + $msg)
                # these answers will not change by skipping the proxy, so go to the next address
                if ($msg -match '\(404\)|\(403\)|could not be resolved|not the expected file') { break }
            }
        }
    }
    throw ($errs -join ' | ')
}

# Replaces $local with $new if $new parses cleanly and differs. Returns 'same', 'updated' or 'errors'.
function Update-LocalFile([string]$local, [string]$new) {
    $perr = $null; $tok = $null
    [void][Management.Automation.Language.Parser]::ParseInput($new, [ref]$tok, [ref]$perr)
    if ($perr -and $perr.Count -gt 0) { return 'errors' }
    $old = ''
    if (Test-Path -LiteralPath $local) { $old = [IO.File]::ReadAllText($local) }
    $a = ($new -replace "`r`n", "`n").TrimEnd(); $b = ($old -replace "`r`n", "`n").TrimEnd()
    if ($a -eq $b) { return 'same' }
    $tmp = $local + '.download'
    [IO.File]::WriteAllText($tmp, $new, (New-Object Text.UTF8Encoding($true)))
    Move-Item -LiteralPath $tmp -Destination $local -Force
    return 'updated'
}

Write-Host 'Word Fixer - checking GitHub for updates...' -ForegroundColor Cyan
try {
    if (Test-Path -LiteralPath $statePath) {
        $j = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($j.source) { $script:state.source = [string]$j.source }
        if ($j.selfCheck) { $script:state.selfCheck = [string]$j.selfCheck }
    }
} catch { }
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

# ---- 1. the macro itself (every run)
$local = Join-Path $PSScriptRoot 'Word_Fixer.ps1'
$oldVer = 'unknown'
if (Test-Path -LiteralPath $local) { $oldVer = Get-Ver ([IO.File]::ReadAllText($local)) 'WF_VERSION' }
$reached = $false
try {
    $new = Get-Remote 'Word_Fixer.ps1' 'WORD FIXER SCRIPT' 5000
    $reached = $true
    $newVer = Get-Ver $new 'WF_VERSION'
    $r = Update-LocalFile $local $new
    $via = ' (via ' + $script:state.source + ')'
    if ($r -eq 'updated') {
        Write-Host ('  Connected to GitHub' + $via + '. UPDATED: version ' + $oldVer + ' -> ' + $newVer) -ForegroundColor Green
    } elseif ($r -eq 'same') {
        Write-Host ('  Connected to GitHub' + $via + '. You already have the latest version (' + $newVer + ').') -ForegroundColor Green
    } else {
        Write-Host ('  GitHub file has errors - kept your version ' + $oldVer + '.') -ForegroundColor Yellow
    }
} catch {
    Write-Host '  Could not reach GitHub - using the copy already here.' -ForegroundColor Yellow
    Write-Host ('  (' + $_.Exception.Message + ')') -ForegroundColor DarkGray
}

# ---- 2. this updater (once a day, only if GitHub was reachable just now; takes effect next run)
$today = (Get-Date).ToString('yyyy-MM-dd')
if ($reached -and $script:state.selfCheck -ne $today) {
    try {
        $selfNew = Get-Remote 'Word_Fixer_Update.ps1' 'WORD FIXER UPDATER' 1000
        $r2 = Update-LocalFile $PSCommandPath $selfNew
        if ($r2 -eq 'updated') { Write-Host ('  The updater was refreshed too (version ' + (Get-Ver $selfNew 'UPD_VERSION') + ').') -ForegroundColor Green }
        $script:state.selfCheck = $today
    } catch { }
}

try {
    if (-not (Test-Path -LiteralPath $stateDir)) { [void](New-Item -ItemType Directory -Path $stateDir -Force) }
    (@{ source = $script:state.source; selfCheck = $script:state.selfCheck } | ConvertTo-Json) | Set-Content -LiteralPath $statePath -Encoding UTF8
} catch { }
Write-Host ''
exit 0
