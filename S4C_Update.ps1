# S4C UPDATE SCRIPT - checks the 'testing' branch on GitHub for ONE tool and replaces the local copy
# S4C-FILE-VERSION: 1.6
# only when GitHub has a HIGHER version number. An older or same-numbered file on GitHub never replaces the copy here.
# Used by "S4C Tools.cmd" for the tool you pick. Never stops the tool from running.
param([string]$Tool)
$info = @{
    'Word_Fixer.ps1'   = @('Word Tool',    'WORD FIXER SCRIPT',    "WF_VERSION\s*=\s*'([^']+)'")
    'IMBK_Macro.ps1'   = @('IMBK Tool',    'IMBK FIGURE LAYOUT',   "IMBK_VERSION\s*=\s*'([^']+)'")
    'SLID_Macro.ps1'   = @('SLID Tool',    'SLID MACRO SCRIPT',    "SLID_VERSION\s*=\s*'([^']+)'")
    'AltText_Tool.ps1' = @('Alt Text Tool', 'ALT TEXT TOOL SCRIPT', "ALT_VERSION\s*=\s*'([^']+)'")
    'Issue_Fixer.ps1'  = @('Issues Fix',   'S4C ISSUE FIXER SCRIPT', "IF_VERSION\s*=\s*'([^']+)'")
    'Link_Fixer.ps1'   = @('Hyperlink Tool',    '# Link_Fixer.ps1',     '(?m)^\$LF_VERSION\s*=\s*(\d+)')
}[$Tool]
if (-not $info) { Write-Host ('Unknown tool: ' + $Tool) -ForegroundColor Red; exit 0 }
$File = $Tool; $Name = $info[0]; $Marker = $info[1]; $VerRx = $info[2]
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # the progress bar makes Windows PowerShell download very slowly
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$bases = @(
    'https://raw.githubusercontent.com/rsashwin225/Fix-Hyperlink/refs/heads/testing/',
    'https://github.com/rsashwin225/Fix-Hyperlink/raw/refs/heads/testing/'
)
if ($env:S4C_BASE) { $bases = @($env:S4C_BASE) }

# S4C Tools (the window version) reads the result from this small file when it asks for one
function Set-Result([string]$status, [string]$msg, [string]$reason) {
    if ($env:S4C_STATUS_FILE) { try { [IO.File]::WriteAllText($env:S4C_STATUS_FILE, ($status + "`t" + $msg + "`t" + $reason), (New-Object Text.UTF8Encoding($false))) } catch {} }
}
function Get-Ver([string]$t) { $m = [regex]::Match($t, $VerRx); if ($m.Success) { $m.Groups[1].Value } else { 'unknown' } }

# Try every way we know, for up to 2 minutes in total:
# two GitHub addresses x two download methods (both use the Windows proxy with your sign-in), then one more round.
function Get-GitHubText {
    $script:why = ''
    $deadline = (Get-Date).AddSeconds(120)
    try { $px = [Net.WebRequest]::GetSystemWebProxy(); $px.Credentials = [Net.CredentialCache]::DefaultCredentials; [Net.WebRequest]::DefaultWebProxy = $px } catch {}
    for ($round = 1; $round -le 2; $round++) {
        foreach ($b in $bases) {
            $url = $b + $File + '?t=' + [DateTime]::UtcNow.Ticks
            $left = [int]($deadline - (Get-Date)).TotalSeconds
            if ($left -lt 5) { return $null }
            $limit = [math]::Min(45, $left)
            try {
                $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $limit
                $t = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
                if ($t) { return $t.TrimStart([char]0xFEFF) }
            } catch { $script:why = $_.Exception.Message }
            $left = [int]($deadline - (Get-Date)).TotalSeconds
            if ($left -lt 5) { return $null }
            try {
                $wc = New-Object Net.WebClient
                $wc.Proxy = [Net.WebRequest]::DefaultWebProxy
                $wc.Headers.Add('User-Agent', 'S4C-Tools')
                $wc.Encoding = [Text.Encoding]::UTF8
                $t = $wc.DownloadString($url)
                if ($t) { return $t.TrimStart([char]0xFEFF) }
            } catch { $script:why = $_.Exception.Message }
        }
        if ($round -eq 1) { Start-Sleep -Seconds 3 }
    }
    return $null
}

Write-Host ($Name + ' - checking GitHub for updates (can take up to a minute)...') -ForegroundColor Cyan
$local = Join-Path $PSScriptRoot $File
$old = if (Test-Path -LiteralPath $local) { [IO.File]::ReadAllText($local) } else { '' }
$oldVer = Get-Ver $old
$new = Get-GitHubText
if ($null -eq $new) {
    Write-Host ('  Could not reach GitHub - using the copy already here (version ' + $oldVer + ').') -ForegroundColor Yellow
    if ($why) { Write-Host ('  (' + $why + ')') -ForegroundColor DarkGray }
    Set-Result 'offline' ('Not connected - using version ' + $oldVer + ' on this PC') $why
    Write-Host ''
    exit 0
}
if ($new.Length -lt 3000 -or $new -notmatch [regex]::Escape($Marker)) {
    Set-Result 'kept' ('Connected, but the GitHub file looks incomplete - kept ' + $oldVer) ''
    Write-Host ('  Connected to GitHub, but the file there looks incomplete - kept your version ' + $oldVer + '.') -ForegroundColor Yellow; Write-Host ''; exit 0
}
$errs = $null; $tok = $null
[void][Management.Automation.Language.Parser]::ParseInput($new, [ref]$tok, [ref]$errs)
if ($errs -and $errs.Count -gt 0) {
    Set-Result 'kept' ('Connected, but the GitHub file has errors - kept ' + $oldVer) ''
    Write-Host ('  Connected to GitHub, but the file there has errors - kept your version ' + $oldVer + '.') -ForegroundColor Yellow; Write-Host ''; exit 0
}
$newVer = Get-Ver $new
$norm = { param($t) ($t -replace "`r`n", "`n").TrimEnd() }
function ConvertTo-Ver([string]$v) { try { if ($v -match '^\d+$') { $v += '.0' }; return [version]$v } catch { return $null } }
$vo = ConvertTo-Ver $oldVer; $vn = ConvertTo-Ver $newVer
if (-not $old -or ($vn -and $vo -and $vn -gt $vo) -or ($vn -and -not $vo)) {
    [IO.File]::WriteAllText($local, $new, (New-Object Text.UTF8Encoding($false)))
    Write-Host ('  Connected to GitHub. UPDATED: version ' + $oldVer + ' -> ' + $newVer) -ForegroundColor Green
    Set-Result 'updated' ('Updated ' + $oldVer + ' -> ' + $newVer) ''
} elseif ($vn -and $vo -and $vn -lt $vo) {
    Write-Host ('  Connected to GitHub. GitHub has an older version (' + $newVer + ') - kept your version ' + $oldVer + '.') -ForegroundColor Green
    Set-Result 'newer' ('GitHub has ' + $newVer + ' - kept your newer ' + $oldVer) ''
} else {
    Write-Host ('  Connected to GitHub. You already have the latest version (' + $oldVer + ').') -ForegroundColor Green
    Set-Result 'latest' ('Newest version (' + $oldVer + ')') ''
}
Write-Host ''
