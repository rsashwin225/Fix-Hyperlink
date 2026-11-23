# S4C SELFUPDATE SCRIPT - keeps the S4C Tools page files up to date from the 'testing' branch on GitHub.
# S4C-FILE-VERSION: 1.1
# Started in the background by S4C_Hub.ps1. A file is replaced only when GitHub has a HIGHER version number
# and the downloaded file is complete. The tool scripts are not touched here (each tool is checked when it is run).
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
$bases = @(
    'https://raw.githubusercontent.com/rsashwin225/Fix-Hyperlink/refs/heads/testing/',
    'https://github.com/rsashwin225/Fix-Hyperlink/raw/refs/heads/testing/'
)
if ($env:S4C_BASE) { $bases = @($env:S4C_BASE) }
$verPs = '# S4C-FILE-VERSION: ([0-9.]+)'
$files = @(
    @{ F = 'S4C_Hub.html';       M = 'name="s4c-hub-version"'; R = 'name="s4c-hub-version" content="([0-9.]+)"'; Ps = $false },
    @{ F = 'S4C_Hub.ps1';        M = 'S4C HUB SCRIPT';        R = $verPs; Ps = $true },
    @{ F = 'S4C_RunTool.ps1';    M = 'S4C RUNTOOL SCRIPT';    R = $verPs; Ps = $true },
    @{ F = 'S4C_Report.ps1';     M = 'S4C REPORT SCRIPT';     R = $verPs; Ps = $true },
    @{ F = 'S4C_Update.ps1';     M = 'S4C UPDATE SCRIPT';     R = $verPs; Ps = $true },
    @{ F = 'S4C_SelfUpdate.ps1'; M = 'S4C SELFUPDATE SCRIPT'; R = $verPs; Ps = $true }
)
function V([string]$t, [string]$rx) { $m = [regex]::Match($t, $rx); if ($m.Success) { try { return [version]$m.Groups[1].Value } catch {} }; return $null }
try { $px = [Net.WebRequest]::GetSystemWebProxy(); $px.Credentials = [Net.CredentialCache]::DefaultCredentials; [Net.WebRequest]::DefaultWebProxy = $px } catch {}
$updated = @(); $reached = $false
foreach ($f in $files) {
    $local = Join-Path $PSScriptRoot $f.F
    $old = if (Test-Path -LiteralPath $local) { [IO.File]::ReadAllText($local) } else { '' }
    $new = $null
    foreach ($b in $bases) {
        try {
            $r = Invoke-WebRequest -Uri ($b + $f.F + '?t=' + [DateTime]::UtcNow.Ticks) -UseBasicParsing -TimeoutSec 30
            $new = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
            $reached = $true; break
        } catch { if ($_.Exception.Message -match '404|Not Found') { $reached = $true; break } }
    }
    if (-not $new) { continue }
    $new = $new.TrimStart([char]0xFEFF)
    if ($new.Length -lt 1000 -or $new -notmatch [regex]::Escape($f.M)) { continue }
    if ($f.Ps) { $e = $null; $t = $null; [void][Management.Automation.Language.Parser]::ParseInput($new, [ref]$t, [ref]$e); if ($e -and $e.Count) { continue } }
    $vn = V $new $f.R; $vo = V $old $f.R
    if ($vn -and (-not $vo -or $vn -gt $vo)) {
        [IO.File]::WriteAllText($local, $new, (New-Object Text.UTF8Encoding($false)))
        $updated += ($f.F + ' ' + $(if ($vo) { $vo } else { 'new' }) + ' -> ' + $vn)
    }
}
if ($env:S4C_SELF_STATUS) {
    $txt = $(if ($reached) { 'connected' } else { 'offline' }) + "`n" + ($updated -join "`n")
    try { [IO.File]::WriteAllText($env:S4C_SELF_STATUS, $txt, (New-Object Text.UTF8Encoding($false))) } catch {}
}
