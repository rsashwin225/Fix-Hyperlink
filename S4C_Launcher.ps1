# S4C LAUNCHER SCRIPT
# S4Carlisle remediation tools - one menu for all tools. The team keeps ONE file (S4C Tools.cmd).
# This script is kept on GitHub; S4C Tools.cmd fetches the newest copy every time it starts, and each tool
# is fetched from GitHub the moment it is chosen. Nothing has to be sent to the team again.
$S4C_LAUNCHER_VERSION = '1.0'

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # the download progress bar makes Windows PowerShell download very slowly
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$S4C_HOME = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [Environment]::GetFolderPath('LocalApplicationData') }) 'S4C_Tools'
if (-not (Test-Path -LiteralPath $S4C_HOME)) { [void](New-Item -ItemType Directory -Path $S4C_HOME -Force) }

# ---- where the files come from (the first address that answers wins) -----------------------------------------
$S4C_BASES = @(
    'https://raw.githubusercontent.com/rsashwin225/Fix-Hyperlink/refs/heads/main/',
    'https://github.com/rsashwin225/Fix-Hyperlink/raw/refs/heads/main/',
    'https://cdn.jsdelivr.net/gh/rsashwin225/Fix-Hyperlink@main/'
)
if ($env:S4C_BASE) { $S4C_BASES = @($env:S4C_BASE.Split(';') | Where-Object { $_ }) }

# ---- download that copes with company networks ---------------------------------------------------------------
# - progress bar off (it makes Windows PowerShell very slow)
# - the Windows proxy with the signed-in user's credentials (many company networks need that)
# - generous time limits, two different download methods, then the next address
# Returns the text, or $null when nothing worked ($script:LastError says why).
function Get-Remote([string]$file, [int]$seconds = 60) {
    $script:LastError = ''
    $deadline = (Get-Date).AddSeconds([math]::Max(120, $seconds * 2))
    try { [Net.WebRequest]::DefaultWebProxy = [Net.WebRequest]::GetSystemWebProxy(); [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultCredentials } catch {}
    $first = $true
    foreach ($base in $S4C_BASES) {
        if ((Get-Date) -gt $deadline) { break }
        $url = $base + $file
        if ($base -notmatch '^file:|localhost|127\.0\.0\.1') { $url += '?nocache=' + [DateTime]::UtcNow.Ticks }
        $limit = if ($first) { $seconds } else { [math]::Min(30, $seconds) }
        $first = $false
        # method 1: Invoke-WebRequest
        try {
            $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $limit
            $t = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
            if ($t) { return $t.TrimStart([char]0xFEFF) }
        } catch { $script:LastError = $_.Exception.Message }
        # method 2: a plain web request that sends the Windows sign-in to the proxy
        try {
            $q = [Net.HttpWebRequest]::Create($url)
            $q.Timeout = $limit * 1000; $q.ReadWriteTimeout = $limit * 1000; $q.UserAgent = 'S4C-Tools'
            $q.Proxy = [Net.WebRequest]::GetSystemWebProxy(); $q.Proxy.Credentials = [Net.CredentialCache]::DefaultCredentials
            $resp = $q.GetResponse()
            try { $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close() } finally { $resp.Close() }
            if ($t) { return $t.TrimStart([char]0xFEFF) }
        } catch { $script:LastError = $_.Exception.Message }
    }
    return $null
}

# ---- the tools -----------------------------------------------------------------------------------------------
$TOOLS = @(
    @{ Key = '1'; Name = 'Word Fixer';   File = 'Word_Fixer.ps1';   Marker = 'WORD FIXER SCRIPT';       VerRx = "WF_VERSION\s*=\s*'([^']+)'";     Info = 'Headings and list levels in Word (.docx) files' },
    @{ Key = '2'; Name = 'IMBK Macro';   File = 'IMBK_Macro.ps1';   Marker = 'IMBK FIGURE LAYOUT';      VerRx = "IMBK_VERSION\s*=\s*'([^']+)'";   Info = 'IMBK chapter decks: captions, credits, alt text layout' },
    @{ Key = '3'; Name = 'SLID Macro';   File = 'SLID_Macro.ps1';   Marker = 'SLID MACRO SCRIPT';       VerRx = "SLID_VERSION\s*=\s*'([^']+)'";   Info = 'Slide titles, shape names, language, title metadata' },
    @{ Key = '4'; Name = 'Alt Text Tool'; File = 'AltText_Tool.ps1'; Marker = 'ALT TEXT TOOL SCRIPT';   VerRx = "ALT_VERSION\s*=\s*'([^']+)'";    Info = 'Fills picture alt text in PowerPoint from the Excel log' },
    @{ Key = '5'; Name = 'Link Fixer';   File = 'Link_Fixer.ps1';   Marker = 'LF_VERSION';              VerRx = '(?m)^\$LF_VERSION\s*=\s*(\d+)';  Info = 'Link text in Word and PowerPoint files' }
)

function Get-Ver($tool, [string]$text) {
    $m = [regex]::Match($text, $tool.VerRx); if ($m.Success) { return $m.Groups[1].Value } else { return 'unknown' }
}
function Test-ToolText($tool, [string]$text) {
    if (-not $text -or $text.Length -lt 3000 -or $text -notmatch [regex]::Escape($tool.Marker)) { return 'the file looks incomplete' }
    $errs = $null; $tok = $null
    [void][Management.Automation.Language.Parser]::ParseInput($text, [ref]$tok, [ref]$errs)
    if ($errs -and $errs.Count -gt 0) { return 'the file has errors' }
    return ''
}
function Get-LocalPath($tool) { return (Join-Path $S4C_HOME $tool.File) }

# Fetches the newest copy of one tool. Never stops the tool from running: the copy already here is used when GitHub cannot be reached.
function Update-Tool($tool) {
    $local = Get-LocalPath $tool
    $old = if (Test-Path -LiteralPath $local) { [IO.File]::ReadAllText($local) } else { '' }
    $oldVer = if ($old) { Get-Ver $tool $old } else { '' }
    Write-Host ($tool.Name + ' - checking GitHub for the newest version (can take up to a minute)...') -ForegroundColor Cyan
    $new = Get-Remote $tool.File
    if ($null -eq $new) {
        if ($old) { Write-Host ('  Could not reach GitHub - using the copy already here (version ' + $oldVer + ').') -ForegroundColor Yellow; Write-Host ('  (' + $script:LastError + ')') -ForegroundColor DarkGray; return $true }
        Write-Host '  Could not reach GitHub and there is no copy of this tool on this computer yet.' -ForegroundColor Red
        Write-Host ('  (' + $script:LastError + ')') -ForegroundColor DarkGray
        Write-Host '  Check the network / VPN and try again. If it keeps failing, send this window to Hari.' -ForegroundColor Red
        return $false
    }
    $bad = Test-ToolText $tool $new
    if ($bad) {
        if ($old) { Write-Host ('  GitHub answered, but ' + $bad + ' - kept your version ' + $oldVer + '.') -ForegroundColor Yellow; return $true }
        Write-Host ('  GitHub answered, but ' + $bad + '. Try again in a minute.') -ForegroundColor Red; return $false
    }
    $norm = { param($t) ($t -replace "`r`n", "`n").TrimEnd() }
    $newVer = Get-Ver $tool $new
    if ((& $norm $new) -ne (& $norm $old)) {
        [IO.File]::WriteAllText($local, $new, (New-Object Text.UTF8Encoding($false)))
        if ($old) { Write-Host ('  Connected to GitHub. UPDATED: version ' + $oldVer + ' -> ' + $newVer) -ForegroundColor Green }
        else { Write-Host ('  Connected to GitHub. Downloaded version ' + $newVer + '.') -ForegroundColor Green }
    } else { Write-Host ('  Connected to GitHub. You already have the newest version (' + $newVer + ').') -ForegroundColor Green }
    return $true
}

function Start-Tool($tool) {
    if (-not (Update-Tool $tool)) { return }
    Write-Host ''
    $exe = (Get-Process -Id $PID).Path
    # the tools keep their own files (Link Fixer's cache) in their own folder; their built-in update check is switched off
    $env:LF_DIR = Join-Path $S4C_HOME 'LinkFixer'; if (-not (Test-Path -LiteralPath $env:LF_DIR)) { [void](New-Item -ItemType Directory -Path $env:LF_DIR -Force) }
    $env:LF_UPDATED = '1'; Remove-Item Env:LF_UPDATE_URL -ErrorAction SilentlyContinue
    foreach ($v in 'WF_UPDATE_URL', 'IMBK_UPDATE_URL', 'SLID_UPDATE_URL', 'ALT_UPDATE_URL') { Remove-Item "Env:$v" -ErrorAction SilentlyContinue }
    & $exe -NoProfile -ExecutionPolicy Bypass -STA -File (Get-LocalPath $tool)
}

function Show-Menu {
    Clear-Host
    Write-Host 'S4Carlisle remediation tools' -ForegroundColor Cyan
    Write-Host ('  launcher version ' + $S4C_LAUNCHER_VERSION + '   (every tool is updated from GitHub when you choose it)') -ForegroundColor DarkGray
    Write-Host ''
    foreach ($t in $TOOLS) {
        $lp = Get-LocalPath $t
        $v = if (Test-Path -LiteralPath $lp) { 'version ' + (Get-Ver $t ([IO.File]::ReadAllText($lp))) } else { 'not downloaded yet' }
        Write-Host ('  ' + $t.Key + '  ' + $t.Name.PadRight(14) + ' ' + $t.Info) -NoNewline
        Write-Host ('   [' + $v + ']') -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Host '  U  Update every tool now'
    Write-Host '  Q  Quit'
    Write-Host ''
}

if ($env:S4C_CHOICE) { $auto = @($env:S4C_CHOICE.Split(',')) } else { $auto = @() }
while ($true) {
    if ($auto.Count -eq 0 -or -not $env:S4C_CHOICE) { Show-Menu }
    if ($auto.Count -gt 0) { $c = $auto[0]; $auto = @($auto | Select-Object -Skip 1) } else { $c = Read-Host 'Type a number and press Enter' }
    $c = ([string]$c).Trim().ToUpperInvariant()
    if ($c -eq 'Q' -or $c -eq '') { if ($c -eq 'Q') { break } else { continue } }
    if ($c -eq 'U') { foreach ($t in $TOOLS) { [void](Update-Tool $t) }; Write-Host ''; [void](Read-Host 'Press Enter to go back to the menu'); continue }
    $tool = $TOOLS | Where-Object { $_.Key -eq $c } | Select-Object -First 1
    if (-not $tool) { continue }
    Start-Tool $tool
    Write-Host ''
    if ($env:S4C_CHOICE) { if ($auto.Count -eq 0) { break } } else { [void](Read-Host 'Finished. Press Enter to go back to the menu') }
}
