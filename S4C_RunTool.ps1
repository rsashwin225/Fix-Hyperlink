# S4C RUNTOOL SCRIPT - opens one tool in its own window, the same way the menu did, and waits for Enter at the end.
# S4C-FILE-VERSION: 1.8
param([string]$Tool, [switch]$Test)
$sta = @{ 'Word_Fixer.ps1' = $true; 'SLID_Macro.ps1' = $true; 'IMBK_Macro.ps1' = $false; 'AltText_Tool.ps1' = $true; 'Link_Fixer.ps1' = $true; 'Issue_Fixer.ps1' = $false }
if (-not $sta.ContainsKey($Tool)) { Write-Host ('Unknown tool: ' + $Tool); exit 1 }
try { $host.UI.RawUI.WindowTitle = 'S4C Tools - ' + $Tool.Replace('.ps1', '').Replace('_', ' ') } catch {}
# GitHub check result from S4C Tools (the check is done before this window opens)
if ($env:S4C_GH_STATUS) {
    $name = $Tool.Replace('.ps1', '').Replace('_', ' ')
    Write-Host ($name + ' - GitHub check (done by S4C Tools before opening this window):') -ForegroundColor Cyan
    switch ($env:S4C_GH_STATUS) {
        'latest'  { Write-Host ('  Connected to GitHub. ' + $env:S4C_GH_MSG + '.') -ForegroundColor Green }
        'updated' { Write-Host ('  Connected to GitHub. ' + $env:S4C_GH_MSG + '.') -ForegroundColor Green }
        'newer'   { Write-Host ('  Connected to GitHub. ' + $env:S4C_GH_MSG + '.') -ForegroundColor Green }
        'offline' { Write-Host '  NOT connected to GitHub - using the copy already on this PC.' -ForegroundColor Yellow; if ($env:S4C_GH_REASON) { Write-Host ('  (' + $env:S4C_GH_REASON + ')') -ForegroundColor DarkGray } }
        default   { Write-Host ('  ' + $env:S4C_GH_MSG) -ForegroundColor Yellow }
    }
    Write-Host ''
    Remove-Item Env:S4C_GH_STATUS, Env:S4C_GH_MSG, Env:S4C_GH_REASON -ErrorAction SilentlyContinue
}
$env:LF_DIR = $PSScriptRoot
$env:LF_UPDATED = '1'   # Link Fixer: S4C Tools has already done the GitHub check
$exe = (Get-Process -Id $PID).Path
$a = @('-NoProfile', '-ExecutionPolicy', 'Bypass')
if ($sta[$Tool] -and -not $Test) { $a += '-STA' }
$a += @('-File', (Join-Path $PSScriptRoot $Tool))
if ($Test) { Write-Host ('[test] would run: ' + ($a -join ' ')); exit 0 }
$igSince = Get-Date
& $exe @a
Write-Host ''
# text integrity check: the original and the fixed file are compared and an Excel report is saved next to the fixed file (never stops the tool)
try {
    $igFile = Join-Path $PSScriptRoot 'S4C_Integrity.ps1'
    if (Test-Path -LiteralPath $igFile) { . $igFile; Invoke-IntegrityRun $Tool $igSince }
} catch { Write-Host ('The text integrity check could not run: ' + $_.Exception.Message) -ForegroundColor Yellow }
Write-Host ''
if ($env:S4C_AUTO) { } else { [void](Read-Host 'Finished. Press Enter to close this window') }
