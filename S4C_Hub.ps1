# S4C HUB SCRIPT - the S4C Tools window. Runs a small page on this PC only (http://localhost) and opens it
# S4C-FILE-VERSION: 2.3
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
function Start-Job2([string]$file, [bool]$launch, [string[]]$files = @(), [hashtable]$envs = @{}, [int]$total = 0) {
    $id = [guid]::NewGuid().ToString('N').Substring(0, 10)
    $out = Join-Path ([IO.Path]::GetTempPath()) ("s4c_upd_$id.txt"); $err = $out + '.err'
    $args1 = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'S4C_Update.ps1') + '" -Tool ' + $file
    $sf = $out + '.status'
    $env:S4C_STATUS_FILE = $sf
    $p = Start-Process -FilePath $exe -ArgumentList $args1 -NoNewWindow -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    Remove-Item Env:S4C_STATUS_FILE -ErrorAction SilentlyContinue
    $jobs[$id] = @{ Id = $id; File = $file; Launch = $launch; State = 'updating'; Upd = $p; Out = $out; Err = $err; Sf = $sf; Status = $null; Tool = $null; Started = (Get-Date); Files = $files; Envs = $envs; Total = $(if ($total -gt 0) { $total } else { $files.Count }); Log = ($out + '.run.txt') }
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
            # files chosen on the page: hand them to the tool and run it hidden, its messages go to a log the page reads
            $fileEnv = @{ 'Word_Fixer.ps1' = 'WF_FILES'; 'SLID_Macro.ps1' = 'SLID_FILES'; 'IMBK_Macro.ps1' = 'IMBK_FILES'; 'Link_Fixer.ps1' = 'LF_FILES' }[$j.File]
            $hidden = ($j.Files -and $j.Files.Count -gt 0)
            if ($hidden) {
                if ($fileEnv) { Set-Item -Path ('Env:' + $fileEnv) -Value ($j.Files -join ';') }
                foreach ($k in $j.Envs.Keys) { Set-Item -Path ('Env:' + $k) -Value ([string]$j.Envs[$k]) }
                $env:S4C_AUTO = '1'
            }
            if ($env:S4C_HUB_TEST) { $j.Tool = Start-Process -FilePath $exe -ArgumentList ($a + ' -Test') -NoNewWindow -PassThru }
            elseif ($hidden) {
                $sp = @{ FilePath = $exe; ArgumentList = $a; RedirectStandardOutput = $j.Log; RedirectStandardError = ($j.Log + '.err'); PassThru = $true }
                if (-not ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows)) { $sp['WindowStyle'] = 'Hidden' }
                $j.Tool = Start-Process @sp
            }
            else { $j.Tool = Start-Process -FilePath $exe -ArgumentList $a -PassThru }
            Remove-Item Env:S4C_GH_STATUS, Env:S4C_GH_MSG, Env:S4C_GH_REASON, Env:WF_FILES, Env:SLID_FILES, Env:IMBK_FILES, Env:LF_FILES, Env:ALTMATCH_MODE, Env:ALTMATCH_PPTX, Env:ALTMATCH_XLSX, Env:ALTMATCH_CHAPTER, Env:ALTMATCH_MAP, Env:S4C_AUTO -ErrorAction SilentlyContinue
            $j.State = 'running'
        } else { $j.State = 'done' }
    } elseif ($j.State -eq 'running' -and $j.Tool.HasExited) { $j.State = 'done' }
}


function Get-JobProgress($j) {
    $p = [ordered]@{ total = [int]$j.Total; done = 0; current = ''; lookDone = 0; lookTotal = 0; serving = $false; saved = 0; errors = @() }
    if (-not $j.Log -or -not (Test-Path -LiteralPath $j.Log)) { return $p }
    $txt = ''
    try { $fs = New-Object IO.FileStream($j.Log, 'Open', 'Read', 'ReadWrite'); $sr = New-Object IO.StreamReader($fs); $txt = $sr.ReadToEnd(); $sr.Close() } catch { return $p }
    $fixN = 0; $errs = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($txt -split "`r?`n")) {
        $t = $line.Trim()
        if ($t -match '^Fixing:\s*(.+)$') { $fixN++; $p.current = ($Matches[1] -replace '^.*[\\/]', '') }
        elseif ($t -match '^Looking up (\d+) of (\d+)') { $p.lookDone = [int]$Matches[1]; $p.lookTotal = [int]$Matches[2] }
        elseif ($t -match '^Saved:') { $p.saved++ }
        elseif ($t -match 'The review page is open at') { $p.serving = $true }
        if ($t -match 'FAILED' -and $t.Length -lt 400) { $errs.Add($t) }
    }
    $p.done = [math]::Min([int]$j.Total, [math]::Max($p.saved, [math]::Max(0, $fixN - 1)))
    # anything the tool wrote to the error stream (for example a crash before it could print a message)
    try {
        $ep = $j.Log + '.err'
        if (Test-Path -LiteralPath $ep) {
            $fs2 = New-Object IO.FileStream($ep, 'Open', 'Read', 'ReadWrite'); $sr2 = New-Object IO.StreamReader($fs2); $et = $sr2.ReadToEnd(); $sr2.Close()
            $et = [regex]::Replace($et, "\x1B\[[0-9;]*[A-Za-z]", '')
            foreach ($el in ($et -split "`r?`n")) { $x = $el.Trim(); if ($x -and $x.Length -lt 400 -and $errs.Count -lt 3) { $errs.Add($x) } }
        }
    } catch {}
    $p.errors = @($errs.ToArray())
    return $p
}

# ---------------------------------------------------------------- the page server
# the version of this script when it was started: if GitHub later replaces the file, the page offers a restart
$hubVerRx = '# S4C-FILE-VERSION: ([0-9.]+)'
function Get-HubVer { try { $m = [regex]::Match([IO.File]::ReadAllText($PSCommandPath), $hubVerRx); if ($m.Success) { return [version]$m.Groups[1].Value } } catch {}; return $null }
$startVer = Get-HubVer
$token = [guid]::NewGuid().ToString('N')
$listener = $null; $port = 0
foreach ($try in 1..40) {
    $port = Get-Random -Minimum 49200 -Maximum 49900
    if ($env:S4C_HUB_PORT) { $port = [int]$env:S4C_HUB_PORT }
    $l = New-Object System.Net.HttpListener
    $l.Prefixes.Add("http://localhost:$port/")
    try { $l.Start(); $listener = $l; break } catch { $l.Close(); if ($env:S4C_HUB_PORT) { Start-Sleep -Milliseconds 700 } }
}
if (-not $listener) { Write-Host 'Could not start S4C Tools (no free port).' -ForegroundColor Red; exit 1 }
$url = "http://localhost:$port/"
Write-Host ''
Write-Host '  S4C Tools is open in your browser.' -ForegroundColor Green
Write-Host ('  Address: ' + $url)
Write-Host '  Keep this window open while you work. Close it to close S4C Tools.' -ForegroundColor DarkGray
Write-Host ''
if (-not $env:S4C_HUB_TEST -and -not $env:S4C_HUB_NOOPEN) { Start-Process $url }

# the page files keep themselves up to date from GitHub in the background (the page opens straight away)
$selfStatus = Join-Path ([IO.Path]::GetTempPath()) ('s4c_self_' + $token.Substring(0, 8) + '.txt')
$selfProc = $null
function Start-SelfUpdate {
    if (Test-Path -LiteralPath (Join-Path $here 'S4C_SelfUpdate.ps1')) {
        Remove-Item -LiteralPath $selfStatus -Force -ErrorAction SilentlyContinue
        $env:S4C_SELF_STATUS = $selfStatus
        try { $script:selfProc = Start-Process -FilePath $exe -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'S4C_SelfUpdate.ps1') + '"') -NoNewWindow -PassThru } catch {}
        Remove-Item Env:S4C_SELF_STATUS -ErrorAction SilentlyContinue
    }
}
Start-SelfUpdate

# the "Check for updates" button at the top of the page: the page files and every tool are checked against GitHub
$script:allProc = $null; $script:allOut = ''
function Start-UpdateAll {
    if ($script:allProc -and -not $script:allProc.HasExited) { return }
    Start-SelfUpdate
    $script:allOut = Join-Path ([IO.Path]::GetTempPath()) ('s4c_all_' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.txt')
    $upd = Join-Path $here 'S4C_Update.ps1'
    $names = ($TOOLS.Keys | ForEach-Object { "'" + $_ + "'" }) -join ','
    $cmd = "foreach (`$t in @($names)) { Write-Host ('### ' + `$t); & '$upd' -Tool `$t }"
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    $sp = @{ FilePath = $exe; ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $enc); PassThru = $true; RedirectStandardOutput = $script:allOut; RedirectStandardError = ($script:allOut + '.err') }
    if (-not ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows)) { $sp['WindowStyle'] = 'Hidden' }
    try { $script:allProc = Start-Process @sp } catch { $script:allProc = $null }
}


# ---------------------------------------------------------------- the Windows file / folder window (opened by this program so the real path is known)
$picks = @{}
$PICK_SCRIPT = @'
$ErrorActionPreference = 'Stop'
$out = $env:S4C_PICK_OUT
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -TypeDefinition 'using System; using System.Threading; using System.Runtime.InteropServices; public class S4CFront { [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h); [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int c); [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool f); [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId(); [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd); [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags); [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra); static IntPtr ownerH; public static void Raise(IntPtr h) { uint pid; uint ft = GetWindowThreadProcessId(GetForegroundWindow(), out pid); uint me = GetCurrentThreadId(); keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero); if (ft != 0 && ft != me) AttachThreadInput(me, ft, true); ShowWindow(h, 9); SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 3); BringWindowToTop(h); SetForegroundWindow(h); if (ft != 0 && ft != me) AttachThreadInput(me, ft, false); } static void Watch() { for (int i = 0; i < 40; i++) { Thread.Sleep(100); IntPtr d = GetWindow(ownerH, 6); if (d != IntPtr.Zero) { Raise(d); if (i > 5) break; } } } public static void RaiseDialogLater(IntPtr owner) { ownerH = owner; Thread t = new Thread(new ThreadStart(Watch)); t.IsBackground = true; t.Start(); } }'
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true; $owner.ShowInTaskbar = $false; $owner.FormBorderStyle = 'None'; $owner.Opacity = 0
    $owner.StartPosition = 'CenterScreen'; $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Show(); $owner.Activate(); [S4CFront]::Raise($owner.Handle)
    $start = $env:S4C_PICK_START
    $res = @()
    if ($env:S4C_PICK_MODE -eq 'folder') {
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = $env:S4C_PICK_TITLE
        if ($start -and (Test-Path -LiteralPath $start)) { $d.SelectedPath = $start }
        [S4CFront]::RaiseDialogLater($owner.Handle); if ($d.ShowDialog($owner) -eq 'OK') { $res = @($d.SelectedPath) }
    } else {
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Title = $env:S4C_PICK_TITLE; $d.Filter = $env:S4C_PICK_FILTER; $d.Multiselect = $true
        if ($start -and (Test-Path -LiteralPath $start)) { $d.InitialDirectory = $start }
        [S4CFront]::RaiseDialogLater($owner.Handle); if ($d.ShowDialog($owner) -eq 'OK') { $res = @($d.FileNames) }
    }
    $owner.Close()
    [IO.File]::WriteAllLines($out, [string[]]$res, (New-Object Text.UTF8Encoding($false)))
} catch {
    [IO.File]::WriteAllText($out + '.err', [string]$_.Exception.Message)
}
[IO.File]::WriteAllText($out + '.done', '1')
'@
function Get-LastFolderFile { $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }; $d = Join-Path $base 'S4C_Tools'; if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d) }; return (Join-Path $d 'lastfolder.txt') }
function Get-LastFolder {
    try { $f = Get-LastFolderFile; if (Test-Path -LiteralPath $f) { $v = ([IO.File]::ReadAllText($f)).Trim(); if ($v -and (Test-Path -LiteralPath $v)) { return $v } } } catch {}
    $dl = if ($env:USERPROFILE) { Join-Path $env:USERPROFILE 'Downloads' } else { '' }
    if ($dl -and (Test-Path -LiteralPath $dl)) { return $dl }
    return [Environment]::GetFolderPath('MyDocuments')
}
function Start-Pick([string]$mode, [string]$filter, [string]$title) {
    $id = [guid]::NewGuid().ToString('N').Substring(0, 10)
    $out = Join-Path ([IO.Path]::GetTempPath()) ("s4c_pick_$id.txt")
    if ($env:S4C_PICK_FAKE) {
        $val = if ($mode -eq 'folder') { $env:S4C_PICK_FAKE_DIR } else { $env:S4C_PICK_FAKE }
        [IO.File]::WriteAllLines($out, [string[]]@($val -split ';' | Where-Object { $_ }))
        [IO.File]::WriteAllText($out + '.done', '1')
        $picks[$id] = @{ Out = $out; Proc = $null; Mode = $mode }
        return $id
    }
    if (Test-WarmAlive) {
        # hand the request to the waiting helper (written to a temp name first, then renamed, so it is never read half written)
        $tmp = $script:warmReq + '.tmp'
        [IO.File]::WriteAllLines($tmp, [string[]]@($mode, $filter, $title, (Get-LastFolder), $out), (New-Object Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $tmp -Destination $script:warmReq -Force
        $picks[$id] = @{ Out = $out; Proc = $null; Mode = $mode; Warm = $true; Req = $script:warmReq; T0 = (Get-Date); Filter = $filter; Title = $title }
        return $id
    }
    $env:S4C_PICK_MODE = $mode; $env:S4C_PICK_FILTER = $filter; $env:S4C_PICK_TITLE = $title; $env:S4C_PICK_OUT = $out; $env:S4C_PICK_START = (Get-LastFolder)
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($PICK_SCRIPT))
    $sp = @{ FilePath = $exe; ArgumentList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $enc); PassThru = $true }
    if (-not ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows)) { $sp['WindowStyle'] = 'Hidden' }
    $proc = Start-Process @sp
    Remove-Item Env:S4C_PICK_MODE, Env:S4C_PICK_FILTER, Env:S4C_PICK_TITLE, Env:S4C_PICK_OUT, Env:S4C_PICK_START -ErrorAction SilentlyContinue
    $picks[$id] = @{ Out = $out; Proc = $proc; Mode = $mode }
    return $id
}

# A helper that is started once when S4C Tools opens and waits with the file window ready, so the window appears at once
# when Choose files is pressed (starting a new PowerShell for every click took 2 to 3 seconds).
$WARM_SCRIPT = @'
$ErrorActionPreference = 'Stop'
$req = $env:S4C_WARM_REQ; $hubPid = [int]$env:S4C_HUB_PID
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition 'using System; using System.Threading; using System.Runtime.InteropServices; public class S4CFront { [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h); [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int c); [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool f); [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId(); [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd); [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags); [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra); static IntPtr ownerH; public static void Raise(IntPtr h) { uint pid; uint ft = GetWindowThreadProcessId(GetForegroundWindow(), out pid); uint me = GetCurrentThreadId(); keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero); if (ft != 0 && ft != me) AttachThreadInput(me, ft, true); ShowWindow(h, 9); SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 3); BringWindowToTop(h); SetForegroundWindow(h); if (ft != 0 && ft != me) AttachThreadInput(me, ft, false); } static void Watch() { for (int i = 0; i < 40; i++) { Thread.Sleep(100); IntPtr d = GetWindow(ownerH, 6); if (d != IntPtr.Zero) { Raise(d); if (i > 5) break; } } } public static void RaiseDialogLater(IntPtr owner) { ownerH = owner; Thread t = new Thread(new ThreadStart(Watch)); t.IsBackground = true; t.Start(); } }'
[System.Windows.Forms.Application]::EnableVisualStyles()
$warmUp = New-Object System.Windows.Forms.OpenFileDialog
$warmUp.Dispose()
[IO.File]::WriteAllText($req + '.ready', '1')
$tick = 0
while ($true) {
    if (Test-Path -LiteralPath $req) {
        $lines = $null
        try { $lines = [IO.File]::ReadAllLines($req, [Text.Encoding]::UTF8); Remove-Item -LiteralPath $req -Force } catch { Start-Sleep -Milliseconds 40; continue }
        if ($lines.Count -ge 5) {
            $mode = $lines[0]; $filter = $lines[1]; $title = $lines[2]; $start = $lines[3]; $out = $lines[4]
            try {
                $owner = New-Object System.Windows.Forms.Form
                $owner.TopMost = $true; $owner.ShowInTaskbar = $false; $owner.FormBorderStyle = 'None'; $owner.Opacity = 0
                $owner.StartPosition = 'CenterScreen'; $owner.Size = New-Object System.Drawing.Size(1, 1)
                $owner.Show(); $owner.Activate(); [S4CFront]::Raise($owner.Handle)
                $res = @()
                if ($mode -eq 'folder') {
                    $d = New-Object System.Windows.Forms.FolderBrowserDialog
                    $d.Description = $title
                    if ($start -and (Test-Path -LiteralPath $start)) { $d.SelectedPath = $start }
                    [S4CFront]::RaiseDialogLater($owner.Handle); if ($d.ShowDialog($owner) -eq 'OK') { $res = @($d.SelectedPath) }
                } else {
                    $d = New-Object System.Windows.Forms.OpenFileDialog
                    $d.Title = $title; $d.Filter = $filter; $d.Multiselect = $true
                    if ($start -and (Test-Path -LiteralPath $start)) { $d.InitialDirectory = $start }
                    [S4CFront]::RaiseDialogLater($owner.Handle); if ($d.ShowDialog($owner) -eq 'OK') { $res = @($d.FileNames) }
                }
                $owner.Close(); $owner.Dispose()
                [IO.File]::WriteAllLines($out, [string[]]$res, (New-Object Text.UTF8Encoding($false)))
            } catch { try { [IO.File]::WriteAllText($out + '.err', [string]$_.Exception.Message) } catch {} }
            try { [IO.File]::WriteAllText($out + '.done', '1') } catch {}
        }
        continue
    }
    Start-Sleep -Milliseconds 50
    $tick++
    if ($tick % 40 -eq 0 -and -not (Get-Process -Id $hubPid -ErrorAction SilentlyContinue)) { break }
}
'@
$script:warm = $null; $script:warmReq = $null
function Start-Warm {
    if ($env:S4C_PICK_FAKE) { return }
    try {
        $script:warmReq = Join-Path ([IO.Path]::GetTempPath()) ('s4c_warm_' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.txt')
        $env:S4C_WARM_REQ = $script:warmReq; $env:S4C_HUB_PID = [string]$PID
        $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($WARM_SCRIPT))
        $sp = @{ FilePath = $exe; ArgumentList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $enc); PassThru = $true }
        if (-not ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows)) { $sp['WindowStyle'] = 'Hidden' }
        $script:warm = Start-Process @sp
        Remove-Item Env:S4C_WARM_REQ, Env:S4C_HUB_PID -ErrorAction SilentlyContinue
    } catch { $script:warm = $null }
}
function Test-WarmAlive { return ($script:warm -and -not $script:warm.HasExited -and $script:warmReq) }

$PICK_FILTERS = @{
    'Word_Fixer.ps1'   = 'Word documents (*.docx)|*.docx'
    'SLID_Macro.ps1'   = 'PowerPoint presentations (*.pptx)|*.pptx'
    'IMBK_Macro.ps1'   = 'PowerPoint presentations (*.pptx)|*.pptx'
    'AltText_Tool.ps1' = 'Deck and Excel files (*.pptx;*.xlsx;*.csv)|*.pptx;*.xlsx;*.csv'
    'Link_Fixer.ps1'   = 'Word and PowerPoint files (*.docx;*.pptx)|*.docx;*.pptx'
}

# ---------------------------------------------------------------- files dropped on the page: find the original, save uploads
function Read-Body($req) { $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8); try { return $sr.ReadToEnd() } finally { $sr.Close() } }
function Get-SearchRoots {
    $r = New-Object System.Collections.Generic.List[string]
    $up = $env:USERPROFILE
    foreach ($p in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'), $(if ($up) { Join-Path $up 'Downloads' }), $env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer, $(if ($up) { Join-Path $up 'Desktop' }), $(if ($up) { Join-Path $up 'Documents' }))) {
        if ($p -and (Test-Path -LiteralPath $p) -and -not ($r | Where-Object { $_ -ieq $p })) { $r.Add($p) }
    }
    return , $r.ToArray()
}
# items: name, size, mtime (ms since 1970). A file is "found" when name, size and saved time all match.
function Find-Originals($items) {
    $want = @{}; foreach ($it in $items) { $want[([string]$it.name).ToLowerInvariant()] = $it }
    $found = @{}
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($root in (Get-SearchRoots)) {
        if ($sw.Elapsed.TotalSeconds -gt 30 -or $found.Count -ge $want.Count) { break }
        try {
            foreach ($f in (Get-ChildItem -LiteralPath $root -Recurse -File -Include '*.pptx', '*.docx', '*.xlsx', '*.csv' -Depth 8 -ErrorAction SilentlyContinue)) {
                $k = $f.Name.ToLowerInvariant()
                if (-not $want.ContainsKey($k) -or $found.ContainsKey($k)) { continue }
                $it = $want[$k]
                $ms = [long]([DateTimeOffset]$f.LastWriteTimeUtc).ToUnixTimeMilliseconds()
                if ([long]$f.Length -eq [long]$it.size -and [math]::Abs($ms - [long]$it.mtime) -le 2500 -and $f.FullName.IndexOf(';') -lt 0) { $found[$k] = $f.FullName }
                if ($found.Count -ge $want.Count -or $sw.Elapsed.TotalSeconds -gt 30) { break }
            }
        } catch {}
    }
    return $found
}
function Get-OutputDir {
    $d = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'S4C Output'
    if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d) }
    return $d
}


# opens File Explorer on the file and brings that window to the front (Windows keeps a window opened by a background program behind, this lifts it)
function Show-InExplorer([string]$p) {
    $isDir = Test-Path -LiteralPath $p -PathType Container
    $dir = if ($isDir) { $p } else { Split-Path -Parent $p }
    if ($isDir) { Start-Process explorer.exe ('"' + $p + '"') } else { Start-Process explorer.exe ('/select,"' + $p + '"') }
    try {
        if (-not ('S4CWin' -as [type])) {
            Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class S4CWin { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c); [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h); [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra); }'
        }
        $sh = New-Object -ComObject Shell.Application
        $want = $dir.TrimEnd('\')
        for ($i = 0; $i -lt 16; $i++) {
            Start-Sleep -Milliseconds 250
            foreach ($w in @($sh.Windows())) {
                $fp = ''
                try { $fp = [string]$w.Document.Folder.Self.Path } catch { continue }
                if ($fp.TrimEnd('\') -ieq $want) {
                    $hw = [IntPtr][int64]$w.HWND
                    if ([S4CWin]::IsIconic($hw)) { [void][S4CWin]::ShowWindow($hw, 9) }
                    [S4CWin]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero); [S4CWin]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
                    [void][S4CWin]::SetForegroundWindow($hw)
                    return
                }
            }
        }
    } catch {}
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

Start-Warm
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
            '/api/pick' {
                $mode = if ($q['mode'] -eq 'folder') { 'folder' } else { 'file' }
                $tool = [string]$q['tool']
                $flt = if ($PICK_FILTERS.ContainsKey($tool)) { $PICK_FILTERS[$tool] } else { 'Word and PowerPoint files (*.docx;*.pptx)|*.docx;*.pptx' }
                $ttl = if ($mode -eq 'folder') { 'Choose the folder this file is in' } else { 'Choose the file(s) to fix' }
                $id = Start-Pick $mode $flt $ttl
                Send $res 200 'application/json' (ToJson @{ id = $id }) $null
            }
            '/api/pickresult' {
                $pk = $picks[$q['id']]
                if (-not $pk) { Send $res 404 'text/plain' 'no such window' $null; break }
                $done = Test-Path -LiteralPath ($pk.Out + '.done'); $err = ''
                if (-not $done -and $pk.Proc -and $pk.Proc.HasExited) { $done = $true; $err = 'The file window closed unexpectedly.' }
                if (-not $done -and $pk.Warm) {
                    # the waiting helper did not pick the request up (not started, or stuck): open the window the slower way instead
                    if (-not (Test-WarmAlive)) { $done = $true; $err = 'The file window could not start.'; Start-Warm }
                    elseif ((Test-Path -LiteralPath $pk.Req) -and ((Get-Date) - $pk.T0).TotalSeconds -gt 8) {
                        Remove-Item -LiteralPath $pk.Req -Force -ErrorAction SilentlyContinue
                        $env:S4C_PICK_MODE = $pk.Mode; $env:S4C_PICK_FILTER = $pk.Filter; $env:S4C_PICK_TITLE = $pk.Title; $env:S4C_PICK_OUT = $pk.Out; $env:S4C_PICK_START = (Get-LastFolder)
                        $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($PICK_SCRIPT))
                        $sp = @{ FilePath = $exe; ArgumentList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $enc); PassThru = $true }
                        if (-not ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows)) { $sp['WindowStyle'] = 'Hidden' }
                        $pk.Proc = Start-Process @sp; $pk.Warm = $false
                        Remove-Item Env:S4C_PICK_MODE, Env:S4C_PICK_FILTER, Env:S4C_PICK_TITLE, Env:S4C_PICK_OUT, Env:S4C_PICK_START -ErrorAction SilentlyContinue
                    }
                }
                $files = @()
                if ($done) {
                    if (Test-Path -LiteralPath $pk.Out) { $files = @([IO.File]::ReadAllLines($pk.Out, [Text.Encoding]::UTF8) | Where-Object { $_ }) }
                    if (Test-Path -LiteralPath ($pk.Out + '.err')) { $err = [IO.File]::ReadAllText($pk.Out + '.err') }
                    if ($files.Count -gt 0) {
                        $lf = if ($pk.Mode -eq 'folder') { $files[0] } else { Split-Path -Parent $files[0] }
                        try { [IO.File]::WriteAllText((Get-LastFolderFile), $lf) } catch {}
                    }
                }
                Send $res 200 'application/json; charset=utf-8' (ToJson @{ done = $done; files = $files; error = $err }) $null
            }
            '/api/locate' {
                $items = @((Read-Body $req) | ConvertFrom-Json)
                $found = Find-Originals $items
                $out = @(); foreach ($it in $items) { $k = ([string]$it.name).ToLowerInvariant(); $out += [ordered]@{ name = $it.name; path = $(if ($found.ContainsKey($k)) { $found[$k] } else { '' }) } }
                Send $res 200 'application/json; charset=utf-8' (ToJson @{ files = $out }) $null
            }
            '/api/upload' {
                # the file as dropped on the page, written into the folder the person chose (an identical file already there is left alone)
                $leaf = [IO.Path]::GetFileName([string]$q['name']) -replace '[;\\/:*?"<>|]', '_'
                if ($leaf -notmatch '(?i)\.(pptx|docx|xlsx|csv)$') { Send $res 400 'text/plain' 'not a file type these tools use' $null; break }
                $dir = [string]$q['dir']
                if ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) { } else { $dir = Get-OutputDir }
                $dest = Join-Path $dir $leaf
                if ((Test-Path -LiteralPath $dest -PathType Leaf) -and ([long](Get-Item -LiteralPath $dest).Length -eq [long]$req.ContentLength64)) {
                    $null = $req.InputStream.CopyTo([IO.Stream]::Null)
                } else {
                    $fs = [IO.File]::Create($dest); try { $req.InputStream.CopyTo($fs) } finally { $fs.Close() }
                }
                Send $res 200 'application/json; charset=utf-8' (ToJson @{ path = $dest }) $null
            }
            '/api/runfiles' {
                $b = (Read-Body $req) | ConvertFrom-Json
                $f = [string]$b.tool
                $allow = @{ 'Word_Fixer.ps1' = @('.docx'); 'SLID_Macro.ps1' = @('.pptx'); 'IMBK_Macro.ps1' = @('.pptx'); 'AltText_Tool.ps1' = @('.pptx', '.xlsx', '.csv'); 'Link_Fixer.ps1' = @('.docx', '.pptx') }
                if (-not $TOOLS.Contains($f) -or -not $allow.ContainsKey($f)) { Send $res 400 'text/plain' 'this tool does not take chosen files' $null; break }
                $ok = @($b.files | Where-Object { $_ -and ($_.IndexOf(';') -lt 0) -and (Test-Path -LiteralPath $_ -PathType Leaf) -and ($allow[$f] -contains [IO.Path]::GetExtension($_).ToLowerInvariant()) })
                if ($ok.Count -eq 0) { Send $res 400 'text/plain' 'no usable files' $null; break }
                $envs = @{}; $jobFiles = $ok
                if ($f -eq 'AltText_Tool.ps1') {
                    $pp = @($ok | Where-Object { $_.ToLowerInvariant().EndsWith('.pptx') }); $xx = @($ok | Where-Object { $_.ToLowerInvariant().EndsWith('.xlsx') }); $cc = @($ok | Where-Object { $_.ToLowerInvariant().EndsWith('.csv') })
                    if ($pp.Count -eq 0 -or $xx.Count -eq 0) { Send $res 400 'text/plain' 'Alt Text Tool needs the deck (.pptx) and the Excel file (.xlsx)' $null; break }
                    $envs['ALTMATCH_PPTX'] = $pp[0]; $envs['ALTMATCH_XLSX'] = $xx[0]
                    if ($cc.Count -gt 0) { $envs['ALTMATCH_MODE'] = '2'; $envs['ALTMATCH_MAP'] = $cc[0] } else { $envs['ALTMATCH_MODE'] = '1'; $envs['ALTMATCH_CHAPTER'] = [string]$b.chapter }
                    $jobFiles = @($pp[0])
                }
                $id = Start-Job2 $f $true $jobFiles $envs $(if ($f -eq 'AltText_Tool.ps1') { 1 } else { $jobFiles.Count })
                Send $res 200 'application/json' (ToJson @{ id = $id }) $null
            }
            '/api/stop' {
                # ends a tool that is waiting for the person to finish in its review page
                $j = $jobs[$q['id']]
                if ($j -and $j.Tool -and -not $j.Tool.HasExited) { try { & taskkill.exe /PID $j.Tool.Id /T /F 2>&1 | Out-Null } catch { try { Stop-Process -Id $j.Tool.Id -Force } catch {} } }
                Send $res 200 'application/json' '{}' $null
            }
            '/api/openfolder' {
                $p = [string]$q['path']
                if ($p -and (Test-Path -LiteralPath $p)) { Show-InExplorer $p }
                Send $res 200 'application/json' '{}' $null
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
                if ($j.Files -and $j.Files.Count -gt 0) { $body['progress'] = (Get-JobProgress $j) }
                Send $res 200 'application/json' (ToJson $body) $null
            }
            '/api/delete' {
                # removes ONE chapter (all its tool runs) from this person's work log for the day; the removed lines are kept so it can be undone
                $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
                $dir = Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog'; $log = Join-Path $dir 'worklog.tsv'; $trash = Join-Path $dir 'worklog.deleted.tsv'
                $day = $q['day']; $key = ([string]$q['key']).ToLowerInvariant(); $removed = 0
                if ($day -and $key -and (Test-Path -LiteralPath $log)) {
                    $keep = New-Object System.Collections.Generic.List[string]; $gone = New-Object System.Collections.Generic.List[string]
                    foreach ($l in [IO.File]::ReadAllLines($log, [Text.Encoding]::UTF8)) {
                        $c = $l.Split("`t"); $hit = $false
                        if ($c.Count -ge 8 -and $c[0] -eq $day) {
                            $leaf = [IO.Path]::GetFileName($c[4])
                            $k = (([IO.Path]::GetFileNameWithoutExtension($leaf) -replace '(?i)(\s+(Remediated|AltText|Links|Fixed)|_BEFORE_IMBKfix\d*)+$', '') + [IO.Path]::GetExtension($leaf)).ToLowerInvariant()
                            if ($k -eq $key) { $hit = $true }
                        }
                        if ($hit) { $gone.Add($l) } else { $keep.Add($l) }
                    }
                    if ($gone.Count) {
                        Copy-Item -LiteralPath $log -Destination (Join-Path $dir 'worklog.backup.tsv') -Force
                        [IO.File]::WriteAllLines($trash, $gone.ToArray(), (New-Object Text.UTF8Encoding($false)))
                        [IO.File]::WriteAllLines($log, $keep.ToArray(), (New-Object Text.UTF8Encoding($false)))
                        $removed = $gone.Count
                    }
                }
                Send $res 200 'application/json; charset=utf-8' (ToJson ([ordered]@{ removed = $removed })) $null
            }
            '/api/undelete' {
                $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
                $dir = Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog'; $log = Join-Path $dir 'worklog.tsv'; $trash = Join-Path $dir 'worklog.deleted.tsv'; $back = 0
                if (Test-Path -LiteralPath $trash) {
                    $lines = [IO.File]::ReadAllLines($trash, [Text.Encoding]::UTF8)
                    if ($lines.Count) { [IO.File]::AppendAllLines($log, $lines, (New-Object Text.UTF8Encoding($false))); $back = $lines.Count }
                    Remove-Item -LiteralPath $trash -Force -ErrorAction SilentlyContinue
                }
                Send $res 200 'application/json; charset=utf-8' (ToJson ([ordered]@{ restored = $back })) $null
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
            '/api/updateall' {
                Start-UpdateAll
                Send $res 200 'application/json' '{}' $null
            }
            '/api/updateallstatus' {
                $done = (-not $script:allProc) -or $script:allProc.HasExited
                $list = @()
                if ($done -and $script:allOut -and (Test-Path -LiteralPath $script:allOut)) {
                    $txt = [IO.File]::ReadAllText($script:allOut)
                    foreach ($blk in ($txt -split '(?m)^### ')) {
                        if (-not $blk.Trim()) { continue }
                        $nl = $blk.IndexOf("`n"); if ($nl -lt 0) { continue }
                        $file = $blk.Substring(0, $nl).Trim()
                        if (-not $TOOLS.Contains($file)) { continue }
                        $st = Read-Status $blk.Substring($nl + 1)
                        $list += [ordered]@{ file = $file; name = $TOOLS[$file].Name; status = $st.status; message = $st.message }
                    }
                }
                Send $res 200 'application/json' (ToJson ([ordered]@{ done = $done; tools = $list })) $null
            }
            '/api/hub' {
                $done = (-not $selfProc) -or $selfProc.HasExited
                $st = ''; $upd = @()
                if ($done -and (Test-Path -LiteralPath $selfStatus)) { $lines = [IO.File]::ReadAllText($selfStatus).Split("`n") | Where-Object { $_ }; if ($lines) { $st = $lines[0]; $upd = @($lines | Select-Object -Skip 1) } }
                $hv = [regex]::Match([IO.File]::ReadAllText((Join-Path $here 'S4C_Hub.html')), 'name="s4c-hub-version" content="([0-9.]+)"').Groups[1].Value
                $nowVer = Get-HubVer; $need = [bool]($startVer -and $nowVer -and $nowVer -gt $startVer)
                Send $res 200 'application/json' (ToJson ([ordered]@{ done = $done; status = $st; updated = $upd; version = $hv; restart = $need })) $null
            }
            '/api/restart' {
                # starts a fresh copy of this script on the same address, then closes this one (used after GitHub updated this script)
                Send $res 200 'application/json' (ToJson ([ordered]@{ ok = $true })) $null
                try {
                    $env:S4C_HUB_PORT = [string]$port; $env:S4C_HUB_NOOPEN = '1'
                    Start-Process -FilePath $exe -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"')
                } catch {}
                $quit = $true
            }
            '/api/quit' { Send $res 200 'text/plain' 'bye' $null; $quit = $true }
            default { Send $res 404 'text/plain' 'not found' $null }
        }
    } catch {
        try { Send $res 500 'text/plain' $_.Exception.Message $null } catch {}
    }
}
$listener.Stop()
try { if ($script:warm -and -not $script:warm.HasExited) { $script:warm.Kill() } } catch {}
