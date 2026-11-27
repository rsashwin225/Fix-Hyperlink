# Link_Fixer.ps1
# Finds links in Word and PowerPoint files whose visible text is a raw web address or a vague phrase
# (click here, link, read more ...) and replaces that visible text with meaningful words.
#
#   1. Pick one or more .docx or .pptx files.
#   2. The tool first checks that each web address works. For a working address it looks up the page
#      (page title, or when the title says little the page heading or description, or the title stored in a
#      PDF; for a YouTube link "Watch the YouTube video of <title>"). When nothing meaningful is found, or the
#      link is broken or blocked, the box stays empty and the review page shows a Go to Link button. Go to Link
#      opens the page in its own Edge or Chrome window and, once the page has loaded, fills the box with
#      the text it reads there. You edit the text and press Apply.
#   3. A copy of each file is saved (<name> Links.docx or .pptx). The original is never changed.
#
# Only the visible text of a link changes. The link address, the link formatting and the language stay.
# Text is written exactly as shown in the box. For PDF links the text ends with .pdf.
# Only the web address is sent over the internet for the lookup, never your document.
# The page opened by Go to Link is read on this computer only.

# ---------- version and automatic update ----------
# LF_VERSION goes up by one with every new release. When the launcher (Link Fixer.cmd) has an update address
# (LF_UPDATE_URL), the newest Link_Fixer.ps1 is looked up there each time the tool starts. A newer one is checked
# (it must be a Link Fixer script of a higher version that PowerShell can read) and then used for this run, and
# a copy is kept for when there is no connection. Without an address, or without a connection, the version inside
# this file is used. Nothing is sent to the address, the tool only reads from it.
$LF_VERSION = 15
# ---- work log: notes each saved file (name, slides or pages) for the daily report in S4C Tools ----
# Only reads the saved file; never changes it and never stops the tool if something goes wrong.
function Add-WorkLog([string]$tool, [string]$inPath, [string]$outPath) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $count = ''; $unit = ''
        $ext = [IO.Path]::GetExtension($outPath).ToLowerInvariant()
        $z = [IO.Compression.ZipFile]::OpenRead($outPath)
        try {
            if ($ext -eq '.pptx') {
                $unit = 'slides'
                $e = $z.GetEntry('ppt/presentation.xml')
                if ($e) { $sr = New-Object IO.StreamReader($e.Open()); $t = $sr.ReadToEnd(); $sr.Close(); $count = [regex]::Matches($t, '<p:sldId\s').Count }
            } elseif ($ext -eq '.docx') {
                $unit = 'pages'
                $e = $z.GetEntry('docProps/app.xml')
                if ($e) { $sr = New-Object IO.StreamReader($e.Open()); $t = $sr.ReadToEnd(); $sr.Close(); $m = [regex]::Match($t, '<Pages>(\d+)</Pages>'); if ($m.Success) { $count = $m.Groups[1].Value } }
            }
        } finally { $z.Dispose() }
        $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
        $dir = Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog'
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $ci = [Globalization.CultureInfo]::InvariantCulture; $now = Get-Date
        $line = @($now.ToString('yyyy-MM-dd', $ci), $now.ToString('HH:mm', $ci), $env:USERNAME, $tool, [IO.Path]::GetFullPath($inPath), [IO.Path]::GetFullPath($outPath), $count, $unit) -join "`t"
        [IO.File]::AppendAllText((Join-Path $dir 'worklog.tsv'), $line + "`r`n", (New-Object Text.UTF8Encoding($false)))
    } catch { }
}
function Get-LfVersion([string]$text) {
    $m = [regex]::Match($text, '(?m)^\$LF_VERSION\s*=\s*(\d+)\s*$')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return 0
}
function Test-LfScript([string]$text, [int]$above) {
    if (-not $text -or $text.Length -lt 20000 -or $text -notmatch '(?m)^# Link_Fixer\.ps1') { return $false }
    if ((Get-LfVersion $text) -le $above) { return $false }
    $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$errs)
    return (@($errs).Count -eq 0)
}
if (-not $env:LF_UPDATED) {
    $env:LF_UPDATED = '1'
    $lfNew = $null; $lfFrom = ''
    $lfUrl = [string]$env:LF_UPDATE_URL
    $lfReached = $false
    if ($lfUrl) { Write-Host 'Link Fixer - checking GitHub for updates...' -ForegroundColor Cyan }
    if ($lfUrl -and ($lfUrl -match '^https://' -or $env:LF_UPDATE_ALLOW_HTTP)) {
        try {
            $lfRq = [Net.HttpWebRequest]::Create($lfUrl)
            $lfRq.Timeout = 6000; $lfRq.ReadWriteTimeout = 6000; $lfRq.UserAgent = 'LinkFixer'
            try { $lfRq.Proxy = [Net.WebRequest]::GetSystemWebProxy(); $lfRq.Proxy.Credentials = [Net.CredentialCache]::DefaultCredentials } catch { }
            $lfRs = $lfRq.GetResponse(); $lfSr = New-Object IO.StreamReader($lfRs.GetResponseStream(), [Text.Encoding]::UTF8)
            $lfText = $lfSr.ReadToEnd(); $lfSr.Close(); $lfRs.Close()
            $lfReached = $true
            if (Test-LfScript $lfText $LF_VERSION) { $lfNew = $lfText; $lfFrom = 'the update address' }
        } catch { }
    }
    $lfCache = $null
    try { $lfBase = $env:LOCALAPPDATA; if (-not $lfBase) { $lfBase = [IO.Path]::GetTempPath() }; $lfCache = Join-Path (Join-Path $lfBase 'LinkFixer') 'Link_Fixer.ps1' } catch { }
    if ($lfNew) {
        try { $lfDir = Split-Path -Parent $lfCache; if (-not (Test-Path -LiteralPath $lfDir)) { $null = New-Item -ItemType Directory -Path $lfDir -Force }; [IO.File]::WriteAllText($lfCache, $lfNew, (New-Object Text.UTF8Encoding($false))) } catch { }
    } elseif ($lfCache -and (Test-Path -LiteralPath $lfCache)) {
        try { $lfC = [IO.File]::ReadAllText($lfCache, [Text.Encoding]::UTF8); if (Test-LfScript $lfC $LF_VERSION) { $lfNew = $lfC; $lfFrom = 'the copy kept on this computer' } } catch { }
    }
    if ($lfNew) {
        if ($lfFrom -eq 'the update address') { Write-Host ('  Connected to GitHub. UPDATED: version ' + $LF_VERSION + ' -> ' + (Get-LfVersion $lfNew)) -ForegroundColor Green }
        else { Write-Host ('  Could not reach GitHub - using the newer copy saved on this computer (version ' + (Get-LfVersion $lfNew) + ').') -ForegroundColor Yellow }
        Write-Host ''
        $lfUpdatedRun = $true
        Invoke-Expression $lfNew
        return
    }
}
if ($lfUrl -and -not $lfUpdatedRun) {
    if ($lfReached) { Write-Host ('  Connected to GitHub. You already have the latest version (' + $LF_VERSION + ').') -ForegroundColor Green }
    else { Write-Host '  Could not reach GitHub - using the copy already here.' -ForegroundColor Yellow }
    Write-Host ''
}
Write-Host ('Link Fixer - running version ' + $LF_VERSION)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls13 } catch { }

$NS_W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$NS_P   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
$XMLNS  = 'http://www.w3.org/XML/1998/namespace'

$ToolDir = $null
if ($env:LF_DIR) { $ToolDir = $env:LF_DIR.TrimEnd('\') }
elseif ($PSScriptRoot) { $ToolDir = $PSScriptRoot }
else { $ToolDir = Join-Path $env:USERPROFILE 'Documents' }
$ListPath = Join-Path $ToolDir 'LinkTexts.csv'

# Folder of a file, never empty: when a file has no folder part (or Windows gives nothing back), fall back to the
# folder of the tool and then to the temp folder, so saving the review page / the fixed copy can not fail on it.
function Get-FolderOf([string]$file) {
    $d = $null
    try { $d = [IO.Path]::GetDirectoryName($file) } catch { }
    if (-not $d) { try { $d = Split-Path -Parent $file } catch { } }
    if (-not $d) { $d = $ToolDir }
    if (-not $d) { $d = [IO.Path]::GetTempPath() }
    return $d
}

$vague = @('click here','here','link','this link','read more','more','this page','website','web site','url','download','see here','learn more','click','click this link','this website','visit','visit website','go here','link here')

function Pick-Files {
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Select the Word or PowerPoint file(s)'
    $dlg.Filter = 'Word and PowerPoint (*.docx;*.pptx)|*.docx;*.pptx'
    $dlg.Multiselect = $true
    $dlg.InitialDirectory = Join-Path $env:USERPROFILE 'Downloads'
    if ($dlg.ShowDialog($owner) -eq 'OK') { return @($dlg.FileNames) }
    return @()
}

function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}
function Write-ZipXml($zip, [string]$name, $xml) {
    $e = $zip.GetEntry($name); if ($e) { $e.Delete() }
    $ne = $zip.CreateEntry($name)
    $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
    $xml.Save($sw); $sw.Close()
}
function Load-Xml($text) {
    $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.LoadXml($text); return $d
}
function New-Ns($doc) {
    $m = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    $m.AddNamespace('w', $NS_W); $m.AddNamespace('a', $NS_A); $m.AddNamespace('p', $NS_P); $m.AddNamespace('r', $NS_R); $m.AddNamespace('pr', $NS_REL)
    return ,$m
}
function Resolve-Part([string]$fromPart, [string]$target) {
    if ($target.StartsWith('/')) { return $target.TrimStart('/') }
    $baseDir = [IO.Path]::GetDirectoryName($fromPart.Replace('/', '\')).Replace('\', '/')
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($s in ($baseDir + '/' + $target).Split('/')) {
        if ($s -eq '' -or $s -eq '.') { continue }
        if ($s -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) } } else { $parts.Add($s) }
    }
    return ($parts -join '/')
}
# external relationships of a part: Id -> address
function Get-ExternalRels($zip, [string]$part) {
    $relName = [IO.Path]::GetDirectoryName($part.Replace('/', '\')).Replace('\', '/') + '/_rels/' + [IO.Path]::GetFileName($part) + '.rels'
    $map = @{}
    $t = Read-ZipText $zip $relName
    if (-not $t) { return $map }
    $d = Load-Xml $t; $m = New-Ns $d
    foreach ($r in $d.SelectNodes('//pr:Relationship', $m)) {
        if ($r.GetAttribute('TargetMode') -eq 'External') { $map[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
    }
    return $map
}

function Test-WebUrl([string]$u) { return ($u -match '^(?i)(https?|ftp)://') }

# Word sometimes stores a link's target without a #fragment or ?query that
# was actually typed into the link's own visible text - two links can then
# look identical at the relationship level while actually pointing at two
# different parts of the same page. Use the fuller one (the text) as the
# real address whenever the text is clearly that same address plus more.
function Get-EffectiveUrl([string]$relUrl, [string]$visibleText) {
    $vt = ($visibleText + '').Trim()
    if ($vt.Length -gt $relUrl.Length -and $vt.StartsWith($relUrl, [StringComparison]::OrdinalIgnoreCase)) {
        $suffix = $vt.Substring($relUrl.Length)
        if ($suffix -match '^[#?]\S*$') { return $relUrl + $suffix }
    }
    return $relUrl
}

# does the visible text need to be replaced?
function Test-NeedsFix([string]$text, [string]$url) {
    $t = ($text + '').Trim()
    if (-not $t) { return $false }
    if ($t -match '^(?i)(https?://|ftp://|www\.)') { return $true }
    if ($t -ieq $url) { return $true }
    if ($t -match '^(?i)[a-z0-9.-]+\.(com|org|net|edu|gov|io|us|uk|pdf)(/\S*)?$') { return $true }
    $n = (($t.ToLowerInvariant() -replace '[^a-z ]', '') -replace '\s+', ' ').Trim()
    if ($vague -contains $n) { return $true }
    return $false
}

# ---------- find links ----------
# each hit: Part, Where, Url, Text, Fix (needs replacing), Node(s) for the apply step
function Get-DocxParts($zip) {
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($e in $zip.Entries) {
        if ($e.FullName -match '^word/(document|footnotes|endnotes|header\d+|footer\d+)\.xml$') { $names.Add($e.FullName) }
    }
    return ,$names.ToArray()
}

function Scan-Docx($zip, [bool]$apply, $texts, [int]$fi = 0) {
    $hits = New-Object System.Collections.Generic.List[object]
    $changed = 0; $skipped = 0; $other = 0; $ord = 0
    foreach ($part in (Get-DocxParts $zip)) {
        $doc = Load-Xml (Read-ZipText $zip $part); $m = New-Ns $doc
        $rels = Get-ExternalRels $zip $part
        $pidx = New-Object 'System.Collections.Generic.Dictionary[object,int]'
        $n = 0; foreach ($p in $doc.SelectNodes('//w:p', $m)) { $n++; $pidx[$p] = $n }
        $partChanged = $false
        foreach ($hl in @($doc.SelectNodes('//w:hyperlink', $m))) {
            $rid = $hl.GetAttribute('id', $NS_R)
            # a link that jumps inside the document (bookmark), a mail link or anything that is not a web address is counted but never changed
            if (-not $rid -or -not $rels.ContainsKey($rid) -or -not (Test-WebUrl $rels[$rid])) { if (-not $apply) { $other++ }; continue }
            $url = $rels[$rid]
            # every web link has its own number in its file, the same number is found again when the changes are written
            $ord++; $key = [string]$fi + ':' + [string]$ord
            $vis = (($hl.SelectNodes('.//w:t', $m) | ForEach-Object { $_.InnerText }) -join '')
            $url = Get-EffectiveUrl $url $vis
            $fix = Test-NeedsFix $vis $url
            $where = 'a header or footer'
            if ($part -like '*document.xml') {
                $anc = $hl.ParentNode; while ($anc -and $anc.LocalName -ne 'p') { $anc = $anc.ParentNode }
                if ($anc -and $pidx.ContainsKey($anc)) { $where = 'paragraph ' + $pidx[$anc] } else { $where = 'the body' }
            } elseif ($part -like '*footnotes.xml') { $where = 'a footnote' } elseif ($part -like '*endnotes.xml') { $where = 'an endnote' }
            if (-not $apply) { $hits.Add([pscustomobject]@{ Where = $where; Url = $url; Text = $vis.Trim(); Fix = $fix; Key = $key }); continue }
            if (-not $fix) { continue }
            if (-not $texts.ContainsKey($key) -or -not $texts[$key]) { $skipped++; continue }
            $runs = @($hl.SelectNodes('w:r', $m))
            if ($runs.Count -eq 0) { $skipped++; continue }
            $first = $runs[0]
            $ts = @($first.SelectNodes('w:t', $m))
            if ($ts.Count -eq 0) { $tn = $doc.CreateElement('w', 't', $NS_W); $null = $first.AppendChild($tn); $ts = @($tn) }
            $ts[0].InnerText = [string]$texts[$key]
            $a = $doc.CreateAttribute('xml', 'space', $XMLNS); $a.Value = 'preserve'; $null = $ts[0].SetAttributeNode($a)
            for ($k = 1; $k -lt $ts.Count; $k++) { $null = $ts[$k].ParentNode.RemoveChild($ts[$k]) }
            for ($k = 1; $k -lt $runs.Count; $k++) { $null = $hl.RemoveChild($runs[$k]) }
            $changed++; $partChanged = $true
        }
        if ($apply -and $partChanged) { Write-ZipXml $zip $part $doc }
    }
    return [pscustomobject]@{ Hits = $hits.ToArray(); Changed = $changed; Skipped = $skipped; Other = $other }
}

function Get-SlideNames($zip) {
    $pres = Load-Xml (Read-ZipText $zip 'ppt/presentation.xml'); $pm = New-Ns $pres
    $rels = Load-Xml (Read-ZipText $zip 'ppt/_rels/presentation.xml.rels'); $rm = New-Ns $rels
    $map = @{}
    foreach ($r in $rels.SelectNodes('//pr:Relationship', $rm)) { $map[$r.GetAttribute('Id')] = (Resolve-Part 'ppt/presentation.xml' $r.GetAttribute('Target')) }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($s in $pres.SelectNodes('//p:sldIdLst/p:sldId', $pm)) { $list.Add($map[$s.GetAttribute('id', $NS_R)]) }
    return ,$list.ToArray()
}

function Scan-Pptx($zip, [bool]$apply, $texts, [int]$fi = 0) {
    $hits = New-Object System.Collections.Generic.List[object]
    $changed = 0; $skipped = 0; $other = 0; $ord = 0
    $slides = Get-SlideNames $zip
    for ($i = 0; $i -lt $slides.Count; $i++) {
        $part = $slides[$i]
        $doc = Load-Xml (Read-ZipText $zip $part); $m = New-Ns $doc
        $rels = Get-ExternalRels $zip $part
        $partChanged = $false
        foreach ($para in $doc.SelectNodes('//a:p', $m)) {
            $groups = New-Object System.Collections.Generic.List[object]
            $cur = $null; $oth = ''
            foreach ($c in @($para.ChildNodes)) {
                $rid = ''
                if ($c.LocalName -eq 'r') {
                    $hc = $c.SelectSingleNode('a:rPr/a:hlinkClick', $m)
                    if ($hc) { $rid = $hc.GetAttribute('id', $NS_R) }
                }
                if ($rid -and $rels.ContainsKey($rid)) {
                    if ($cur -and $cur.Rid -eq $rid) { $cur.Runs.Add($c) }
                    else { $cur = [pscustomobject]@{ Rid = $rid; Runs = (New-Object System.Collections.Generic.List[object]) }; $cur.Runs.Add($c); $groups.Add($cur) }
                } else {
                    # a link to another slide or an action: counted, never changed
                    if ($rid -and -not $apply -and -not ($oth -and $oth -eq $rid)) { $other++ }
                    $oth = $rid; $cur = $null
                }
            }
            foreach ($g in $groups) {
                $url = $rels[$g.Rid]
                if (-not (Test-WebUrl $url)) { if (-not $apply) { $other++ }; continue }
                $ord++; $key = [string]$fi + ':' + [string]$ord
                $vis = (($g.Runs | ForEach-Object { $_.SelectSingleNode('a:t', $m).InnerText }) -join '')
                $url = Get-EffectiveUrl $url $vis
                $fix = Test-NeedsFix $vis $url
                if (-not $apply) { $hits.Add([pscustomobject]@{ Where = 'slide ' + ($i + 1); Url = $url; Text = $vis.Trim(); Fix = $fix; Key = $key }); continue }
                if (-not $fix) { continue }
                if (-not $texts.ContainsKey($key) -or -not $texts[$key]) { $skipped++; continue }
                $g.Runs[0].SelectSingleNode('a:t', $m).InnerText = [string]$texts[$key]
                for ($k = 1; $k -lt $g.Runs.Count; $k++) { $null = $para.RemoveChild($g.Runs[$k]) }
                $changed++; $partChanged = $true
            }
        }
        if ($apply -and $partChanged) { Write-ZipXml $zip $part $doc }
    }
    return [pscustomobject]@{ Hits = $hits.ToArray(); Changed = $changed; Skipped = $skipped; Other = $other }
}

function Scan-File([string]$path, [bool]$apply, $texts, [string]$outPath, [int]$fi = 0) {
    $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
    if ($apply) { Copy-Item -LiteralPath $path -Destination $outPath -Force; $zip = [IO.Compression.ZipFile]::Open($outPath, 'Update') }
    else { $zip = [IO.Compression.ZipFile]::OpenRead($path) }
    try {
        if ($ext -eq '.docx') { return (Scan-Docx $zip $apply $texts $fi) }
        elseif ($ext -eq '.pptx') { return (Scan-Pptx $zip $apply $texts $fi) }
        else { throw "Not a .docx or .pptx file: $path" }
    } finally { $zip.Dispose() }
}

# ---------- lookup of a meaningful text for an address ----------
function Clean-Title([string]$t) {
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = ($t -replace '\s+', ' ').Trim()
    $dashes = [string][char]0x2013 + [string][char]0x2014
    $segs = @([regex]::Split($t, '\s+[|' + $dashes + ']\s+|\s+-\s+|\s+::\s+') | Where-Object { $_.Trim().Length -gt 0 })
    if ($segs.Count -gt 1) {
        $best = $segs[0]
        foreach ($s in $segs) { if ($s.Length -gt $best.Length) { $best = $s } }
        $t = $best.Trim()
    }
    if ($t.Length -gt 120) { $t = ($t.Substring(0, 120) -replace '\s+\S*$', '') }
    return $t
}

# ---------- reading more of the page (all local, nothing is sent anywhere) ----------

# the characters a document must not gain: comma, colon, semicolon, hyphen, underscore, quotes, backtick and full stops
$script:LT_Quotes = '[' + [char]0x201C + [char]0x201D + [char]0x201E + [char]0x2033 + '"`]'
$script:LT_Dashes = '[' + [char]0x2010 + [char]0x2011 + [char]0x2012 + [char]0x2013 + [char]0x2014 + [char]0x2015 + [char]0x2212 + '\-_]'
$script:LT_Commas = '[,;:' + [char]0x3001 + ']'

function Clean-LinkText([string]$t, [int]$max = 110) {
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace $script:LT_Quotes, ''
    $t = $t -replace $script:LT_Dashes, ' '
    $t = $t -replace $script:LT_Commas, ' '
    $t = $t -replace '(?<=\d)\.(?=\d)', ' point '
    $t = $t -replace '(?<=\w)\.(?=\w)', ''
    $t = $t -replace '\.', ' '
    $t = ($t -replace '\s+', ' ').Trim()
    if ($t.Length -gt $max) { $t = ($t.Substring(0, $max) -replace '\s+\S*$', '').Trim() }
    return $t
}

function Strip-Tags([string]$s) {
    $s = [regex]::Replace($s, '(?is)<(script|style|noscript|svg)[^>]*>.*?</\1>', ' ')
    $s = [regex]::Replace($s, '<[^>]+>', ' ')
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    return ($s -replace '\s+', ' ').Trim()
}

function Get-MetaContent([string]$html, [string[]]$names) {
    foreach ($n in $names) {
        $q = [regex]::Escape($n)
        $m = [regex]::Match($html, '(?is)<meta[^>]+(?:name|property)\s*=\s*["'']' + $q + '["''][^>]*?content\s*=\s*["'']([^"'']*)["'']')
        if (-not $m.Success) { $m = [regex]::Match($html, '(?is)<meta[^>]+content\s*=\s*["'']([^"'']*)["''][^>]*?(?:name|property)\s*=\s*["'']' + $q + '["'']') }
        if ($m.Success) { $v = ([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) -replace '\s+', ' ').Trim(); if ($v) { return $v } }
    }
    return ''
}

function Get-WordCount([string]$t) { return @(($t -split '\s+') | Where-Object { $_ }).Count }

# a title that says little: a site name, "Home", an error or a bot check page
function Test-WeakTitle([string]$t, [string]$site) {
    $n = (($t.ToLowerInvariant() -replace '[^a-z0-9 ]', ' ') -replace '\s+', ' ').Trim()
    if (-not $n) { return $true }
    if ($n -match '^(home|homepage|home page|welcome|welcome home|index|default|main page|main|untitled|untitled document|document|loading|redirecting|please wait|just a moment|attention required|access denied|forbidden|not found|page not found|404|404 not found|error|an error occurred|sign in|log in|login|403 forbidden|500 internal server error|service unavailable|website|web site)$') { return $true }
    if ($n -match '^(404|403|500|503)\b') { return $true }
    if ($n -match '(just a moment|access denied|page not found|are you a robot|verify you are human|enable javascript|checking your browser|checking if the site connection is secure|one moment please|ddos protection|security check|captcha)') { return $true }
    if ($site) { $sn = (($site.ToLowerInvariant() -replace '[^a-z0-9 ]', ' ') -replace '\s+', ' ').Trim(); if ($sn -and $n -eq $sn) { return $true } }
    if ((Get-WordCount $n) -le 2 -and $n.Length -lt 18) { return $true }
    return $false
}

# first sentence of a description, short enough to be link text
function Get-FirstSentence([string]$d) {
    $d = ($d -replace '\s+', ' ').Trim()
    if (-not $d) { return '' }
    $m = [regex]::Match($d, '^(.+?[a-z0-9\)])[.!?](\s|$)')
    if ($m.Success) { $d = $m.Groups[1].Value }
    return $d
}

# best local wording for a web page: its title when it is a real title, otherwise its main heading,
# otherwise the start of its description. Returns Text and Source.
function Get-PageSuggestion([string]$html) {
    $site = Get-MetaContent $html @('og:site_name', 'application-name')
    $rawTitle = Get-MetaContent $html @('og:title', 'twitter:title')
    if (-not $rawTitle) { $tm = [regex]::Match($html, '(?is)<title[^>]*>(.*?)</title>'); if ($tm.Success) { $rawTitle = $tm.Groups[1].Value } }
    $title = ''; if ($rawTitle) { $title = Clean-LinkText (Clean-Title $rawTitle) }
    $h1 = ''; $hm = [regex]::Match($html, '(?is)<h1[^>]*>(.*?)</h1>'); if ($hm.Success) { $h1 = Clean-LinkText (Strip-Tags $hm.Groups[1].Value) }
    $blockRx = '(just a moment|access denied|page not found|are you a robot|verify you are human|checking your browser|checking if the site connection is secure|one moment please|ddos protection|captcha|^404|^403|not found$)'
    if (($rawTitle -match $blockRx) -or ($h1 -match $blockRx)) { return [pscustomobject]@{ Text = ''; Source = 'the page is a security check or error page' } }
    $desc = Get-MetaContent $html @('description', 'og:description', 'twitter:description')
    $descText = ''; if ($desc) { $descText = Clean-LinkText (Get-FirstSentence $desc) 100 }

    if ($title -and -not (Test-WeakTitle $title $site)) { return [pscustomobject]@{ Text = $title; Source = 'page title' } }
    # the title says little: look for something that says more
    if ($h1 -and (Get-WordCount $h1) -ge 3 -and -not (Test-WeakTitle $h1 $site)) { return [pscustomobject]@{ Text = $h1; Source = 'page heading, please check' } }
    if ($descText -and (Get-WordCount $descText) -ge 4) { return [pscustomobject]@{ Text = $descText; Source = 'page description, please check' } }
    if ($h1 -and (Get-WordCount $h1) -ge 2) { return [pscustomobject]@{ Text = $h1; Source = 'page heading, please check' } }
    return [pscustomobject]@{ Text = ''; Source = '' }
}

# page bytes to text, using the character set the page declares
function Decode-Html([byte[]]$bytes, [string]$contentType) {
    $enc = $null
    try {
        $cs = [regex]::Match([string]$contentType, '(?i)charset\s*=\s*["'']?([\w\-]+)')
        if (-not $cs.Success) {
            $peek = [Text.Encoding]::GetEncoding(28591).GetString($bytes, 0, [Math]::Min(4000, $bytes.Length))
            $cs = [regex]::Match($peek, '(?i)<meta[^>]+charset\s*=\s*["'']?([\w\-]+)')
        }
        if ($cs.Success) { $enc = [Text.Encoding]::GetEncoding($cs.Groups[1].Value) }
    } catch { $enc = $null }
    if (-not $enc) { $enc = [Text.Encoding]::UTF8 }
    return $enc.GetString($bytes)
}

function Get-PdfTitle([byte[]]$bytes) {
    $latin = [Text.Encoding]::GetEncoding(28591)
    $s = $latin.GetString($bytes)
    $title = ''
    $m = [regex]::Match($s, '/Title\s*\(((?:\\.|[^\\)])*)\)')
    if ($m.Success) {
        $raw = $m.Groups[1].Value
        $raw = [regex]::Replace($raw, '\\([()\\])', '$1')
        $raw = $raw.Replace('\n', ' ').Replace('\r', ' ')
        if ($raw.Length -ge 2 -and [int][char]$raw[0] -eq 254 -and [int][char]$raw[1] -eq 255) {
            $b = $latin.GetBytes($raw.Substring(2)); $title = [Text.Encoding]::BigEndianUnicode.GetString($b)
        } else { $title = $raw }
    }
    if (-not $title) {
        $m = [regex]::Match($s, '/Title\s*<([0-9A-Fa-f\s]+)>')
        if ($m.Success) {
            $hex = ($m.Groups[1].Value -replace '\s', '')
            if ($hex.Length % 2 -eq 0 -and $hex.Length -gt 0) {
                $b = New-Object byte[] ($hex.Length / 2)
                for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16) }
                if ($b.Length -ge 2 -and $b[0] -eq 254 -and $b[1] -eq 255) { $title = [Text.Encoding]::BigEndianUnicode.GetString($b, 2, $b.Length - 2) } else { $title = $latin.GetString($b) }
            }
        }
    }
    if (-not $title) {
        $u8 = [Text.Encoding]::UTF8.GetString($bytes)
        $m = [regex]::Match($u8, '(?s)<dc:title>.*?<rdf:li[^>]*>(.*?)</rdf:li>')
        if ($m.Success) { $title = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) }
    }
    $title = ($title -replace '\s+', ' ').Trim()
    $title = ($title -replace '^(?i)microsoft (word|powerpoint) - ', '').Trim()
    if ($title -match '^(?i)(untitled|document\d*|slide \d+)$') { $title = '' }
    return $title
}

# A direct request can be blocked by bot-protection that checks things a
# plain script cannot reproduce (the TLS handshake itself, for example) -
# this asks a real browser already on the machine (Edge, or Chrome if
# Edge is not there) to fetch the page instead, headless, which looks
# like actual browser traffic because it is one. Only used when the
# direct request has already failed. Degrades silently (returns $null)
# if neither browser is installed or it does not respond - never holds
# up the rest of the tool.
function Get-BrowserPath {
    # LF_BROWSER: use this browser instead of looking for Edge or Chrome (for testing)
    if ($env:LF_BROWSER -and (Test-Path -LiteralPath $env:LF_BROWSER)) { return $env:LF_BROWSER }
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LocalAppData) | Where-Object { $_ }
    # Edge first - it ships with every current copy of Windows, so it is
    # the one most likely to already be there - Chrome next, if installed.
    foreach ($leaf in @('Microsoft\Edge\Application\msedge.exe', 'Google\Chrome\Application\chrome.exe')) {
        foreach ($r in $roots) {
            $c = Join-Path $r $leaf
            if (Test-Path -LiteralPath $c) { return $c }
        }
    }
    foreach ($exe in @('msedge.exe', 'chrome.exe')) {
        $cmd = Get-Command $exe -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Get-PageViaEdge([string]$url) {
    $edge = Get-BrowserPath
    if (-not $edge) { return $null }
    $profileDir = Join-Path ([IO.Path]::GetTempPath()) ('lf_edge_' + [guid]::NewGuid().ToString('N'))
    $p = $null
    try {
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = $edge
        $psi.Arguments = '--headless=new --disable-gpu --disable-extensions --no-first-run --mute-audio --virtual-time-budget=9000 ' + $env:LF_BROWSER_ARGS + ' --user-data-dir="' + $profileDir + '" --dump-dom "' + $url + '"'
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $p = [Diagnostics.Process]::Start($psi)
        $outTask = $p.StandardOutput.ReadToEndAsync()
        $null = $p.StandardError.ReadToEndAsync()
        $finished = $p.WaitForExit(14000)
        if (-not $finished) { try { $p.Kill() } catch { }; return $null }
        $html = $outTask.GetAwaiter().GetResult()
        if (-not $html -or $html.Trim().Length -lt 20) { return $null }
        return $html
    } catch { return $null }
    finally {
        if ($p) { try { if (-not $p.HasExited) { $p.Kill() } } catch { } }
        try { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
}

# ---------- is the destination working? ----------
# The destination is checked FIRST. Only a destination that works may be used to build
# a suggestion from the page itself (title, heading, description, PDF title). For a broken
# or inaccessible destination nothing from the page is ever used - see Get-LinkSuggestion.

# wording that marks an error page, an access-denied page, a CAPTCHA or a bot check.
# Title / heading patterns are anchored so that real titles that merely contain a word
# like "error" ("Error handling in Python") are not mistaken for error pages.
# words that can also be part of a real article title: they only count on a SHORT page (a challenge page has almost no text)
$script:LF_ErrTitleWeakRx = 'captcha|security check\b|security verification|human verification|bot (check|verification|detection|protection)\b|unusual traffic'
$script:LF_ErrTitleRx = '^(404|403|401|500|502|503|504)\b|\b(404|403|500|503) (error|not found|forbidden|service unavailable|internal server error)|\berror (404|403|500|503)\b|\bhttp error\b|^(error|not found|forbidden|access denied|oops|unauthorized|bad gateway|server error|internal server error|service unavailable|gateway time.?out|site not found|page unavailable|an error occurred|something went wrong|sorry)[ !.]*$|\b(page|file|url|resource|site|document) (was |is )?not found\b|\bnot found$|access denied|permission denied|has been blocked|request blocked|just a moment|attention required|are you a robot|not a robot|verify you are human|verifying you are human|human verification|checking your browser|checking if the site connection is secure|one moment please|ddos protection|pardon our interruption|this site can.t be reached|can.t reach this page|no longer available|(page|content) (is )?(unavailable|not available)|(does not|doesn.t) exist|(can.t|cannot|couldn.t|could not) (be )?(find|found)'
# the same kind of wording inside the visible text of a SHORT page (a real article may mention these words)
$script:LF_ErrBodyRx = 'verify you are human|verifying you are human|are you a robot|not a robot|checking your browser|checking if the site connection is secure|enable javascript and cookies to continue|please enable cookies|please turn javascript on|complete the security check|performing security verification|press (and|&) hold|ddos protection by|unusual traffic from your|your request has been blocked|has been blocked|access to this page has been denied|access denied|you don.t have permission to access|errors? 40[134]|\b40[134] (not found|error|forbidden)|http error 40|page (cannot|can.t|could not|couldn.t) be found|(page|content) (is )?(not available|unavailable)|this page (isn.t available|is not available|doesn.t exist|does not exist)|(we|i) (can.t|cannot|couldn.t|could not) (find|locate) (that|the|this) page|(page|url|file) (you (are looking for|requested|were looking for)|requested) (is|was|could|cannot|can.t|does)|requested url was not found|not found on this server|this site can.t be reached|can.t reach this page|refused to connect|took too long to respond|err_[a-z_]{4,}|dns_probe|your connection is not private|something went wrong|robot check|bot check|before you continue to'
# scripts that bot-protection services place on their challenge pages
$script:LF_ChallengeRx = 'g-recaptcha|h-captcha|hcaptcha|cf-turnstile|cf-chl|challenge-platform|_incapsula_resource|px-captcha|captcha-delivery|recaptcha/api'

function Get-ErrorKind([string]$s) {
    if ($s -match 'captcha|robot|human|checking (your|if)|just a moment|one moment|security (check|verification)|bot |ddos|interruption|unusual traffic|press (and|&) hold|cookies|javascript') { return 'the page is a CAPTCHA or bot check page' }
    if ($s -match 'denied|forbidden|blocked|permission|unauthorized|^40[13]') { return 'the page says access is denied' }
    if ($s -match 'not found|40[4]|no longer|exist|find|found|available|gone') { return 'the page says the address was not found' }
    return 'the page is an error page'
}

# '' when the page looks like a real page, otherwise the reason it does not.
# -RequireContent is for a page that was only read through the browser (the browser cannot tell us the HTTP status):
# a branded error page often has a normal-looking title but almost no text, so a page with under 300 characters
# of visible text is not accepted there as proof that the destination works.
function Test-ErrorOrBlockPage([string]$html, [bool]$requireContent = $false) {
    if (-not $html -or -not $html.Trim()) { return 'the page is empty' }
    $head = $html; if ($head.Length -gt 400000) { $head = $head.Substring(0, 400000) }
    $title = ''; $tm = [regex]::Match($head, '(?is)<title[^>]*>(.*?)</title>'); if ($tm.Success) { $title = [System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value) }
    $h1 = ''; $hm = [regex]::Match($head, '(?is)<h1[^>]*>(.*?)</h1>'); if ($hm.Success) { $h1 = Strip-Tags $hm.Groups[1].Value }
    $og = Get-MetaContent $head @('og:title', 'twitter:title')
    $vis = Strip-Tags $head
    foreach ($n in @($title, $og, $h1)) {
        $x = (([string]$n -replace '\s+', ' ').Trim()).ToLowerInvariant()
        if ($x -and ($x -match $script:LF_ErrTitleRx)) { return (Get-ErrorKind $x) }
        if ($x -and $vis.Length -lt 1500 -and ($x -match $script:LF_ErrTitleWeakRx)) { return (Get-ErrorKind $x) }
    }
    if ($vis.Length -lt 1500) {
        $lv = $vis.ToLowerInvariant()
        if ($lv -match $script:LF_ErrBodyRx) { return (Get-ErrorKind $Matches[0]) }
        if ($head -match $script:LF_ChallengeRx) { return 'the page is a CAPTCHA or bot check page' }
    }
    if ($requireContent -and $vis.Length -lt 300) { return 'the page has almost no content' }
    return ''
}

# The name of the website, read from the page the site itself shows for a link that does not work (its "page not
# found" page). Used as the text for a broken link. It is looked for in this order: the site-name tag of the page,
# the application name, an element the page marks as its site name or logo, the structured data of the page, the
# first line of the page header that reads like a name (banners such as "An official website" are skipped), and
# last the part of the page title that is not the error. Nothing is made from the web address.
function Get-SiteName([string]$html, [string]$url) {
    if (-not $html) { return '' }
    $head = $html; if ($head.Length -gt 400000) { $head = $head.Substring(0, 400000) }
    $head = [regex]::Replace($head, '(?is)<(script|style|noscript|svg|template)\b(?![^>]*ld\+json)[^>]*>.*?</\1>', ' ')
    $skip = '(?i)skip|official (website|site)|here.s how you know|\.gov|\bgovernment\b|secure|cookie|privacy|menu|search|log ?in|sign ?in|register|subscribe|contact|toggle|navigation|close|language|translate|espa.ol|home page|homepage|^home$|not found|error|404|oops|sorry|javascript|loading|^(skip|main|content)'
    $cands = New-Object System.Collections.Generic.List[string]
    $m1 = Get-MetaContent $head @('og:site_name', 'application-name', 'twitter:site_name'); if ($m1) { $cands.Add($m1) }
    foreach ($m in [regex]::Matches($head, '(?is)<(?:a|div|span|p|h1|h2|h3|strong|b|figure)\b[^>]*\b(?:class|id)\s*=\s*"[^"]*(?:site-?(?:name|title)|brand|org(?:anization)?-?name|agency-?name|masthead-?title|header-?title|logo-?text)[^"]*"[^>]*>(.*?)</(?:a|div|span|p|h1|h2|h3|strong|b|figure)>')) { $cands.Add((Strip-Tags $m.Groups[1].Value)) }
    foreach ($m in [regex]::Matches($head, '(?is)<img\b[^>]*\bclass\s*=\s*"[^"]*(?:logo|brand)[^"]*"[^>]*\balt\s*=\s*"([^"]+)"')) { $cands.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)) }
    foreach ($m in [regex]::Matches($head, '(?is)<script[^>]+ld\+json[^>]*>(.*?)</script>')) {
        $j = $m.Groups[1].Value
        if ($j -match '"@type"\s*:\s*\[?\s*"(Organization|GovernmentOrganization|WebSite|NewsMediaOrganization|EducationalOrganization|MedicalOrganization|Corporation|LocalBusiness)"') {
            $nm = [regex]::Match($j, '"name"\s*:\s*"((?:[^"\\]|\\.)*)"'); if ($nm.Success) { $cands.Add(($nm.Groups[1].Value -replace '\\"', '"')) }
        }
    }
    # the first line of the page header that reads like a name: a few capitalised words, no sentence
    $hm = [regex]::Match($head, '(?is)<header\b[^>]*>(.*?)</header>')
    $region = ''
    if ($hm.Success) { $region = $hm.Groups[1].Value } else { $bm = [regex]::Match($head, '(?is)<body\b[^>]*>(.{0,12000})'); if ($bm.Success) { $region = $bm.Groups[1].Value } }
    # what is specific to this one page (its own heading and title tags) is never taken as the site name
    $own = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($head, '(?is)<h[12]\b[^>]*>(.*?)</h[12]>')) { $own.Add((((Strip-Tags $m.Groups[1].Value) -replace '\s+', ' ').Trim().ToLowerInvariant())) }
    $ogt = Get-MetaContent $head @('og:title', 'twitter:title'); if ($ogt) { $own.Add($ogt.ToLowerInvariant()) }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($region, '>([^<>]{8,140})<')) {
        $t = ([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) -replace '\s+', ' ').Trim()
        if (-not $t -or $t -match $skip -or $t -match '[.!?:;|]' -or $own.Contains($t.ToLowerInvariant())) { continue }
        $ws = @($t -split ' ' | Where-Object { $_ })
        if ($ws.Count -lt 2 -or $ws.Count -gt 12) { continue }
        $cap = @($ws | Where-Object { $_ -cmatch '^[A-Z0-9(]' -or $_ -cmatch '^(of|and|the|for|de|la|du|von|van|&)$' }).Count
        if ($cap / $ws.Count -ge 0.75) { $lines.Add($t) }
    }
    if ($lines.Count -gt 0) { $cands.Add($lines[0]) }
    # the part of the page title that is not the error (Page not found | Site name)
    $tm = [regex]::Match($head, '(?is)<title[^>]*>(.*?)</title>')
    if ($tm.Success) {
        $dashes = [string][char]0x2013 + [string][char]0x2014
        $segs = New-Object System.Collections.Generic.List[string]
        foreach ($seg in @([regex]::Split([System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value), '\s+[|' + $dashes + ']\s+|\s+-\s+|\s+::\s+|\s*\|\s*'))) {
            $q = ($seg -replace '\s+', ' ').Trim()
            if (-not $q -or $q -match $skip -or $q -match '(?i)page|found|unavailable|exist') { continue }
            $ql = $q.ToLowerInvariant(); $isOwn = $false
            foreach ($o in $own) { if ($o -and ($o -eq $ql -or $o.Contains($ql) -or $ql.Contains($o))) { $isOwn = $true } }
            if (-not $isOwn) { $segs.Add($q) }
        }
        # only a title part that is clearly the site, not another page title: the one that is left once the page's own heading is taken out
        if ($segs.Count -eq 1) { $cands.Add($segs[0]) }
    }
    $generic = '^(home|homepage|welcome|menu|search|error|not found|page not found|untitled|loading|website|site)$'
    foreach ($c in $cands) {
        $t = ([string]$c -replace '\s&\s', ' and ' -replace '\s+', ' ').Trim()
        if (-not $t) { continue }
        if (($t -ceq $t.ToUpperInvariant()) -and ($t -cmatch '[A-Z]') -and $t.Length -gt 6) { $t = (New-Object Globalization.CultureInfo('en-US')).TextInfo.ToTitleCase($t.ToLowerInvariant()) }
        $t = Clean-LinkText $t
        $n = Get-WordCount $t
        if ($n -ge 1 -and $n -le 12 -and $t -notmatch $generic -and $t -notmatch '(?i)not found|page not found|\b404\b|access denied|just a moment') { return $t }
    }
    return ''
}

# Decides whether the destination works, before any suggestion is made from it.
#   1. A direct request, the way a script can make it.
#   2. If that fails (or lands on an error, CAPTCHA or bot check page) a real browser installed on this
#      computer (Edge, or Chrome) is asked to open the address. A page that opens there is a working
#      destination - being blocked for an automated request does not make a link broken.
#      "Not found" (404) and "gone" (410) answers and unknown site names are final; the browser is not asked.
# Returns Working, Reason (why it is not working), Bytes and ContentType (direct reply), EdgeHtml (page as the browser shows it), IsPdf.
function Get-LinkStatus([string]$url) {
    $isPdf = ($url -match '(?i)\.pdf($|[?#])')
    $reason = ''; $final = $false
    $bytes = $null; $ct = ''; $errHtml = ''; $errKind = ''
    try {
        $req = [Net.HttpWebRequest]::Create($url)
        $req.Timeout = 15000; $req.ReadWriteTimeout = 15000
        $req.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
        $req.Headers.Add('Accept-Language', 'en-US,en;q=0.9')
        $req.Headers.Add('Upgrade-Insecure-Requests', '1')
        $req.Headers.Add('Sec-Fetch-Site', 'none')
        $req.Headers.Add('Sec-Fetch-Mode', 'navigate')
        $req.Headers.Add('Sec-Fetch-User', '?1')
        $req.Headers.Add('Sec-Fetch-Dest', 'document')
        $req.KeepAlive = $true
        $req.AllowAutoRedirect = $true; $req.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,application/pdf,*/*;q=0.8'
        $resp = $req.GetResponse()
        try {
            $ct = [string]$resp.ContentType
            if ($ct -match 'pdf') { $isPdf = $true }
            $stream = $resp.GetResponseStream()
            $ms = New-Object IO.MemoryStream
            $buf = New-Object byte[] 65536; $limit = 3000000
            while ($ms.Length -lt $limit) { $n = $stream.Read($buf, 0, $buf.Length); if ($n -le 0) { break }; $ms.Write($buf, 0, $n) }
            $bytes = $ms.ToArray()
        } finally { $resp.Close() }
        # the server answered. A PDF or another kind of file is a working destination; a web page is
        # only a working destination when it is not an error, CAPTCHA or bot check page.
        if ($isPdf -or ($ct -and $ct -notmatch '(?i)html|xml|text/plain')) {
            return [pscustomobject]@{ Working = $true; Reason = ''; Bytes = $bytes; ContentType = $ct; EdgeHtml = $null; IsPdf = $isPdf }
        }
        $why = Test-ErrorOrBlockPage (Decode-Html $bytes $ct)
        if (-not $why) {
            return [pscustomobject]@{ Working = $true; Reason = ''; Bytes = $bytes; ContentType = $ct; EdgeHtml = $null; IsPdf = $isPdf }
        }
        $reason = $why
        $errHtml = Decode-Html $bytes $ct; $errKind = $why
    } catch {
        $reason = 'the page could not be reached'
        # A direct .NET method call that throws comes back wrapped in a
        # MethodInvocationException - the actual WebException (and the
        # HTTP status code on it) is one level deeper, in InnerException.
        $we = $_.Exception
        if ($we -and ($we -isnot [Net.WebException]) -and ($we.InnerException -is [Net.WebException])) { $we = $we.InnerException }
        if ($we -is [Net.WebException] -and $we.Response) {
            try {
                $code = [int]$we.Response.StatusCode
                $reason = 'the page replied with ' + $code + ' ' + $we.Response.StatusDescription
                if ($code -eq 404 -or $code -eq 410) { $final = $true }
                # the error page the site shows for this answer, when it sends one
                try {
                    $ect = [string]$we.Response.ContentType
                    if ($ect -match '(?i)html|xml|text') {
                        $rs2 = $we.Response.GetResponseStream(); $ms2 = New-Object IO.MemoryStream; $b2 = New-Object byte[] 65536
                        while ($ms2.Length -lt 600000) { $n2 = $rs2.Read($b2, 0, $b2.Length); if ($n2 -le 0) { break }; $ms2.Write($b2, 0, $n2) }
                        $errHtml = Decode-Html $ms2.ToArray() $ect
                    }
                } catch { }
                if ($code -eq 404 -or $code -eq 410) { $errKind = 'the page says the address was not found' }
                elseif ($code -eq 401 -or $code -eq 403) { $errKind = 'the page says access is denied' }
                else { $bk = ''; if ($errHtml) { $bk = Test-ErrorOrBlockPage $errHtml }; if ($bk) { $errKind = $bk } else { $errKind = 'the page is an error page' } }
            } catch { }
        } elseif ($we -and $we.Message -match '(?i)timed out') { $reason = 'the page took too long to reply' }
        elseif ($we -and $we.Message -match '(?i)SSL|TLS|certificate|trust') { $reason = 'a secure connection could not be made' }
        if ($we -is [Net.WebException] -and $we.Status -eq [Net.WebExceptionStatus]::NameResolutionFailure) { $final = $true; $reason = 'the site could not be found' }
    }
    # not working for a script: ask a real browser before calling the link broken
    if (-not $isPdf -and -not $final) {
        $edgeHtml = Get-PageViaEdge $url
        if ($edgeHtml) {
            $why2 = Test-ErrorOrBlockPage $edgeHtml $true
            if (-not $why2) {
                return [pscustomobject]@{ Working = $true; Reason = $reason; Bytes = $null; ContentType = ''; EdgeHtml = $edgeHtml; IsPdf = $isPdf }
            }
            # what the real browser shows is an error page: that page (not the answer the script got) says why
            if ($why2 -ne 'the page is empty') { $errHtml = $edgeHtml; $errKind = $why2 }
        }
    }
    return [pscustomobject]@{ Working = $false; Reason = $reason; Bytes = $null; ContentType = ''; EdgeHtml = $null; IsPdf = $isPdf; ErrHtml = $errHtml; ErrKind = $errKind }
}

# ---------- YouTube ----------
# A YouTube link is described by the title of the video: "Watch the YouTube video of <title>".
function Get-YouTubeInfo([string]$url) {
    $u = $null
    try { $u = [uri]$url } catch { return $null }
    $h = $u.Host.ToLowerInvariant()
    $isYt = ($h -match '^(www\.|m\.|music\.)?youtube(-nocookie)?\.com$')
    $isShort = ($h -eq 'youtu.be')
    if (-not ($isYt -or $isShort)) { return $null }
    $path = $u.AbsolutePath; $q = $u.Query
    $id = ''
    if ($isShort) { $m = [regex]::Match($path, '^/([A-Za-z0-9_-]{6,})'); if ($m.Success) { $id = $m.Groups[1].Value } }
    else {
        $m = [regex]::Match($path, '^/(shorts|live|embed|v)/([A-Za-z0-9_-]{6,})')
        if ($m.Success) { $id = $m.Groups[2].Value }
        elseif ($path -eq '/watch') { $m = [regex]::Match($q, '(?:^\?|&)v=([A-Za-z0-9_-]{6,})'); if ($m.Success) { $id = $m.Groups[1].Value } }
    }
    if ($id) { return [pscustomobject]@{ Kind = 'video'; Id = $id } }
    $m = [regex]::Match($q, '(?:^\?|&)list=([A-Za-z0-9_-]{6,})')
    if ($m.Success -and ($path -eq '/playlist' -or $path -eq '/watch')) { return [pscustomobject]@{ Kind = 'playlist'; Id = $m.Groups[1].Value } }
    return $null
}

function Get-YouTubeSentence($info, [string]$title) {
    $t = Clean-LinkText $title
    if (-not $t -or $t -match '^(?i)youtube$') { return '' }
    if ($info.Kind -eq 'playlist') { return 'Watch the YouTube playlist of ' + $t }
    return 'Watch the YouTube video of ' + $t
}

# YouTube's public oEmbed address gives the title of a video without loading the page (and without a consent page).
# Returns Status ok / notfound / unknown and Title.
function Get-YouTubeOEmbed($info) {
    $target = 'https://www.youtube.com/watch?v=' + $info.Id
    if ($info.Kind -eq 'playlist') { $target = 'https://www.youtube.com/playlist?list=' + $info.Id }
    $base = 'https://www.youtube.com/oembed'
    if ($env:LF_YT_OEMBED) { $base = $env:LF_YT_OEMBED }
    $api = $base + '?format=json&url=' + [uri]::EscapeDataString($target)
    try {
        $req = [Net.HttpWebRequest]::Create($api)
        $req.Timeout = 15000; $req.ReadWriteTimeout = 15000
        $req.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
        $resp = $req.GetResponse()
        try { $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8); $json = $sr.ReadToEnd() } finally { $resp.Close() }
        $o = $json | ConvertFrom-Json
        return [pscustomobject]@{ Status = 'ok'; Title = [string]$o.title }
    } catch {
        $we = $_.Exception
        if ($we -and ($we -isnot [Net.WebException]) -and ($we.InnerException -is [Net.WebException])) { $we = $we.InnerException }
        if ($we -is [Net.WebException] -and $we.Response) {
            try { if ([int]$we.Response.StatusCode -eq 404) { return [pscustomobject]@{ Status = 'notfound'; Title = '' } } } catch { }
        }
    }
    return [pscustomobject]@{ Status = 'unknown'; Title = '' }
}

# The wording for a web page that works. For a YouTube video it is the video title; for every other page it is
# Get-PageSuggestion exactly as before (title, then heading, then description).
function Get-PageTextSuggestion([string]$url, [string]$html) {
    $yt = Get-YouTubeInfo $url
    if ($yt) {
        $raw = Get-MetaContent $html @('og:title', 'twitter:title')
        if (-not $raw) { $tm = [regex]::Match($html, '(?is)<title[^>]*>(.*?)</title>'); if ($tm.Success) { $raw = [System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value) } }
        $raw = ((($raw + '') -replace '\s+', ' ').Trim()) -replace '(?i)\s*[-|]\s*YouTube$', ''
        $s = Get-YouTubeSentence $yt $raw
        if ($s) { return [pscustomobject]@{ Text = $s; Source = 'YouTube video title' } }
    }
    return (Get-PageSuggestion $html)
}

# ---------- Go to Link: open the page, then read it once it has loaded ----------
# Go to Link opens the address in its own Edge (or Chrome) window, with a temporary profile and a debugging
# port that only this computer can reach. The person can pass any check in that window. When the page has
# finished loading, the tool reads its title, heading and description from the window and suggests the text.
# Nothing is sent anywhere. If neither browser is installed, Go to Link simply opens the address normally.
$script:Watch = @{}

# One helper browser (own temporary profile) is started for the first Go to Link. Every later Go to Link opens a
# new tab in that same window, which is much quicker than starting the browser again.
$script:Shared = $null

function Start-WatchBrowser([string]$url) {
    $exe = Get-BrowserPath
    if (-not $exe) { return $null }
    $sh = $script:Shared
    if ($sh -and $sh.Proc -and -not $sh.Proc.HasExited -and (Get-WatchPort $sh)) {
        try {
            $wc = New-Object Net.WebClient; $wc.Proxy = $null
            $r = $wc.UploadString("http://127.0.0.1:$($sh.Port)/json/new?about:blank", 'PUT', '') | ConvertFrom-Json
            if ($r -and $r.id) {
                try { $null = $wc.DownloadString("http://127.0.0.1:$($sh.Port)/json/activate/$($r.id)") } catch { }
                # the address is opened from inside the tab, so a #section or ?query part is kept exactly
                $null = Invoke-Cdp $r.webSocketDebuggerUrl ('location.href = ' + ($url | ConvertTo-Json -Compress))
                return [pscustomobject]@{ Proc = $sh.Proc; Profile = $sh.Profile; Port = $sh.Port; Started = (Get-Date); Connected = $true; TargetId = [string]$r.id; Url = $url }
            }
        } catch { }
    }
    $profileDir = Join-Path ([IO.Path]::GetTempPath()) ('lf_watch_' + [guid]::NewGuid().ToString('N'))
    try {
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = $exe
        $psi.Arguments = '--user-data-dir="' + $profileDir + '" --remote-debugging-port=0 --no-first-run --no-default-browser-check --log-level=3 --new-window ' + $env:LF_BROWSER_ARGS + ' "' + $url + '"'
        $psi.UseShellExecute = $false
        $p = [Diagnostics.Process]::Start($psi)
        $w = [pscustomobject]@{ Proc = $p; Profile = $profileDir; Port = 0; Started = (Get-Date); Connected = $false; TargetId = ''; Url = $url }
        $script:Shared = $w
        return $w
    } catch { return $null }
}

function Get-WatchPort($w) {
    if ($w.Port) { return $w.Port }
    $f = Join-Path $w.Profile 'DevToolsActivePort'
    if (Test-Path -LiteralPath $f) {
        try { $l = [IO.File]::ReadAllLines($f); if ($l.Count -ge 1 -and $l[0] -match '^\d+$') { $w.Port = [int]$l[0] } } catch { }
    }
    return $w.Port
}

# one DevTools call: evaluates a JavaScript expression in the page and returns the reply
function Invoke-Cdp([string]$wsUrl, [string]$expr) {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $cts = New-Object Threading.CancellationTokenSource(8000)
    try {
        $ws.ConnectAsync([uri]$wsUrl, $cts.Token).GetAwaiter().GetResult()
        $msg = @{ id = 1; method = 'Runtime.evaluate'; params = @{ expression = $expr; returnByValue = $true } } | ConvertTo-Json -Depth 5 -Compress
        $bytes = [Text.Encoding]::UTF8.GetBytes($msg)
        $null = $ws.SendAsync((New-Object 'ArraySegment[byte]' (, $bytes)), [Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).GetAwaiter().GetResult()
        $buf = New-Object byte[] 65536
        $seg = New-Object 'ArraySegment[byte]' (, $buf)
        $ms = New-Object IO.MemoryStream
        while ($true) {
            $r = $ws.ReceiveAsync($seg, $cts.Token).GetAwaiter().GetResult()
            if ($r.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) { return $null }
            $ms.Write($buf, 0, $r.Count)
            if ($r.EndOfMessage) {
                $o = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
                if ($o.id -eq 1) { return $o }
                $ms.SetLength(0)
            }
        }
    } catch { return $null } finally { $ws.Dispose(); $cts.Dispose() }
}

# Reading what the person sees on the opened page, for when its title, heading tag and description say nothing.
# The page itself is asked (inside the browser window) for its main headings and its blocks of text, left out are
# menus, footers and side bars; capitals produced by the page's styling are what innerText returns.
$script:LF_VisibleJs = @'
(function(){
var root=document.querySelector("main,[role=main]")||document.body; if(!root) return "{}";
function bad(e){return e.closest("nav,footer,aside,[role=navigation],[role=banner],[role=contentinfo]")||(e.closest("header")&&!e.closest("main,article,[role=main]"));}
function vis(e){return e.getClientRects().length>0;}
var a="",fr="";try{fr=decodeURIComponent(location.hash.replace(/^#\/?/,""));}catch(e){fr=location.hash.replace(/^#\/?/,"");}
if(fr){var el=document.getElementById(fr);if(!el){var ids=document.querySelectorAll("[id],[name]");for(var q=0;q<ids.length;q++){var iv=(ids[q].id||ids[q].getAttribute("name")||"").toLowerCase();if(iv===fr.toLowerCase()){el=ids[q];break;}}}
if(el){var HT=/^H[1-6]$/,tx=function(n){return n?(n.innerText||"").trim():"";};
if(HT.test(el.tagName)||el.getAttribute("role")==="heading")a=tx(el);
if(!a){var hh=el.querySelector("h1,h2,h3,h4,h5,h6,[role=heading]");a=tx(hh);}
if(!a&&/^(A|BUTTON|SUMMARY|LI|SPAN|LABEL)$/.test(el.tagName)){var st=tx(el);if(st&&st.length<=90)a=st;}
if(!a){var nx=el.nextElementSibling;if(nx&&(HT.test(nx.tagName)))a=tx(nx);}
if(!a&&el.parentElement){var ph=el.parentElement.querySelector("h1,h2,h3,h4,h5,h6");a=tx(ph);}}}
var h=[],hs=root.querySelectorAll("h1,h2,[role=heading]");
for(var i=0;i<hs.length&&h.length<12;i++){var e=hs[i]; if(bad(e)||!vis(e)) continue; var t=(e.innerText||"").trim(); if(t) h.push(t);}
var b=[],bs=root.querySelectorAll("p,li,div,span,td,dd,blockquote,h3,h4,h5,h6");
for(var j=0;j<bs.length&&b.length<40;j++){var x=bs[j]; if(bad(x)||!vis(x)) continue; var blk=false;
for(var k=0;k<x.children.length;k++){ if(/^(P|DIV|UL|OL|LI|TABLE|SECTION|ARTICLE|FORM|H[1-6])$/.test(x.children[k].tagName)){blk=true;break;} }
if(blk) continue; var s=(x.innerText||"").trim(); if(s.length>=40) b.push(s.slice(0,600));}
return JSON.stringify({a:a,h:h,b:b});
})()
'@

# the main heading of the page: the first heading of two or more words that is not a menu or site label
function Select-HeadingText($heads, [int]$MinWords = 2) {
    $generic = '^(home|homepage|welcome|menu|main menu|navigation|search|skip to content|cookies?|privacy( policy)?|sign in|log in|login|register|subscribe|share|related|contact( us)?|follow us|footer|table of contents|on this page|loading|error)$'
    foreach ($h in @($heads)) {
        $t = ([string]$h -replace '\s+', ' ').Trim() -replace '\s&\s', ' and '
        if (-not $t) { continue }
        # a heading the page shows in capitals becomes normal capitalisation: SUPPORT MATERIALS -> Support Materials
        if (($t -ceq $t.ToUpperInvariant()) -and ($t -cmatch '[A-Z]')) { $t = (New-Object Globalization.CultureInfo('en-US')).TextInfo.ToTitleCase($t.ToLowerInvariant()) }
        $c = Clean-LinkText $t
        $n = Get-WordCount $c
        if ($n -ge $MinWords -and $n -le 14 -and ($c.ToLowerInvariant() -notmatch $generic)) { return $c }
    }
    return ''
}

# the first real sentence of the page's text: a block of at least six words, not a cookie, privacy or footer line
function Select-SentenceText($blocks) {
    $skip = 'toll free|switchboard|email us|call us|\bphone\b|\bfax\b|@|\d{3}[\s.-]?\d{3,4}[\s.-]?\d{4}|cookie|privacy|copyright|all rights reserved|terms (of|and)|subscribe|sign up|log in|sign in|skip to|javascript|your browser|accept all|\bmenu\b|' + [char]0x00A9
    foreach ($b in @($blocks)) {
        $t = ([string]$b -replace '\s+', ' ').Trim()
        if (-not $t -or $t -match $skip) { continue }
        if ($t -ceq $t.ToUpperInvariant()) { continue }
        $n = Get-WordCount $t
        if ($n -lt 6 -or $n -gt 80) { continue }
        $sents = @([regex]::Split($t, '(?<=[.!?])\s+') | Where-Object { $_.Trim() })
        $pick = $sents[0]
        if ((Get-WordCount $pick) -lt 6 -and $sents.Count -gt 1) { $pick = $sents[0] + ' ' + $sents[1] }
        $c = Clean-LinkText $pick 125
        if ((Get-WordCount $c) -ge 5) { return $c }
    }
    return ''
}

function Get-VisiblePageText([string]$wsUrl) {
    $res = Invoke-Cdp $wsUrl $script:LF_VisibleJs
    if (-not ($res -and $res.result -and $res.result.result -and $res.result.result.value)) { return '' }
    $d = $null
    try { $d = [string]$res.result.result.value | ConvertFrom-Json } catch { return '' }
    # a link that points to a section of the page (#something) is named after that section when it can be found
    if ($d.a) { $t = Select-HeadingText @($d.a) 1; if ($t) { return $t } }
    $t = Select-HeadingText $d.h
    if ($t) { return $t }
    return (Select-SentenceText $d.b)
}

# A link that points to a section (#something): the name of that section is read straight from the page. It is tried
# in this order: the element with that id when it is a heading, a heading whose words make the same name
# (Testing & Quizzing = testingandquizzing), the page's own jump link to that section, then the heading inside the id.
$script:LF_AnchorJs = @'
(function(fr){
try{fr=decodeURIComponent(fr);}catch(e){}
fr=fr.replace(/^\//,""); if(!fr) return "";
function tx(n){return n?(n.innerText||n.textContent||"").replace(/\s+/g," ").trim():"";}
function slug(t){return t.toLowerCase().replace(/&/g,"and").replace(/[^a-z0-9]/g,"");}
var want=slug(fr), lf=fr.toLowerCase(), HT=/^H[1-6]$/, el=document.getElementById(fr);
if(!el){var ids=document.querySelectorAll("[id],[name]");for(var q=0;q<ids.length;q++){var iv=(ids[q].id||ids[q].getAttribute("name")||"").toLowerCase();if(iv===lf){el=ids[q];break;}}}
if(el&&(HT.test(el.tagName)||el.getAttribute("role")==="heading")){var t0=tx(el);if(t0)return t0;}
var hs=document.querySelectorAll("h1,h2,h3,h4,h5,h6,[role=heading]");
for(var i=0;i<hs.length;i++){var t1=tx(hs[i]);if(t1&&t1.length<=90&&slug(t1)===want)return t1;}
var as=document.querySelectorAll("a[href*='#']");
for(var j=0;j<as.length;j++){var h=as[j].getAttribute("href")||"";var f=h.slice(h.lastIndexOf("#")+1);try{f=decodeURIComponent(f);}catch(e){}if(f.toLowerCase()===lf){var t2=tx(as[j]);if(t2&&t2.length<=90)return t2;}}
if(el){var hh=el.querySelector("h1,h2,h3,h4,h5,h6,[role=heading]");var t3=tx(hh);if(t3)return t3;
 if(/^(A|BUTTON|SUMMARY|LI|SPAN|LABEL)$/.test(el.tagName)){var t4=tx(el);if(t4&&t4.length<=90)return t4;}
 var nx=el.nextElementSibling;if(nx&&HT.test(nx.tagName)){var t5=tx(nx);if(t5)return t5;}}
return "";
})
'@

function Get-AnchorText([string]$wsUrl, [string]$fragment) {
    if (-not $fragment) { return '' }
    $res = Invoke-Cdp $wsUrl ('(' + $script:LF_AnchorJs.Trim() + ')(' + ($fragment | ConvertTo-Json -Compress) + ')')
    if (-not ($res -and $res.result -and $res.result.result)) { return '' }
    $t = ([string]$res.result.result.value -replace '\s+', ' ').Trim()
    if (-not $t) { return '' }
    return (Select-HeadingText @($t) 1)
}

# What the window shows right now: opening / loading / blocked / weak / found / closed / error
# titles that sites give their robot-check pages
$script:LF_CheckTitleRx = 'just a moment|attention required|verify you are human|verifying you are human|performing security verification|security check|checking your browser|checking if the site connection is secure|one moment,? please|ddos protection|are you a robot|not a robot|pardon our interruption|human verification|access denied|request blocked'

function Get-WatchState($w) {
    if ($w.Proc -and $w.Proc.HasExited) { return @{ state = 'closed' } }
    $port = Get-WatchPort $w
    if (-not $port) {
        if (((Get-Date) - $w.Started).TotalSeconds -gt 30) { return @{ state = 'error' } }
        return @{ state = 'opening' }
    }
    $targets = New-Object System.Collections.Generic.List[object]
    try {
        $wc = New-Object Net.WebClient; $wc.Proxy = $null
        $parsed = $wc.DownloadString("http://127.0.0.1:$port/json/list") | ConvertFrom-Json
        # Windows PowerShell 5.1 hands a JSON array over as one object, PowerShell 7 as separate items: accept both
        foreach ($t in @($parsed)) { if ($t -is [array]) { foreach ($u in $t) { $targets.Add($u) } } else { $targets.Add($t) } }
    } catch { }
    if ($targets.Count -eq 0) {
        if ($w.Connected) { return @{ state = 'closed' } }
        if (((Get-Date) - $w.Started).TotalSeconds -gt 30) { return @{ state = 'error' } }
        return @{ state = 'opening' }
    }
    $w.Connected = $true
    $pages = @($targets | Where-Object { $_.type -eq 'page' })
    if ($pages.Count -eq 0) { return @{ state = 'closed' } }
    if ($w.TargetId) {
        # this link has its own tab: when the tab is gone, the person closed it
        $page = $pages | Where-Object { $_.id -eq $w.TargetId } | Select-Object -First 1
        if (-not $page) { return @{ state = 'closed' } }
        if ($page.url -match '^chrome-error') { return @{ state = 'blocked' } }
        if ($page.url -match '^(devtools|chrome|edge|chrome-extension|about:)' -or -not $page.webSocketDebuggerUrl) { return @{ state = 'loading' } }
    } else {
        $page = $pages | Where-Object { $_.url -notmatch '^(devtools|chrome|edge|chrome-extension|chrome-error|about:)' -and $_.webSocketDebuggerUrl } | Select-Object -First 1
        if (-not $page) {
            if (@($pages | Where-Object { $_.url -match '^chrome-error' }).Count -gt 0) { return @{ state = 'blocked' } }
            return @{ state = 'loading' }
        }
        $w.TargetId = [string]$page.id
    }
    # While the site shows its robot check, the tool does not connect to the page at all: the title is taken from the
    # browser's own list of windows, so nothing on the page can tell that it is being watched. The page is read after
    # the title has changed to the real one.
    if ([string]$page.title -match $script:LF_CheckTitleRx) { return @{ state = 'blocked' } }
    # the page's real title is not known yet: wait a moment before the first look (a page without any title is looked at after 2 seconds)
    if (([string]$page.title -eq [string]$page.url -or -not [string]$page.title) -and ((Get-Date) - $w.Started).TotalSeconds -lt 2) { return @{ state = 'loading' } }
    $res = Invoke-Cdp $page.webSocketDebuggerUrl 'document.readyState+"\u0001"+location.href'
    $val = ''
    if ($res -and $res.result -and $res.result.result) { $val = [string]$res.result.result.value }
    if (-not $val) { return @{ state = 'loading' } }
    $parts = $val.Split([char[]]@([char]1), 2)
    if ($parts.Count -lt 2) { return @{ state = 'loading' } }
    $ready = $parts[0]; $pageUrl = $parts[1]
    if ($ready -eq 'loading') { return @{ state = 'loading' } }
    # a link to a section of a page: the section's name is read as soon as the page structure is there, without waiting for images and scripts
    $frag = ''
    $fm = [regex]::Match([string]$w.Url, '#(.+)$'); if ($fm.Success) { $frag = $fm.Groups[1].Value }
    if ($frag -and -not (Get-YouTubeInfo $pageUrl)) {
        $sec = Get-AnchorText $page.webSocketDebuggerUrl $frag
        if ($sec) { return @{ state = 'found'; text = $sec; source = 'page section' } }
    }
    # otherwise wait for the page to finish (a page that stays busy is read after a few seconds anyway)
    if ($ready -ne 'complete' -and ((Get-Date) - $w.Started).TotalSeconds -lt 8) { return @{ state = 'loading' } }
    $res = Invoke-Cdp $page.webSocketDebuggerUrl 'document.documentElement.outerHTML.slice(0,600000)'
    $html = ''
    if ($res -and $res.result -and $res.result.result) { $html = [string]$res.result.result.value }
    if (-not $html) { return @{ state = 'loading' } }
    # the same check as for the first lookup: an error, CAPTCHA or bot check page is never read for text
    $why = Test-ErrorOrBlockPage $html $true
    if ($why -eq 'the page is empty') { return @{ state = 'loading' } }
    if ($why -and $why -ne 'the page has almost no content') {
        # a robot check or an access-denied page can still be passed in the window; a not-found or error page of the site cannot
        if ($why -match 'CAPTCHA|denied') { return @{ state = 'blocked' } }
        $nm = ''; try { $nm = Get-SiteName $html $pageUrl } catch { }
        return @{ state = 'broken'; text = $nm; source = 'site name' }
    }
    if ($why) { return @{ state = 'weak' } }
    $s = Get-PageTextSuggestion $pageUrl $html
    if ($s.Text) { return @{ state = 'found'; text = $s.Text; source = $s.Source } }
    # nothing in the title, heading tag or description: read what the page shows (main heading, then first real sentence)
    if (-not (Get-YouTubeInfo $pageUrl)) {
        $seen = Get-VisiblePageText $page.webSocketDebuggerUrl
        if ($seen) { return @{ state = 'found'; text = $seen; source = 'page text' } }
    }
    return @{ state = 'weak' }
}

# closes the tab of one link (the helper window stays for the next Go to Link)
function Stop-Watch($w) {
    if (-not $w) { return }
    try {
        if ($w.TargetId -and $w.Port -and $w.Proc -and -not $w.Proc.HasExited) {
            $wc = New-Object Net.WebClient; $wc.Proxy = $null
            $null = $wc.DownloadString("http://127.0.0.1:$($w.Port)/json/close/$($w.TargetId)")
        }
    } catch { }
}

# ends the helper browser and removes its temporary profile (when the tool is finished)
function Stop-Shared {
    $w = $script:Shared; if (-not $w) { return }
    try {
        if ($w.Proc -and -not $w.Proc.HasExited) {
            if ($env:OS -eq 'Windows_NT') { & taskkill.exe /PID $w.Proc.Id /T /F 2>&1 | Out-Null } else { $w.Proc.Kill() }
        }
    } catch { }
    try { Start-Sleep -Milliseconds 300; Remove-Item -LiteralPath $w.Profile -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    $script:Shared = $null
}

# Decides the status of a web address and the text that goes with it. Status is
#   working  - the page works and a meaningful sentence was found
#   notext   - the page works but no meaningful sentence was found
#   broken   - the link does not work, or is blocked and could not be verified
# The destination is checked FIRST; only a working destination is ever read for text. For a broken link and
# for a page with nothing meaningful the box is left empty - nothing is made up from the web address.
function Get-LinkSuggestion([string]$url) {
    $isPdf = ($url -match '(?i)\.pdf($|[?#])')
    $text = ''; $source = ''; $status = 'working'; $reason = ''; $done = $false
    if ($env:LF_NOLOOKUP) { $status = 'notext'; $source = 'lookup is switched off'; $done = $true }
    if (-not $done) {
        $yt = Get-YouTubeInfo $url
        if ($yt) {
            $o = Get-YouTubeOEmbed $yt
            if ($o.Status -eq 'ok') {
                $s = Get-YouTubeSentence $yt $o.Title
                if ($s) { $text = $s; $source = 'YouTube video title'; $done = $true }
            } elseif ($o.Status -eq 'notfound') { $status = 'broken'; $reason = 'the video was not found'; $done = $true }
        }
    }
    if (-not $done) {
        # 1. Is the destination working? Decided before anything is read from the page.
        $st = Get-LinkStatus $url
        $isPdf = $st.IsPdf
        if (-not $st.Working) {
            $status = 'broken'; $reason = $st.Reason
            # a not-found or error page the site itself shows: the name of the site is the text (never the page title of a CAPTCHA or access-denied page)
            if ($st.ErrHtml -and $st.ErrKind -and $st.ErrKind -notmatch 'CAPTCHA|denied|empty|almost no content') {
                $nm = ''; try { $nm = Get-SiteName $st.ErrHtml $url } catch { }
                if ($nm) { $text = $nm; $source = 'site name' }
            }
        }
        else {
            # 2. Only now may the page itself be used.
            $source = $st.Reason
            try {
                if ($st.EdgeHtml) {
                    # a script was blocked, but a real browser opens the page: it works
                    $ps2 = Get-PageTextSuggestion $url $st.EdgeHtml
                    if ($ps2.Text) { $text = $ps2.Text; $source = $ps2.Source + ' (fetched through Edge)' }
                } elseif ($isPdf) {
                    $pt = Get-PdfTitle $st.Bytes
                    if ($pt) { $text = Clean-LinkText (Clean-Title $pt); $source = 'PDF title' }
                } else {
                    $html = Decode-Html $st.Bytes $st.ContentType
                    $ps = Get-PageTextSuggestion $url $html
                    $text = $ps.Text; $source = $ps.Source
                }
            } catch { $text = ''; $source = '' }
            if (-not $text) { $status = 'notext'; $source = '' }
        }
    }
    if ($status -ne 'working' -and $source -ne 'site name') { $text = '' }
    if ($text -and $isPdf -and $text -notmatch '(?i)\.pdf$') { $text = $text + '.pdf' }
    return [pscustomobject]@{ Text = $text; Source = $source; IsPdf = $isPdf; Status = $status; Reason = $reason }
}

function Html([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }
$jsq = { param($s) (ConvertTo-Json -InputObject ([string]$s) -Compress).Replace('<', ([string][char]92 + 'u003c')) }

# ---------- review page ----------
function Build-Page($items, $fineCount, $fileNames, [int]$otherCount = 0) {
    $itemJson = New-Object System.Collections.Generic.List[string]
    foreach ($it in $items) {
        $wh = ($it.Where | Select-Object -First 6 | ForEach-Object { & $jsq $_ }) -join ','
        $itemJson.Add('{"key":' + (& $jsq $it.Key) + ',"url":' + (& $jsq $it.Url) + ',"count":' + $it.Count + ',"where":[' + $wh + '],"extra":' + [Math]::Max(0, $it.Where.Count - 6) + ',"text":' + (& $jsq $it.Text) + ',"source":' + (& $jsq $it.Source) + ',"status":' + (& $jsq $it.Status) + ',"reason":' + (& $jsq $it.Reason) + '}')
    }
    $fn = (($fileNames | ForEach-Object { Html $_ }) -join ', ')
    $fineText = [string]$items.Count + ' of ' + [string]($items.Count + $fineCount + $otherCount) + ' hyperlinks need new text'
    if ($fineCount -gt 0) { $fineText += ', ' + [string]$fineCount + ' already read well' }
    if ($otherCount -gt 0) { $fineText += ', ' + [string]$otherCount + ' jump inside the document or are not web links (left alone)' }
    $html = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Link text review</title>
<style>
:root{--bg:#f4f5f7;--card:#fff;--text:#1b1f23;--muted:#59636e;--line:#d0d7de;--accent:#0b5cad;--accent-bg:#ddeeff;--ok:#1a7f37;--ok-bg:#e6f6ea;--warn:#8a5a00;--warn-bg:#fff3d1;--err:#a40e26;--soft:#eef1f4}
@media (prefers-color-scheme: dark){:root{--bg:#161b22;--card:#1f2630;--text:#e6edf3;--muted:#9aa4af;--line:#38414d;--accent:#58a6ff;--accent-bg:#14304f;--ok:#56d364;--ok-bg:#12301b;--warn:#e3b341;--warn-bg:#3a2d0b;--err:#ff7b72;--soft:#262e3a}}
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--text);padding-bottom:96px}
.wrap{max-width:1000px;margin:0 auto;padding:0 16px}
header{position:sticky;top:0;z-index:5;background:var(--bg);padding:14px 0 10px;border-bottom:1px solid var(--line)}
.top{display:flex;gap:16px;align-items:center;flex-wrap:wrap}.top h1{margin:0;font-size:18px;font-weight:600}.sub{font-size:13px;color:var(--muted)}
.grow{flex:1}.prog{width:250px}.prog .t{display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px}
.track{height:8px;border-radius:4px;background:var(--line)}.fillbar{height:8px;border-radius:4px;background:var(--ok);width:0;transition:width .25s}
.tools{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap}
.chip{border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:999px;padding:5px 12px;font-size:13px;cursor:pointer}
.chip.on{background:var(--accent-bg);color:var(--accent);border-color:var(--accent)}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;margin-top:12px}
.card.ok{border-color:var(--ok)}
.url{font-family:Consolas,monospace;font-size:13px;word-break:break-all}
.top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px}
.top .url{flex:1;padding-top:6px}
.wk{background:var(--ok-bg);color:var(--ok);border-radius:999px;padding:3px 12px;font-size:12px;font-weight:600;white-space:nowrap;margin-top:4px}
a.golink{display:inline-flex;align-items:center;gap:6px;white-space:nowrap;background:#1a7f37;color:#fff;border-radius:999px;padding:7px 16px;font-size:13px;font-weight:600;text-decoration:none}
a.golink:hover{background:#17692f}
.meta{font-size:12px;color:var(--muted);margin-top:4px}
.note{font-size:13px;color:var(--muted);margin-top:8px}
.sno{flex:none;min-width:30px;height:30px;padding:0 8px;border-radius:15px;background:var(--accent-bg);color:var(--accent);font-weight:700;font-size:14px;display:inline-flex;align-items:center;justify-content:center}.top .url{padding-top:5px}.cp{display:inline-flex;align-items:center;gap:4px;vertical-align:middle;margin-left:8px;padding:3px 7px;border:1px solid var(--line);border-radius:6px;background:transparent;color:var(--muted);font:inherit;font-size:12px;cursor:pointer}.cp:hover{color:var(--accent);border-color:var(--accent)}.cp.done{color:var(--ok);border-color:var(--ok)}.note a{color:var(--accent);font-weight:600}.sameRow{margin:6px 0 2px;display:flex;flex-wrap:wrap;gap:6px;align-items:center}.sameLbl{font-size:12px;color:var(--muted)}.same{font:inherit;font-size:12px;font-weight:600;padding:2px 10px;border-radius:999px;border:1px solid var(--accent);background:var(--accent-bg);color:var(--accent);cursor:pointer}.same:hover{filter:brightness(1.15)}.card.flash{outline:3px solid var(--accent);outline-offset:2px}
.note.red{color:var(--err)}.note.grn{color:var(--ok)}.note.gry{color:var(--muted)}
.lbl{font-size:13px;font-weight:600;margin-top:12px}
input.txt{width:100%;padding:9px 10px;font-size:14px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-family:inherit;margin-top:4px}
input.txt.empty{border-color:var(--err)}
.empty-msg{padding:40px;text-align:center;color:var(--muted)}
#bar{position:fixed;left:0;right:0;bottom:0;background:var(--card);border-top:1px solid var(--line);padding:12px 16px;z-index:9}
#bar .in{max-width:1000px;margin:0 auto;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
#msg{font-size:13px;color:var(--muted);flex:1}
button.btn{font:inherit;font-size:14px;padding:9px 16px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--text);cursor:pointer}
button.main-btn{background:var(--text);color:var(--card);border-color:var(--text)}button:disabled{opacity:.5;cursor:default}
#panel{position:fixed;left:16px;right:16px;bottom:84px;max-width:968px;margin:0 auto;max-height:60vh;overflow:auto;background:var(--card);border:1px solid var(--ok);border-radius:12px;padding:16px 18px;z-index:10;box-shadow:0 6px 24px rgba(0,0,0,.25)}
#panel.err{border-color:var(--err)}#panel h2{font-size:16px;margin:0 0 8px}#panel ul{margin:4px 0 10px 18px;padding:0;font-size:13px}#panel .row{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap}
</style></head><body>
<header><div class="wrap">
<div class="top"><div class="grow"><h1>Link text review</h1><div class="sub">__FILES__ &middot; __FINE__</div></div>
<div class="prog"><div class="t"><b id="cnt"></b><span id="pct"></span></div><div class="track"><div class="fillbar" id="fill"></div></div></div></div>
<div class="tools">
<button class="chip on" data-f="all">All <span id="n-all"></span></button>
<button class="chip" data-f="need">Needs text <span id="n-need"></span></button>
<button class="chip" data-f="ready">Ready <span id="n-ready"></span></button>
</div></div></header>
<div class="wrap" id="list"></div>
<div id="panel" hidden></div>
<div id="bar"><div class="in"><span id="msg">Edit the text in each box.</span><button class="btn main-btn" id="apply">Apply to the files</button></div></div>
<script>
var ITEMS = __ITEMS__;
var TOKEN = "__TOKEN__";
var CANWATCH = __CANWATCH__;
var S = ITEMS.map(function(it){ return {text: it.text}; });
var W = ITEMS.map(function(){ return {note: "", busy: false, weak: 0, found: "", kept: "", stuck: 0, brokenPage: false}; });
var GOICON = '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 2h5v5M14 2L7.5 8.5M12 9.5V13a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h3.5"/></svg>';
var COPYICON = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" fill="currentColor"><path d="M0 6.75C0 5.78.78 5 1.75 5h1.5a.75.75 0 0 1 0 1.5h-1.5a.25.25 0 0 0-.25.25v7.5c0 .14.11.25.25.25h7.5a.25.25 0 0 0 .25-.25v-1.5a.75.75 0 0 1 1.5 0v1.5A1.75 1.75 0 0 1 9.25 16h-7.5A1.75 1.75 0 0 1 0 14.25Z"/><path d="M5 1.75C5 .78 5.78 0 6.75 0h7.5C15.22 0 16 .78 16 1.75v7.5A1.75 1.75 0 0 1 14.25 11h-7.5A1.75 1.75 0 0 1 5 9.25Zm1.75-.25a.25.25 0 0 0-.25.25v7.5c0 .14.11.25.25.25h7.5a.25.25 0 0 0 .25-.25v-7.5a.25.25 0 0 0-.25-.25Z"/></svg>';
var TICKICON = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" fill="currentColor"><path d="M13.78 4.22a.75.75 0 0 1 0 1.06l-7.25 7.25a.75.75 0 0 1-1.06 0L2.22 9.28a.75.75 0 0 1 1.06-1.06L6 10.94l6.72-6.72a.75.75 0 0 1 1.06 0Z"/></svg>';
function copyText(t){
  if(navigator.clipboard && window.isSecureContext){ return navigator.clipboard.writeText(t); }
  return new Promise(function(ok, no){ try{ var a = document.createElement("textarea"); a.value = t; a.style.position = "fixed"; a.style.opacity = "0"; document.body.appendChild(a); a.select(); var r = document.execCommand("copy"); document.body.removeChild(a); r ? ok() : no(); }catch(e){ no(e); } });
}
document.addEventListener("click", function(e){
  var b = e.target.closest ? e.target.closest("[data-cp]") : null; if(!b) return;
  copyText(ITEMS[+b.getAttribute("data-cp")].url).then(function(){
    b.classList.add("done"); b.innerHTML = TICKICON + '<span>Copied</span>';
    setTimeout(function(){ b.classList.remove("done"); b.innerHTML = COPYICON; }, 1500);
  }).catch(function(){ b.title = "Could not copy, select the address and press Ctrl+C"; });
});
var filt = "all";
function esc(s){ return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function stateOf(i){ return S[i].text.trim() ? "ready" : "need"; }
function cap(t){ return t.charAt(0).toUpperCase() + t.slice(1); }
function noteInfo(i){
  var it = ITEMS[i], w = W[i], has = S[i].text.trim() !== "";
  if(w.busy) return {t: w.note, c: "gry", a: w.stuck >= 12};
  if(has){
    var nameUsed = (it.source === "site name" && S[i].text === it.text) || (w.brokenPage && w.found && S[i].text === w.found);
    if(w.kept) return {t: w.kept, c: "grn"};
    if(nameUsed) return {t: "Link not working, the site name is used, please check", c: "grn"};
    if(w.found && S[i].text === w.found) return {t: "Text taken from the page, please check", c: "grn"};
    if(it.status === "working") return {t: /please check/.test(it.source) ? cap(it.source) : "", c: "grn"};
    return {t: "Text added", c: "grn"};
  }
  if(w.note) return {t: w.note, c: "red"};
  if(it.status === "broken") return {t: "Link not working, please type the text", c: "red"};
  if(it.status === "working") return {t: "Type the text for the link", c: "red"};
  return {t: "No meaningful sentence found, please type the text", c: "red"};
}
// the link to open the address in the person's own browser is always there on a card that has Go to Link
function noteHtml(x, i){
  var open = ITEMS[i].status === "working" ? "" : '<a class="own" href="' + esc(ITEMS[i].url) + '" target="_blank" rel="noopener noreferrer">Open in my normal browser</a>';
  return esc(x.t) + (x.t && open ? " &middot; " : "") + open;
}
function refreshNote(i){ var n = document.querySelector('[data-note="' + i + '"]'); if(!n) return; var x = noteInfo(i); n.innerHTML = noteHtml(x, i); n.className = "note " + x.c; }
function setNote(i, t){ W[i].note = t; refreshNote(i); }
function sameTag(i){
  var o = []; ITEMS.forEach(function(x, k){ if(k !== i && x.url === ITEMS[i].url) o.push(k); });
  if(!o.length) return "";
  return '<div class="sameRow"><span class="sameLbl">Same address as</span>' + o.map(function(k){
    return '<button type="button" class="same" data-go="' + k + '" title="Go to this link">S.No ' + (k + 1) + '</button>';
  }).join("") + '</div>';
}
document.addEventListener("click", function(e){
  var b = e.target.closest ? e.target.closest("[data-go]") : null; if(!b) return;
  var c = document.getElementById("card-" + b.getAttribute("data-go"));
  if(!c){ filt = "all"; render(); c = document.getElementById("card-" + b.getAttribute("data-go")); }
  if(c){ c.scrollIntoView({behavior: "smooth", block: "center"}); c.classList.add("flash"); setTimeout(function(){ c.classList.remove("flash"); }, 1800); }
});
function render(){
  var c = {need: 0, ready: 0}, h = "";
  ITEMS.forEach(function(it, i){ c[stateOf(i)]++; });
  ITEMS.forEach(function(it, i){
    var st = stateOf(i);
    if(filt !== "all" && filt !== st) return;
    var right = it.status === "working" ? '<span class="wk">Working</span>' : '<a class="golink" data-i="' + i + '" href="' + esc(it.url) + '" target="_blank" rel="noopener noreferrer">' + GOICON + 'Go to Link</a>';
    h += '<section class="card' + (st === "ready" ? " ok" : "") + '" id="card-' + i + '"><div class="top"><span class="sno" title="Serial number">' + (i + 1) + '</span><span class="url">' + esc(it.url) + '<button type="button" class="cp" data-cp="' + i + '" title="Copy the link" aria-label="Copy the link">' + COPYICON + '</button></span>' + right + '</div>' +
      (it.where.join("") ? '<div class="meta">' + it.where.map(esc).join(", ") + '</div>' : '') + sameTag(i) +
      '<div class="lbl">New text for the link</div><input class="txt' + (!S[i].text.trim() ? ' empty' : '') + '" data-i="' + i + '" type="text" value="' + esc(S[i].text) + '" placeholder="Type what the link is about" aria-label="New link text">' +
      '<div class="note ' + noteInfo(i).c + '" data-note="' + i + '" title="' + esc(it.reason || "") + '">' + noteHtml(noteInfo(i), i) + '</div></section>';
  });
  document.getElementById("list").innerHTML = h || '<div class="empty-msg">Nothing here. Try another filter.</div>';
  counts();
}
function counts(){
  var c = {need: 0, ready: 0};
  ITEMS.forEach(function(it, i){ c[stateOf(i)]++; });
  var total = ITEMS.length, done = c.ready, pc = total ? Math.round(done / total * 100) : 100;
  document.getElementById("cnt").textContent = done + " of " + total + " links have text";
  document.getElementById("pct").textContent = pc + "%"; document.getElementById("fill").style.width = pc + "%";
  document.getElementById("n-all").textContent = total; document.getElementById("n-need").textContent = c.need;
  document.getElementById("n-ready").textContent = c.ready;
}document.addEventListener("input", function(e){ var i = e.target.getAttribute && e.target.getAttribute("data-i"); if(i !== null && i !== undefined){ S[+i].text = e.target.value; var r = e.target.classList; if(S[+i].text.trim()) r.remove("empty"); else r.add("empty"); counts(); refreshNote(+i); } });
[].slice.call(document.querySelectorAll(".chip")).forEach(function(c){ c.addEventListener("click", function(){ filt = c.getAttribute("data-f"); [].slice.call(document.querySelectorAll(".chip")).forEach(function(x){ x.classList.toggle("on", x === c); }); render(); }); });
function list(items){ return items.length ? "<ul>" + items.map(function(x){ return "<li>" + esc(x) + "</li>"; }).join("") + "</ul>" : ""; }
function post(path, body){
  return fetch(path, {method: "POST", headers: {"Content-Type": "application/json", "X-Token": TOKEN}, body: JSON.stringify(body || {})})
    .then(function(r){ return r.json().then(function(j){ if(!r.ok){ throw new Error(j.error || r.statusText); } return j; }); });
}
function fillText(i, t){
  var w = W[i]; w.busy = false;
  if(!S[i].text.trim()){
    S[i].text = t; w.found = t; w.note = ""; w.kept = "";
    var inp = document.querySelector('input.txt[data-i="' + i + '"]'); if(inp){ inp.value = t; inp.classList.remove("empty"); }
    counts();
  } else if(S[i].text.trim() === t){ w.found = t; w.note = ""; w.kept = ""; }
  else { w.kept = "A suggestion was found, your own text was kept: " + t; }
  refreshNote(i);
}
function poll(i, n){
  var w = W[i];
  if(n > 400){ w.busy = false; setNote(i, "Stopped waiting for the page. Press Go to Link to try again."); return; }
  post("/pagestatus", {i: i}).then(function(r){
    if(r.state === "found"){ fillText(i, r.text); return; }
    if(r.state === "broken"){
      w.brokenPage = true;
      if(r.text){ fillText(i, r.text); } else { w.busy = false; setNote(i, "Link not working, please type the text"); }
      return;
    }
    if(r.state === "weak"){ w.weak++; if(w.weak >= 4){ w.busy = false; setNote(i, "The page opened but no meaningful sentence was found, please type the text"); return; } setNote(i, "Reading the page..."); }
    else if(r.state === "closed"){ w.busy = false; setNote(i, "The window was closed before the page was read, please type the text"); return; }
    else if(r.state === "error"){ w.busy = false; setNote(i, "The page could not be read, please type the text"); return; }
    else if(r.state === "blocked"){
      w.stuck++;
      setNote(i, w.stuck >= 12 ? "The site keeps showing its check in this window. Open the link in your normal browser, then type the text here (use the copy button for the address)." : "The page shows a check or an error. Finish it in the window and the text is read when the page has loaded.");
      setTimeout(function(){ poll(i, n + 1); }, 2000); return;
    }
    else { setNote(i, "Waiting for the page to load..."); }
    setTimeout(function(){ poll(i, n + 1); }, 500);
  }).catch(function(){ w.busy = false; setNote(i, "Lost contact with the helper window, please type the text"); });
}
function startWatch(i){
  var w = W[i]; if(w.busy) return;
  w.busy = true; w.weak = 0; w.stuck = 0; w.kept = ""; w.brokenPage = false; setNote(i, "Opening the page in a new window...");
  post("/golink", {i: i}).then(function(r){
    if(!r.ok){ w.busy = false; setNote(i, "Opened as a normal link. The page cannot be read in this case, please type the text"); window.open(ITEMS[i].url, "_blank", "noopener"); return; }
    poll(i, 0);
  }).catch(function(){ w.busy = false; setNote(i, "Could not open the page, please type the text"); });
}
document.addEventListener("click", function(e){
  var a = e.target.closest ? e.target.closest("a.golink") : null;
  if(!a || !CANWATCH || !TOKEN || location.protocol === "file:") return;
  e.preventDefault(); startWatch(+a.getAttribute("data-i"));
});
function showPanel(html, isErr){ var p = document.getElementById("panel"); p.hidden = false; p.className = isErr ? "err" : ""; p.innerHTML = html; }
document.getElementById("apply").addEventListener("click", function(){
  if(location.protocol === "file:" || !TOKEN){ showPanel("<h2>This page needs the CMD window</h2><p>This is only a saved copy. Run Link Fixer.cmd and use the page it opens (its address starts with http://localhost).</p>", true); return; }
  var picked = ITEMS.filter(function(it, i){ return S[i].text.trim(); }).length;
  if(!picked){ showPanel("<h2>Nothing to apply</h2><p>Type the text for at least one link first.</p>", true); return; }
  if(!confirm(picked + " links will get their new text in a copy of each file. The originals are not changed.\n\nContinue?")) return;
  var btn = document.getElementById("apply"); btn.disabled = true; document.getElementById("msg").textContent = "Writing...";
  var items = ITEMS.map(function(it, i){ return {key: it.key, url: it.url, text: S[i].text.trim()}; });
  post("/apply", {items: items}).then(function(r){
    document.getElementById("msg").textContent = "Done";
    showPanel("<h2>Saved " + r.Files.length + " file" + (r.Files.length === 1 ? "" : "s") + "</h2>" +
      "<ul>" + r.Files.map(function(f){ return "<li>" + esc(f.OutFile) + ": " + f.Changed + " links changed" + (f.Skipped ? ", " + f.Skipped + " left as they are" : "") + "</li>"; }).join("") + "</ul>" +
      (r.Notes.length ? "<p>Notes:</p>" + list(r.Notes) : "") +
      '<div class="row"><button class="btn main-btn" id="openf">Open the folder</button><button class="btn" id="finish">Finish and close the helper</button><button class="btn" id="hide">Hide this</button></div>', false);
    document.getElementById("openf").onclick = function(){ post("/open"); };
    document.getElementById("hide").onclick = function(){ document.getElementById("panel").hidden = true; };
    document.getElementById("finish").onclick = function(){ post("/quit").then(function(){ document.getElementById("msg").textContent = "The helper window has stopped. You can close this page."; document.getElementById("apply").disabled = true; }); };
  }).catch(function(e){ document.getElementById("msg").textContent = ""; showPanel("<h2>It did not work</h2><p>" + esc(e.message) + "</p><p>If the CMD window was closed, run Link Fixer.cmd again. If a file or its copy is open in Word or PowerPoint, close it and try again.</p>", true); })
  .then(function(){ btn.disabled = false; });
});
render();
</script></body></html>
'@
    return $html.Replace('__ITEMS__', ('[' + ($itemJson -join ',') + ']')).Replace('__FILES__', $fn).Replace('__FINE__', $fineText).Replace('__CANWATCH__', $(if (Get-BrowserPath) { 'true' } else { 'false' }))
}

# ---------- helper server ----------
function Start-LinkServer([string]$html, [string]$token, $files, $urls) {
    $listener = $null; $port = 0
    foreach ($try in 1..40) {
        $port = Get-Random -Minimum 49200 -Maximum 49900
        $l = New-Object System.Net.HttpListener
        $l.Prefixes.Add("http://localhost:$port/")
        try { $l.Start(); $listener = $l; break } catch { $l.Close() }
    }
    if (-not $listener) { throw 'Could not start the local page server.' }
    $url = "http://localhost:$port/"
    if ($env:LF_PORTFILE) { Set-Content -LiteralPath $env:LF_PORTFILE -Value "$port|$token" -Encoding ASCII } else { Start-Process $url }
    Write-Host ''
    Write-Host ('The review page is open at ' + $url)
    Write-Host 'Leave this window open while you use the page. Close it when you are finished.'
    $quit = $false
    while ($listener.IsListening -and -not $quit) {
        $ctx = $listener.GetContext(); $req = $ctx.Request; $res = $ctx.Response
        $status = 200; $ctype = 'application/json; charset=utf-8'; $body = ''
        try {
            $path = $req.Url.AbsolutePath
            if ($req.HttpMethod -eq 'GET' -and $path -eq '/') { $ctype = 'text/html; charset=utf-8'; $body = $html }
            elseif ($req.HttpMethod -eq 'POST') {
                if ($req.Headers['X-Token'] -ne $token) { $status = 403; $body = '{"error":"Not allowed"}' }
                else {
                    $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8); $raw = $sr.ReadToEnd(); $sr.Close()
                    if ($path -eq '/apply') {
                        $data = $raw | ConvertFrom-Json
                        $texts = @{}
                        foreach ($it in @($data.items)) { if (([string]$it.text).Trim()) { $texts[[string]$it.key] = ([string]$it.text).Trim() } }
                        $result = Invoke-ApplyAll $files $texts
                        Write-Host ('Applied: ' + (($result.Files | ForEach-Object { $_.OutFile + ' (' + $_.Changed + ')' }) -join ', '))
                        $script:LastOut = $result.Files[0].OutPath
                        $body = $result | ConvertTo-Json -Depth 5 -Compress
                    }
                    elseif ($path -eq '/open') {
                        if ($script:LastOut -and (Test-Path -LiteralPath $script:LastOut)) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $script:LastOut + '"') }
                        $body = '{"ok":true}'
                    }
                    elseif ($path -eq '/golink') {
                        # open the address in the helper's own browser window so the page can be read once it has loaded
                        $i = [int]($raw | ConvertFrom-Json).i
                        if ($i -lt 0 -or $i -ge @($urls).Count) { $status = 404; $body = '{"error":"Not found"}' }
                        else {
                            Stop-Watch $script:Watch[$i]
                            $w = Start-WatchBrowser ([string]@($urls)[$i])
                            if ($w) { $script:Watch[$i] = $w; $body = '{"ok":true}' } else { $body = '{"ok":false}' }
                        }
                    }
                    elseif ($path -eq '/pagestatus') {
                        $i = [int]($raw | ConvertFrom-Json).i
                        $w = $null; if ($script:Watch.ContainsKey($i)) { $w = $script:Watch[$i] }
                        if ($w) { $body = (Get-WatchState $w) | ConvertTo-Json -Compress } else { $body = '{"state":"closed"}' }
                    }
                    elseif ($path -eq '/quit') { $body = '{"ok":true}'; $quit = $true }
                    else { $status = 404; $body = '{"error":"Not found"}' }
                }
            } else { $status = 404; $body = '{"error":"Not found"}' }
        } catch { $status = 500; $body = (@{ error = $_.Exception.Message } | ConvertTo-Json -Compress) }
        $bytes = [Text.Encoding]::UTF8.GetBytes($body)
        $res.StatusCode = $status; $res.ContentType = $ctype; $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes, 0, $bytes.Length); $res.OutputStream.Close()
    }
    $listener.Stop()
    foreach ($k in @($script:Watch.Keys)) { Stop-Watch $script:Watch[$k] }
    Stop-Shared
    Write-Host 'Page closed.'
}

function Invoke-ApplyAll($files, $texts) {
    $out = New-Object System.Collections.Generic.List[object]
    $notes = New-Object System.Collections.Generic.List[string]
    $fi = -1
    foreach ($f in $files) {
        $fi++
        $ext = [IO.Path]::GetExtension($f)
        $outPath = Join-Path (Get-FolderOf $f) ([IO.Path]::GetFileNameWithoutExtension($f) + ' Links' + $ext)
        $r = Scan-File $f $true $texts $outPath $fi
        $out.Add([pscustomobject]@{ InFile = (Split-Path -Leaf $f); OutFile = (Split-Path -Leaf $outPath); OutPath = $outPath; Changed = $r.Changed; Skipped = $r.Skipped })
        Add-WorkLog 'Link Fixer' $f $outPath
    }
    $report = New-Object System.Collections.Generic.List[string]
    foreach ($o in $out) { $report.Add("$($o.InFile) -> $($o.OutFile): $($o.Changed) links changed, $($o.Skipped) left as they are") }
    $notes.Add('Links inside a document (bookmarks) and mail links are not changed.')
    return [pscustomobject]@{ Files = $out.ToArray(); Notes = $notes.ToArray() }
}

# =====================================================================
# main
# =====================================================================
try {
    $files = @()
    if ($env:LF_FILES) { $files = @($env:LF_FILES.Split(';') | Where-Object { $_ }) } else { $files = Pick-Files }
    # A file given just by name, with no folder in front of it, leaves
    # Split-Path -Parent with nothing to return later on - resolve every
    # file to its full path up front so that never happens.
    $files = @($files | ForEach-Object { [IO.Path]::GetFullPath($_) })
    if ($files.Count -eq 0) { Write-Host 'No file selected.' }
    else {
        Write-Host 'Reading the links...'
        $fine = 0; $other = 0; $webTotal = 0; $list = New-Object System.Collections.Generic.List[object]
        $fi = -1
        foreach ($f in $files) {
            $fi++
            $name = Split-Path -Leaf $f
            $r = Scan-File $f $false @{} $null $fi
            $other += $r.Other
            foreach ($h in $r.Hits) {
                $webTotal++
                if (-not $h.Fix) { $fine++; continue }
                # every link gets its own card, also when the same address is used again
                $list.Add([pscustomobject]@{ Key = $h.Key; Url = $h.Url; Where = $(& { $loc = $h.Where; if ($loc -match '^paragraph \d+$') { $loc = '' }; $parts = @(); if ($files.Count -gt 1) { $parts += $name }; if ($loc) { $parts += $loc }; ($parts -join ' ') }); Text = $h.Text })
            }
        }
        Write-Host ("Hyperlinks found: " + ($webTotal + $other) + " (web links: $webTotal, links inside the document or other kinds: $other)")
        Write-Host ("Web links that need new text: " + $list.Count + "; web links that already read well: " + $fine)
        if ($list.Count -eq 0) { Write-Host 'Nothing to fix in these files.' }
        else {
            $items = New-Object System.Collections.Generic.List[object]
            $n = 0; $cache = @{}
            foreach ($e in $list) {
                $n++
                Write-Host ("Looking up $n of " + $list.Count + ': ' + $e.Url)
                # the same address is looked up once, the answer is used for each of its links
                if (-not $cache.ContainsKey($e.Url)) { $cache[$e.Url] = Get-LinkSuggestion $e.Url }
                $s = $cache[$e.Url]
                $items.Add([pscustomobject]@{ Key = $e.Key; Url = $e.Url; Count = 1; Where = @($e.Where); Texts = @($e.Text); Text = $s.Text; Source = $s.Source; Status = $s.Status; Reason = $s.Reason })
            }
            $names = @($files | ForEach-Object { Split-Path -Leaf $_ })
            $html = Build-Page $items.ToArray() $fine $names $other
            $backup = Join-Path (Get-FolderOf $files[0]) 'Link Fixer review.html'
            [IO.File]::WriteAllText($backup, $html.Replace('__TOKEN__', ''), (New-Object Text.UTF8Encoding($false)))
            Write-Host ('Review page saved as a backup copy: ' + $backup)
            if (-not $env:LF_NOSERVE) {
                $token = [guid]::NewGuid().ToString('N')
                Start-LinkServer $html.Replace('__TOKEN__', $token) $token $files @($items | ForEach-Object { $_.Url })
            }
        }
    }
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ('  (version ' + $LF_VERSION + ', at line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + ($_.InvocationInfo.Line + '').Trim() + ')') -ForegroundColor DarkGray
}
