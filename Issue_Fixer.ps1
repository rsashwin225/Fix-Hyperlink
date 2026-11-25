# Issue Fixer - reads the pre-check / tester issue sheet (.xlsx) and the PowerPoint deck it was made for, fixes what can be fixed
# safely (reading order, language tags, colour contrast), says which findings look like false positives, and fills the sheet's
# Status / Owner / Comment columns. The text and objects of the deck are NEVER changed or added to.
# S4C ISSUE FIXER SCRIPT
$IF_VERSION = '1.0'   # shown at the start of the window; the updater compares this number with GitHub
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$ci = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object Text.UTF8Encoding($false)

function Add-WorkLog([string]$tool, [string]$inPath, [string]$outPath) {
    try {
        $count = ''; $unit = ''
        $z = [IO.Compression.ZipFile]::OpenRead($outPath)
        try {
            $unit = 'slides'
            $e = $z.GetEntry('ppt/presentation.xml')
            if ($e) { $sr = New-Object IO.StreamReader($e.Open()); $t = $sr.ReadToEnd(); $sr.Close(); $count = [regex]::Matches($t, '<p:sldId\s').Count }
        } finally { $z.Dispose() }
        $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
        $dir = Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog'
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $now = Get-Date
        $line = @($now.ToString('yyyy-MM-dd', $ci), $now.ToString('HH:mm', $ci), $env:USERNAME, $tool, [IO.Path]::GetFullPath($inPath), [IO.Path]::GetFullPath($outPath), $count, $unit) -join "`t"
        [IO.File]::AppendAllText((Join-Path $dir 'worklog.tsv'), $line + "`r`n", $utf8)
    } catch { }
}

# ------------------------------------------------------------------ small zip / xml helpers
function Read-Entry($zip, [string]$name) {
    $e = $zip.GetEntry($name); if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), $utf8); try { return $sr.ReadToEnd() } finally { $sr.Close() }
}
function Write-Entry($zip, [string]$name, [string]$text) {
    $e = $zip.GetEntry($name); if ($e) { $e.Delete() }
    $ne = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
    $st = $ne.Open(); try { $b = $utf8.GetBytes($text); $st.Write($b, 0, $b.Length) } finally { $st.Dispose() }
}
function New-Xml([string]$text) { $x = New-Object Xml.XmlDocument; $x.PreserveWhitespace = $true; $x.LoadXml($text); return $x }
function Save-Xml($x) {
    # written without the declaration, then the standard UTF-8 declaration is put in front (the default writer would say utf-16)
    $ms = New-Object IO.MemoryStream
    $set = New-Object Xml.XmlWriterSettings; $set.Encoding = $utf8; $set.OmitXmlDeclaration = $true; $set.Indent = $false
    $w = [Xml.XmlWriter]::Create($ms, $set); $x.Save($w); $w.Close()
    $body = $utf8.GetString($ms.ToArray()).TrimStart([char]0xFEFF)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n" + $body)
}
function Col-Index([string]$letters) { $n = 0; foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }; return $n }

# ------------------------------------------------------------------ colour maths (contrast)
function Get-Lum([double]$r, [double]$g, [double]$b) {
    $out = @()
    foreach ($v in @($r, $g, $b)) { $x = $v / 255.0; if ($x -le 0.03928) { $out += ($x / 12.92) } else { $out += [math]::Pow(($x + 0.055) / 1.055, 2.4) } }
    return 0.2126 * $out[0] + 0.7152 * $out[1] + 0.0722 * $out[2]
}
function Hex-Rgb([string]$h) { return @([Convert]::ToInt32($h.Substring(0, 2), 16), [Convert]::ToInt32($h.Substring(2, 2), 16), [Convert]::ToInt32($h.Substring(4, 2), 16)) }
function Get-Ratio([string]$h1, [string]$h2) {
    $a = Hex-Rgb $h1; $b = Hex-Rgb $h2
    $l1 = Get-Lum $a[0] $a[1] $a[2]; $l2 = Get-Lum $b[0] $b[1] $b[2]
    if ($l1 -lt $l2) { $t = $l1; $l1 = $l2; $l2 = $t }
    return ($l1 + 0.05) / ($l2 + 0.05)
}
function Rgb-Hsl([double]$r, [double]$g, [double]$b) {
    $r /= 255; $g /= 255; $b /= 255
    $mx = [math]::Max($r, [math]::Max($g, $b)); $mn = [math]::Min($r, [math]::Min($g, $b)); $l = ($mx + $mn) / 2; $h = 0; $s = 0
    if ($mx -ne $mn) {
        $d = $mx - $mn
        if ($l -gt 0.5) { $s = $d / (2 - $mx - $mn) } else { $s = $d / ($mx + $mn) }
        if ($mx -eq $r) { $h = ($g - $b) / $d; if ($g -lt $b) { $h += 6 } } elseif ($mx -eq $g) { $h = ($b - $r) / $d + 2 } else { $h = ($r - $g) / $d + 4 }
        $h /= 6
    }
    return @($h, $s, $l)
}
function Hue-Rgb([double]$p, [double]$q, [double]$t) {
    if ($t -lt 0) { $t += 1 }; if ($t -gt 1) { $t -= 1 }
    if ($t -lt 1 / 6) { return $p + ($q - $p) * 6 * $t }
    if ($t -lt 1 / 2) { return $q }
    if ($t -lt 2 / 3) { return $p + ($q - $p) * (2 / 3 - $t) * 6 }
    return $p
}
function Hsl-Hex([double]$h, [double]$s, [double]$l) {
    if ($s -eq 0) { $r = $l; $g = $l; $b = $l } else {
        if ($l -lt 0.5) { $q = $l * (1 + $s) } else { $q = $l + $s - $l * $s }
        $p = 2 * $l - $q
        $r = Hue-Rgb $p $q ($h + 1 / 3); $g = Hue-Rgb $p $q $h; $b = Hue-Rgb $p $q ($h - 1 / 3)
    }
    return ('{0:X2}{1:X2}{2:X2}' -f [int][math]::Round($r * 255), [int][math]::Round($g * 255), [int][math]::Round($b * 255))
}
# the same colour, only darker, until it reaches the ratio on that background
function Get-Darker([string]$hex, [string]$bg, [double]$target) {
    $c = Hex-Rgb $hex; $hsl = Rgb-Hsl $c[0] $c[1] $c[2]; $l = $hsl[2]; $cur = $hex
    while ((Get-Ratio $cur $bg) -lt $target -and $l -gt 0.02) { $l -= 0.005; $cur = Hsl-Hex $hsl[0] $hsl[1] $l }
    return $cur
}

function Fix-Issues {
param([string]$PptxPath, [string]$XlsxPath)

$src = (Resolve-Path -LiteralPath $PptxPath).Path
$xin = (Resolve-Path -LiteralPath $XlsxPath).Path
$dir = Split-Path -Parent $src
$base = [IO.Path]::GetFileNameWithoutExtension($src)
$outPptx = Join-Path $dir ($base + ' Fixed.pptx')
$outXlsx = Join-Path $dir ($base + ' Issue Report.xlsx')
Write-Host ('Fixing: ' + [IO.Path]::GetFileName($src))

# ---------------------------------------------------------------- 1. the issue sheet
$xz = [IO.Compression.ZipFile]::OpenRead($xin)
try {
    $wb = Read-Entry $xz 'xl/workbook.xml'; $wbr = Read-Entry $xz 'xl/_rels/workbook.xml.rels'
    $ssText = Read-Entry $xz 'xl/sharedStrings.xml'
    $sheets = @{}
    foreach ($m in [regex]::Matches($wb, '<sheet\s[^>]*?>')) { $n1 = [regex]::Match($m.Value, '\sname="([^"]+)"').Groups[1].Value; $r1 = [regex]::Match($m.Value, 'r:id="([^"]+)"').Groups[1].Value; if ($n1) { $sheets[$n1] = $r1 } }
    $targetOf = @{}
    foreach ($m in [regex]::Matches($wbr, '<Relationship\s[^>]*?>')) { $i1 = [regex]::Match($m.Value, '\sId="([^"]+)"').Groups[1].Value; $t1 = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value; if ($i1) { $targetOf[$i1] = $t1 } }
    $findName = $null
    foreach ($k in $sheets.Keys) { if ($k -match '(?i)find') { $findName = $k } }
    if (-not $findName) { throw 'The issue sheet has no "Findings" sheet.' }
    $findPath = 'xl/' + ($targetOf[$sheets[$findName]] -replace '^/?(xl/)?', '')
    $findXml = Read-Entry $xz $findPath
} finally { $xz.Dispose() }

$shared = @()
if ($ssText) { foreach ($si in ([xml]$ssText).GetElementsByTagName('si')) { $shared += $si.InnerText } }
$fx = New-Xml $findXml
$nsm = New-Object Xml.XmlNamespaceManager($fx.NameTable); $nsm.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
$rowsEl = @($fx.SelectNodes('//m:sheetData/m:row', $nsm))
function Cell-Text($c) {
    if ($c.GetAttribute('t') -eq 's') { $v = $c.SelectSingleNode('m:v', $nsm); if ($v) { return $shared[[int]$v.InnerText] } }
    $is = $c.SelectSingleNode('m:is', $nsm); if ($is) { return $is.InnerText }
    $v = $c.SelectSingleNode('m:v', $nsm); if ($v) { return $v.InnerText }
    return ''
}
$hdr = @{}
foreach ($c in $rowsEl[0].SelectNodes('m:c', $nsm)) { $col = ($c.GetAttribute('r') -replace '\d', ''); $hdr[(Cell-Text $c).Trim().ToLowerInvariant()] = $col }
function HCol([string]$like) { foreach ($k in $hdr.Keys) { if ($k -like $like) { return $hdr[$k] } }; return $null }
$cLoc = HCol 'location*'; $cRule = HCol 'rule*'; $cCond = HCol 'condition*'; $cMean = HCol 'what it means*'; $cStat = HCol 'status*'; $cOwn = HCol 'owner*'; $cCom = HCol 'comment*'; $cId = HCol 'id'
if (-not $cLoc -or -not $cCond -or -not $cStat -or -not $cCom) { throw 'The issue sheet does not have the columns Location, Condition, Status and Comment.' }
$findings = New-Object System.Collections.Generic.List[object]
foreach ($r in ($rowsEl | Select-Object -Skip 1)) {
    $vals = @{}
    foreach ($c in $r.SelectNodes('m:c', $nsm)) { $vals[($c.GetAttribute('r') -replace '\d', '')] = Cell-Text $c }
    if (-not $vals[$cLoc] -and -not $vals[$cCond]) { continue }
    $findings.Add(@{ Row = [int]$r.GetAttribute('r'); Loc = [string]$vals[$cLoc]; Rule = [string]$vals[$cRule]; Cond = [string]$vals[$cCond]; Mean = [string]$vals[$cMean]; Id = [string]$vals[$cId] })
}
Write-Host ('Looking up ' + $findings.Count + ' of ' + $findings.Count + ' findings')

# ---------------------------------------------------------------- 2. the deck
Copy-Item -LiteralPath $src -Destination $outPptx -Force
$zip = [IO.Compression.ZipFile]::Open($outPptx, 'Update')
$docs = @{}; $dirty = @{}; $origText = @{}
$nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main'; $nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
try {
    function Get-Doc([string]$name) {
        if ($docs.ContainsKey($name)) { return $docs[$name] }
        $t = Read-Entry $zip $name; if (-not $t) { return $null }
        $d = New-Xml $t; $docs[$name] = $d; return $d
    }
    function Get-Nsm($d) { $n = New-Object Xml.XmlNamespaceManager($d.NameTable); $n.AddNamespace('p', $nsP); $n.AddNamespace('a', $nsA); $n.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'); return ,$n }
    $pres = Read-Entry $zip 'ppt/presentation.xml'; $prel = Read-Entry $zip 'ppt/_rels/presentation.xml.rels'
    $relT = @{}
    foreach ($m in [regex]::Matches($prel, '<Relationship\s[^>]*?>')) {
        $id = [regex]::Match($m.Value, 'Id="([^"]+)"').Groups[1].Value; $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
        if ($id) { $relT[$id] = $tg }
    }
    $slideParts = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($pres, '<p:sldId\s[^>]*?r:id="([^"]+)"')) { $slideParts.Add('ppt/' + ($relT[$m.Groups[1].Value] -replace '^/?(ppt/)?', '')) }
    Write-Host ('Slides in the deck: ' + $slideParts.Count)

    function Slide-Part([int]$n) { if ($n -ge 1 -and $n -le $slideParts.Count) { return $slideParts[$n - 1] }; return $null }
    function Rel-Target([string]$part, [string]$type) {
        $rp = ($part -replace '[^/]+$', '_rels/') + ([IO.Path]::GetFileName($part)) + '.rels'
        $t = Read-Entry $zip $rp; if (-not $t) { return $null }
        foreach ($m in [regex]::Matches($t, '<Relationship\s[^>]*?>')) {
            if ($m.Value -match ('Type="[^"]*/' + $type + '"')) {
                $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
                $parts = New-Object System.Collections.Generic.List[string]
                foreach ($s in (($part -replace '[^/]+$', '') + $tg).Split('/')) { if ($s -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count - 1) } } elseif ($s -and $s -ne '.') { $parts.Add($s) } }
                return ($parts -join '/')
            }
        }
        return $null
    }
    function Shape-Children($tree) {
        $l = New-Object System.Collections.Generic.List[object]
        foreach ($c in $tree.ChildNodes) { if ($c.NodeType -eq 'Element' -and @('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'contentPart') -contains $c.LocalName) { $l.Add($c) } }
        return ,$l
    }
    function Is-Title($sh) {
        if ($sh.LocalName -ne 'sp') { return $false }
        $ph = $sh.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', (Get-Nsm $sh.OwnerDocument))
        return ($ph -and @('title', 'ctrTitle') -contains $ph.GetAttribute('type'))
    }
    function Shape-Text($sh) { $n = Get-Nsm $sh.OwnerDocument; return ((@($sh.SelectNodes('.//a:t', $n)) | ForEach-Object { $_.InnerText }) -join '') }
    function Slide-Title([int]$n) {
        $d = Get-Doc (Slide-Part $n); if (-not $d) { return $null }
        $tree = $d.SelectSingleNode('//p:cSld/p:spTree', (Get-Nsm $d))
        foreach ($sh in (Shape-Children $tree)) { if (Is-Title $sh) { return (Shape-Text $sh).Trim() } }
        return $null
    }
    function Get-Bounds($sh, [string]$slidePart) {
        $n = Get-Nsm $sh.OwnerDocument
        $x = $sh.SelectSingleNode('p:spPr/a:xfrm', $n)
        if (-not $x) { $x = $sh.SelectSingleNode('p:xfrm', $n) }
        if (-not $x) { $x = $sh.SelectSingleNode('p:grpSpPr/a:xfrm', $n) }
        if (-not $x) {
            $ph = $sh.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $n)
            if ($ph) {
                $lay = Rel-Target $slidePart 'slideLayout'
                foreach ($part in @($lay, $(if ($lay) { Rel-Target $lay 'slideMaster' }))) {
                    if (-not $part) { continue }
                    $ld = Get-Doc $part; if (-not $ld) { continue }
                    $ln = Get-Nsm $ld; $idx = $ph.GetAttribute('idx'); $ty = $ph.GetAttribute('type')
                    foreach ($ps in $ld.SelectNodes('//p:cSld/p:spTree/p:sp', $ln)) {
                        $lp = $ps.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $ln); if (-not $lp) { continue }
                        $match = $false
                        if ($idx -and $lp.GetAttribute('idx') -eq $idx) { $match = $true } elseif (-not $idx -and $ty -and $lp.GetAttribute('type') -eq $ty) { $match = $true } elseif (-not $idx -and -not $ty -and -not $lp.GetAttribute('type')) { $match = $true }
                        if ($match) { $x = $ps.SelectSingleNode('p:spPr/a:xfrm', $ln); if ($x) { break } }
                    }
                    if ($x) { break }
                }
            }
        }
        if (-not $x) { return $null }
        $o = $x.SelectSingleNode('a:off', $n); $e = $x.SelectSingleNode('a:ext', $n); if (-not $o -or -not $e) { return $null }
        return @([long]$o.GetAttribute('x'), [long]$o.GetAttribute('y'), [long]$e.GetAttribute('cx'), [long]$e.GetAttribute('cy'))
    }
    # the order a person would read the slide in: title first, then top to bottom, left to right
    function Get-WantedOrder($shapes, [string]$part) {
        $items = @(); $i = 0
        foreach ($sh in $shapes) {
            $b = Get-Bounds $sh $part
            if (-not $b) { return $null }
            $items += [pscustomobject]@{ Sh = $sh; T = [bool](Is-Title $sh); X = $b[0]; Y = $b[1]; W = $b[2]; H = $b[3]; I = $i }; $i++
        }
        for ($a = 0; $a -lt $items.Count; $a++) { for ($c = $a + 1; $c -lt $items.Count; $c++) {
            $p = $items[$a]; $q = $items[$c]
            $ox = [math]::Min($p.X + $p.W, $q.X + $q.W) - [math]::Max($p.X, $q.X); $oy = [math]::Min($p.Y + $p.H, $q.Y + $q.H) - [math]::Max($p.Y, $q.Y)
            if ($ox -gt 12700 -and $oy -gt 12700) { return 'overlap' }
        } }
        $sorted = @($items | Sort-Object @{ Expression = { if ($_.T) { 0 } else { 1 } } }, @{ Expression = { [math]::Round($_.Y / 342900) } }, @{ Expression = { $_.X } }, @{ Expression = { $_.I } })
        return ,$sorted
    }
    function Apply-Order($tree, $sorted) {
        $first = $null
        foreach ($c in $tree.ChildNodes) { if ($c.NodeType -eq 'Element' -and $c.LocalName -eq 'grpSpPr') { $first = $c } }
        $anchor = $first
        foreach ($it in $sorted) {
            $tree.RemoveChild($it.Sh) | Out-Null
            if ($anchor.NextSibling) { $tree.InsertBefore($it.Sh, $anchor.NextSibling) | Out-Null } else { $tree.AppendChild($it.Sh) | Out-Null }
            $anchor = $it.Sh
        }
    }
    function Mark-Dirty([string]$part) { $dirty[$part] = $true }
    function Text-Bag($d) { $n = Get-Nsm $d; return ((@($d.SelectNodes('//a:t', $n)) | ForEach-Object { $_.InnerText } | Sort-Object) -join "`n") }

    # titles of every slide (for the repeated-title check)
    $titles = @{}
    for ($i = 1; $i -le $slideParts.Count; $i++) { $titles[$i] = Slide-Title $i }
    $quizWords = @('review', 'review questions', 'chapter review', 'quiz', 'check your understanding', 'test your knowledge', 'knowledge check')

    # ---------------------------------------------------------------- 3. go through the findings
    $res = @{}
    $nDone = 0
    foreach ($f in $findings) {
        $nDone++
        if ($nDone % 10 -eq 0) { Write-Host ('Looking up ' + $nDone + ' of ' + $findings.Count + ' findings') }
        $status = 'Not fixed - needs a person'; $why = 'This tool has no automatic check for this kind of finding yet.'
        $sl = 0; $m = [regex]::Match($f.Loc, '(?i)slide\s+(\d+)'); if ($m.Success) { $sl = [int]$m.Groups[1].Value }
        $part = if ($sl) { Slide-Part $sl } else { $null }
        $d = if ($part) { Get-Doc $part } else { $null }
        try {
        if (-not $d) { $why = 'The slide named in this finding was not found in the deck.' }
        elseif ($f.Cond -match '(?i)different regional language') {
            $mm = [regex]::Match($f.Mean, "declare '([A-Za-z-]+)' while the presentation's default is '([A-Za-z-]+)'")
            if ($mm.Success) {
                $from = $mm.Groups[1].Value; $to = $mm.Groups[2].Value; $n = Get-Nsm $d; $cnt = 0
                foreach ($nd in @($d.SelectNodes('//a:rPr[@lang]|//a:endParaRPr[@lang]|//a:defRPr[@lang]', $n))) { if ($nd.GetAttribute('lang') -eq $from) { $nd.SetAttribute('lang', $to); $cnt++ } }
                if ($cnt) { Mark-Dirty $part; $status = 'Fixed'; $why = 'Set ' + $cnt + ' text run(s) from ' + $from + ' to ' + $to + ' (the presentation default). Only the language tag changed, no text.' }
                else { $why = 'No text with the ' + $from + ' tag was found on this slide any more.' }
            }
        }
        elseif ($f.Cond -match '(?i)duplicate slide title') {
            $mm = [regex]::Match($f.Mean, "repeats the title '(.*)' \(also on slide (\d+)\)")
            $tl = if ($mm.Success) { $mm.Groups[1].Value } else { [string]$titles[$sl] }
            $run = 1; $k = $sl - 1; while ($k -ge 1 -and $titles[$k] -eq $titles[$sl]) { $run++; $k-- }; $k = $sl + 1; while ($k -le $slideParts.Count -and $titles[$k] -eq $titles[$sl]) { $run++; $k++ }
            if ($quizWords -contains $tl.ToLowerInvariant() -and $run -ge 3) {
                $status = 'False positive (likely)'; $why = 'The title "' + $tl + '" is a label repeated on ' + $run + ' slides in a row (question and answer slides). Nothing changed. A person should confirm that the repeated label is intended.'
            } else {
                $status = 'Not fixed - needs a content change'; $why = 'Making the title unique means changing the title text, which this tool never does. SLID Macro numbers repeated titles as (1 of N).'
            }
        }
        elseif ($f.Cond -match '(?i)contrast') {
            $mm = [regex]::Match($f.Mean, 'text #([0-9A-Fa-f]{6}) on background #([0-9A-Fa-f]{6}) has a contrast ratio of ([0-9.]+):1, below the required ([0-9.]+):1')
            if ($mm.Success) {
                $fg = $mm.Groups[1].Value.ToUpperInvariant(); $bg = $mm.Groups[2].Value.ToUpperInvariant(); $need = [double]::Parse($mm.Groups[4].Value, $ci)
                $new = Get-Darker $fg $bg ($need + 0.2); $ratio = Get-Ratio $new $bg
                $n = Get-Nsm $d; $cnt = 0
                foreach ($nd in @($d.SelectNodes('//a:rPr/a:solidFill/a:srgbClr|//a:endParaRPr/a:solidFill/a:srgbClr|//a:defRPr/a:solidFill/a:srgbClr', $n))) { if ($nd.GetAttribute('val').ToUpperInvariant() -eq $fg) { $nd.SetAttribute('val', $new); $cnt++ } }
                if ($cnt) { Mark-Dirty $part; $status = 'Fixed'; $why = 'Darkened the text colour of ' + $cnt + ' run(s) from #' + $fg + ' to #' + $new + ' (contrast on #' + $bg + ' is now ' + [math]::Round($ratio, 2).ToString($ci) + ':1, needed ' + $need.ToString($ci) + ':1). Only the colour changed, no text.' }
                else { $why = 'The colour #' + $fg + ' is not set on the text of this slide (it comes from the theme or layout), so it was left alone.' }
            }
        }
        elseif ($f.Cond -match '(?i)title is not first') {
            $tree = $d.SelectSingleNode('//p:cSld/p:spTree', (Get-Nsm $d)); $sh = Shape-Children $tree
            $t = $null; foreach ($s in $sh) { if (Is-Title $s) { $t = $s; break } }
            if (-not $t) { $why = 'No title was found on this slide.' }
            elseif ($sh[0] -eq $t) { $status = 'False positive (likely)'; $why = 'The title is already the first shape in the reading order.' }
            else {
                $grp = $null; foreach ($c in $tree.ChildNodes) { if ($c.NodeType -eq 'Element' -and $c.LocalName -eq 'grpSpPr') { $grp = $c } }
                $tree.RemoveChild($t) | Out-Null; $tree.InsertAfter($t, $grp) | Out-Null; Mark-Dirty $part
                $status = 'Fixed'; $why = 'Moved the title to the start of the reading order. Nothing was moved on the slide and no text changed.'
            }
        }
        elseif ($f.Cond -match '(?i)reading order may not match') {
            $tree = $d.SelectSingleNode('//p:cSld/p:spTree', (Get-Nsm $d)); $sh = Shape-Children $tree
            $w = Get-WantedOrder $sh $part
            if ($null -eq $w) { $why = 'The position of some shapes could not be read, so the order was not changed.' }
            elseif ($w -is [string]) { $why = 'Some shapes overlap, so changing their order could change how the slide looks. Please set the order in the Selection Pane.' }
            else {
                $same = $true; for ($i = 0; $i -lt $w.Count; $i++) { if (-not [object]::ReferenceEquals($w[$i].Sh, $sh[$i])) { $same = $false; break } }
                if ($same) { $status = 'False positive (likely)'; $why = 'The reading order already follows the visual order (title first, then top to bottom, left to right).' }
                else { Apply-Order $tree $w; Mark-Dirty $part; $status = 'Fixed'; $why = 'Put the ' + $w.Count + ' shapes in reading order (title first, then top to bottom, left to right). The shapes do not overlap, so nothing changes on screen, and no text changed.' }
            }
        }
        elseif ($f.Cond -match '(?i)loose text box') {
            $tree = $d.SelectSingleNode('//p:cSld/p:spTree', (Get-Nsm $d)); $sh = Shape-Children $tree
            $loose = @(); foreach ($s in $sh) { if ($s.LocalName -eq 'sp' -and -not $s.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', (Get-Nsm $d)) -and (Shape-Text $s).Trim()) { $loose += $s } }
            $cr = [string][char]169
            $credit = $true; foreach ($s in $loose) { if ((Shape-Text $s).Trim() -notmatch ('(?i)^(' + $cr + '|\(c\)|source|photo|credit|courtesy|image|reprinted|used with)')) { $credit = $false } }
            $w = Get-WantedOrder $sh $part
            $okOrder = $false
            if ($w -and -not ($w -is [string])) { $okOrder = $true; for ($i = 0; $i -lt $w.Count; $i++) { if (-not [object]::ReferenceEquals($w[$i].Sh, $sh[$i])) { $okOrder = $false; break } } }
            if ($loose.Count -gt 0 -and $credit) { $status = 'False positive (likely)'; $why = 'The text box(es) are credit or source lines, which normally sit next to their picture as text boxes. Reading order was checked.' }
            elseif ($okOrder) { $status = 'False positive (likely)'; $why = 'The text boxes are already in the right reading order (title first, then top to bottom, left to right). Moving the text into placeholders would change the layout, so it was left as it is.' }
            else { $why = 'The text boxes are not in a clear reading order and the slide has overlapping or unreadable shapes. Please check the Selection Pane.' }
        }
        } catch { $status = 'Not fixed - needs a person'; $why = 'The tool could not process this finding: ' + $_.Exception.Message }
        $res[$f.Row] = @{ Status = $status; Why = $why }
    }

    # ---------------------------------------------------------------- 4. the content guard: no text may have changed
    foreach ($part in @($dirty.Keys)) {
        $before = Read-Entry $zip $part
        $bagB = Text-Bag (New-Xml $before); $bagA = Text-Bag $docs[$part]
        if ($bagB -ne $bagA) { throw ('Stopped: the text on ' + $part + ' would have changed. Nothing was saved.') }
    }
    foreach ($part in @($dirty.Keys)) { Write-Entry $zip $part (Save-Xml $docs[$part]) }
} catch {
    try { $zip.Dispose() } catch {}
    Remove-Item -LiteralPath $outPptx -Force -ErrorAction SilentlyContinue
    throw
}
$zip.Dispose()

# ---------------------------------------------------------------- 5. the report: the same sheet with Status / Owner / Comment filled in
Copy-Item -LiteralPath $xin -Destination $outXlsx -Force
$oz = [IO.Compression.ZipFile]::Open($outXlsx, 'Update')
try {
    foreach ($f in $findings) {
        $row = $rowsEl | Where-Object { [int]$_.GetAttribute('r') -eq $f.Row } | Select-Object -First 1
        foreach ($pair in @(@($cStat, $res[$f.Row].Status), @($cOwn, 'S4C Issue Fixer'), @($cCom, $res[$f.Row].Why))) {
            if (-not $pair[0]) { continue }
            $ref = $pair[0] + $f.Row
            $cell = $row.SelectSingleNode("m:c[@r='" + $ref + "']", $nsm)
            if (-not $cell) {
                $cell = $fx.CreateElement('c', $nsm.LookupNamespace('m')); $cell.SetAttribute('r', $ref)
                $row.AppendChild($cell) | Out-Null
            }
            while ($cell.HasChildNodes) { [void]$cell.RemoveChild($cell.FirstChild) }
            $cell.SetAttribute('t', 'inlineStr')
            $is = $fx.CreateElement('is', $nsm.LookupNamespace('m')); $t = $fx.CreateElement('t', $nsm.LookupNamespace('m')); $t.InnerText = [string]$pair[1]; $is.AppendChild($t) | Out-Null; $cell.AppendChild($is) | Out-Null
        }
    }
    Write-Entry $oz $findPath (Save-Xml $fx)
    # a short summary on the Notice sheet
    $noticeKey = $null; foreach ($k in $sheets.Keys) { if ($k -match '(?i)notice') { $noticeKey = $k } }
    if ($noticeKey) {
        $np = 'xl/' + ($targetOf[$sheets[$noticeKey]] -replace '^/?(xl/)?', '')
        $nx = New-Xml (Read-Entry $oz $np); $nn = New-Object Xml.XmlNamespaceManager($nx.NameTable); $nn.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $sd = $nx.SelectSingleNode('//m:sheetData', $nn); $last = 0; foreach ($r in $sd.SelectNodes('m:row', $nn)) { $last = [math]::Max($last, [int]$r.GetAttribute('r')) }
        $cnt = @{}; foreach ($f in $findings) { $s = $res[$f.Row].Status; if ($cnt.ContainsKey($s)) { $cnt[$s]++ } else { $cnt[$s] = 1 } }
        $verLine = 'Version ' + $IF_VERSION + ', ' + (Get-Date).ToString('yyyy-MM-dd HH:mm', $ci)
        $lines = @(@('Issue Fixer (S4C)', $verLine), @('Findings checked', [string]$findings.Count))
        foreach ($k in ($cnt.Keys | Sort-Object)) { $lines += , @($k, [string]$cnt[$k]) }
        $lines += , @('Fixed copy of the deck', [IO.Path]::GetFileName($outPptx))
        $rn = $last + 2
        foreach ($ln in $lines) {
            $rEl = $nx.CreateElement('row', $nn.LookupNamespace('m')); $rEl.SetAttribute('r', [string]$rn)
            for ($ci2 = 0; $ci2 -lt 2; $ci2++) {
                $cEl = $nx.CreateElement('c', $nn.LookupNamespace('m')); $cEl.SetAttribute('r', @('A', 'B')[$ci2] + $rn); $cEl.SetAttribute('t', 'inlineStr')
                $is = $nx.CreateElement('is', $nn.LookupNamespace('m')); $t = $nx.CreateElement('t', $nn.LookupNamespace('m')); $t.InnerText = [string]$ln[$ci2]; $is.AppendChild($t) | Out-Null; $cEl.AppendChild($is) | Out-Null; $rEl.AppendChild($cEl) | Out-Null
            }
            $sd.AppendChild($rEl) | Out-Null; $rn++
        }
        Write-Entry $oz $np (Save-Xml $nx)
    }
} finally { $oz.Dispose() }

Add-WorkLog 'Issue Fixer' $src $outPptx
$sum = @{}; foreach ($f in $findings) { $s = $res[$f.Row].Status; if ($sum.ContainsKey($s)) { $sum[$s]++ } else { $sum[$s] = 1 } }
foreach ($k in ($sum.Keys | Sort-Object)) { Write-Host ('  ' + $k + ': ' + $sum[$k]) }
Write-Host ('Saved: ' + $outPptx)
Write-Host ('Saved: ' + $outXlsx)
}

# ---------------------------------------------------------------- start
Write-Host ('Issue Fixer ' + $IF_VERSION) -ForegroundColor Cyan
$p = $env:IF_PPTX; $x = $env:IF_XLSX
if (-not $p) { $p = (Read-Host 'Path of the PowerPoint file (.pptx)').Trim('"') }
if (-not $x) { $x = (Read-Host 'Path of the issue sheet (.xlsx)').Trim('"') }
try { Fix-Issues -PptxPath $p -XlsxPath $x }
catch { Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red }
if (-not $env:S4C_AUTO) { [void](Read-Host 'Finished. Press Enter to close this window') }
