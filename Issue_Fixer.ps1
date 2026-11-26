# Issue Fixer - reads the pre-check / tester issue sheet (.xlsx) and the PowerPoint deck it was made for, fixes what can be fixed
# safely (reading order, language tags, colour contrast), says which findings look like false positives, and fills the sheet's
# Status / Owner / Comment columns. The text and objects of the deck are NEVER changed or added to.
# S4C ISSUE FIXER SCRIPT
$IF_VERSION = '1.4'   # shown at the start of the window; the updater compares this number with GitHub
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$ci = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object Text.UTF8Encoding($false)

# notes the saved file in today's work, but with no slide or page count (running Issue Fixer must not add to the slides / pages goal)
function Add-WorkLog([string]$tool, [string]$inPath, [string]$outPath) {
    try {
        $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
        $dir = Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog'
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $now = Get-Date
        $line = @($now.ToString('yyyy-MM-dd', $ci), $now.ToString('HH:mm', $ci), $env:USERNAME, $tool, [IO.Path]::GetFullPath($inPath), [IO.Path]::GetFullPath($outPath), '', '') -join "`t"
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
        $fromLayout = $false
        if (-not $x) {
            $fromLayout = $true
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
        if ($fromLayout) { return @([long]$o.GetAttribute('x'), [long]$o.GetAttribute('y'), [long]$e.GetAttribute('cx'), [long]$e.GetAttribute('cy'), 'L') }
        return @([long]$o.GetAttribute('x'), [long]$o.GetAttribute('y'), [long]$e.GetAttribute('cx'), [long]$e.GetAttribute('cy'))
    }
    # ---- the reading order and shape names SLID Macro uses: title first, then column by column (or row by row) as the slide is laid out;
    # a shape that sits mostly behind a picture stays behind it. Shapes are named Title / Picture / TextBox (TextBox 1, TextBox 2 ... when there are several).
    $sldW = 12192000.0; $mW = [regex]::Match($pres, '<p:sldSz\s[^>]*?cx="(\d+)"'); if ($mW.Success) { $sldW = [double]$mW.Groups[1].Value }
    $slideWpt = $sldW / 12700
    function Get-SlideOrder([string]$part, [int]$slideIdx) {
        $d = Get-Doc $part; $m = Get-Nsm $d
        $tree = $d.SelectSingleNode('//p:cSld/p:spTree', $m)
        $items = New-Object System.Collections.Generic.List[object]; $order = 0
        foreach ($c in $tree.ChildNodes) {
            if ($c.NodeType -ne 'Element') { continue }
            $order++
            if (@('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp') -notcontains $c.LocalName) { continue }
            $kind = ''; $nv = $null
            switch ($c.LocalName) {
                'sp' {
                    $nv = $c.SelectSingleNode('p:nvSpPr/p:cNvPr', $m)
                    $ph = $c.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $m)
                    $pt = if ($ph) { $ph.GetAttribute('type') } else { '' }
                    $txt = Shape-Text $c
                    if ($pt -eq 'title' -or $pt -eq 'ctrTitle') { $kind = 'Title' }
                    elseif ($pt -eq 'pic') { $kind = 'Picture' }
                    elseif ($txt.Length -gt 0) { if ($pt -eq 'subTitle') { $kind = 'Subtitle' } else { $kind = 'TextBox' } }
                }
                'pic' { $nv = $c.SelectSingleNode('p:nvPicPr/p:cNvPr', $m); $kind = 'Picture' }
                'grpSp' { $nv = $c.SelectSingleNode('p:nvGrpSpPr/p:cNvPr', $m); if ($c.SelectSingleNode('.//p:pic', $m)) { $kind = 'Picture' } }
            }
            $rk = 1; if ($kind -eq 'Title') { $rk = 0 }
            $b = Get-Bounds $c $part
            if ($b) { $x = [double]$b[0]; $y = [double]$b[1]; $w = [double]$b[2]; $h = [double]$b[3]; if ($b.Count -gt 4) { $h = 0.0 } } else { $x = 0.0; $y = 1e12; $w = 0.0; $h = 0.0 }
            $items.Add([pscustomobject]@{ Kind = $kind; Nv = $nv; Node = $c; RK = $rk; Y = [Math]::Round($y / 12700); X = $x; CX = (($x + $w / 2) / 12700); Order = $order; L = $x; T = $y; W = $w; H = $h })
        }
        if ($slideIdx -eq 0 -and -not @($items | Where-Object { $_.Kind -eq 'Subtitle' }).Count) {
            $tbs = @($items | Where-Object { $_.Kind -eq 'TextBox' }); if ($tbs.Count -eq 1) { $tbs[0].Kind = 'Subtitle' }
        }
        $sorted = New-Object System.Collections.Generic.List[object]
        foreach ($it in @($items | Where-Object { $_.RK -eq 0 } | Sort-Object Y, X, Order)) { $sorted.Add($it) }
        $rest = @($items | Where-Object { $_.RK -ne 0 })
        $placed = @($rest | Where-Object { $_.Y -lt 1e9 } | Sort-Object CX, Y, Order)
        $unplaced = @($rest | Where-Object { $_.Y -ge 1e9 } | Sort-Object Order)
        $clusters = New-Object System.Collections.Generic.List[object]; $cur = $null
        foreach ($it in $placed) {
            if ($null -eq $cur -or ($it.CX - $cur.Last) -gt ($slideWpt * 0.15)) { $cur = [pscustomobject]@{ Last = $it.CX; Items = (New-Object System.Collections.Generic.List[object]) }; $clusters.Add($cur) }
            $cur.Items.Add($it); $cur.Last = $it.CX
        }
        $multi = @($clusters | Where-Object { $_.Items.Count -ge 2 }).Count
        if ($clusters.Count -ge 2 -and $multi -ge 2) {
            foreach ($cl in $clusters) { foreach ($it in @($cl.Items | Sort-Object Y, X, Order)) { $sorted.Add($it) } }
        } else {
            $rowList = New-Object System.Collections.Generic.List[object]; $rowItems = $null; $rowStart = 0
            foreach ($it in @($placed | Sort-Object Y, X, Order)) {
                if ($null -eq $rowItems -or ($it.Y - $rowStart) -gt 18) { $rowItems = New-Object System.Collections.Generic.List[object]; $rowStart = $it.Y; $rowList.Add($rowItems) }
                $rowItems.Add($it)
            }
            foreach ($r in $rowList) { foreach ($it in @($r | Sort-Object X, Order)) { $sorted.Add($it) } }
        }
        foreach ($it in $unplaced) { $sorted.Add($it) }
        $guard = 0; $moved = $true
        while ($moved -and $guard -lt 50) {
            $moved = $false; $guard++
            for ($i = 0; $i -lt $sorted.Count -and -not $moved; $i++) {
                $s1 = $sorted[$i]
                if ($s1.RK -eq 0 -or $s1.Kind -eq 'Picture' -or $s1.W -le 0 -or $s1.H -le 0 -or $s1.T -ge 1e11) { continue }
                for ($j = 0; $j -lt $i; $j++) {
                    $pc = $sorted[$j]
                    if ($pc.Kind -ne 'Picture' -or $pc.Order -le $s1.Order -or $pc.W -le 0 -or $pc.H -le 0 -or $pc.T -ge 1e11) { continue }
                    $ow = [Math]::Min($s1.L + $s1.W, $pc.L + $pc.W) - [Math]::Max($s1.L, $pc.L)
                    $oh = [Math]::Min($s1.T + $s1.H, $pc.T + $pc.H) - [Math]::Max($s1.T, $pc.T)
                    if ($ow -le 0 -or $oh -le 0) { continue }
                    if (($ow * $oh) / ($s1.W * $s1.H) -gt 0.5) { $sorted.RemoveAt($i); $sorted.Insert($j, $s1); $moved = $true; break }
                }
            }
        }
        return [pscustomobject]@{ Tree = $tree; Items = $items; Sorted = @($sorted.ToArray()) }
    }
    # puts the shapes of a slide in that reading order; -Rename also gives the text shapes their TextBox names. Returns what changed.
    function Apply-SlideOrder([string]$part, [int]$slideIdx, [bool]$rename) {
        $o = Get-SlideOrder $part $slideIdx
        $orderChanged = $false
        for ($i = 0; $i -lt $o.Sorted.Count; $i++) { if ($o.Sorted[$i].Order -ne $o.Items[$i].Order) { $orderChanged = $true; break } }
        if ($orderChanged) {
            $anchor = $o.Tree.SelectSingleNode('p:extLst', (Get-Nsm $o.Tree.OwnerDocument))
            foreach ($it in $o.Sorted) {
                [void]$o.Tree.RemoveChild($it.Node)
                if ($anchor) { [void]$o.Tree.InsertBefore($it.Node, $anchor) } else { [void]$o.Tree.AppendChild($it.Node) }
            }
        }
        $renamed = 0
        if ($rename) {
            # a placeholder that was already on the slide is never renamed or replaced; only plain text boxes and the placeholders this tool just gave them
            $tot = @($o.Sorted | Where-Object { $_.Nv -and ($_.Kind -eq 'TextBox' -or $_.Kind -eq 'Subtitle') -and ((-not $_.Node.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', (Get-Nsm $_.Node.OwnerDocument))) -or ($script:convNodes -contains $_.Node)) })
            $byKind = @{}; foreach ($it in $tot) { if ($byKind.ContainsKey($it.Kind)) { $byKind[$it.Kind]++ } else { $byKind[$it.Kind] = 1 } }
            $run = @{}
            foreach ($it in $tot) {
                if ($run.ContainsKey($it.Kind)) { $run[$it.Kind]++ } else { $run[$it.Kind] = 1 }
                $nm = if ($byKind[$it.Kind] -eq 1) { $it.Kind } else { $it.Kind + ' ' + $run[$it.Kind] }
                if ($it.Nv.GetAttribute('name') -ne $nm) { $it.Nv.SetAttribute('name', $nm); $renamed++ }
            }
        }
        return [pscustomobject]@{ OrderChanged = $orderChanged; Renamed = $renamed; Count = $o.Sorted.Count }
    }
    # ---- loose text boxes get the placeholder their role needs: a credit line / other text gets a text (body) placeholder, a title gets the title placeholder,
    # the only text on the first slide gets the subtitle placeholder. Position, size, text and look stay as they are (the box keeps its own xfrm; bullets and spacing the
    # master would add to a body placeholder are switched off, and a missing font size is set to the 18 pt a text box has).
    $script:convNodes = New-Object System.Collections.Generic.List[object]
    function Convert-Loose([string]$part, [int]$slideIdx) {
        $out = [pscustomobject]@{ Credit = 0; Title = 0; Subtitle = 0; Body = 0 }
        $d = Get-Doc $part; $n = Get-Nsm $d
        $tree = $d.SelectSingleNode('//p:cSld/p:spTree', $n)
        $hasTitle = $false; $hasSub = $false; $used = New-Object System.Collections.Generic.List[int]; $loose = New-Object System.Collections.Generic.List[object]
        foreach ($sh in (Shape-Children $tree)) {
            if ($sh.LocalName -ne 'sp') { continue }
            $ph = $sh.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $n)
            if ($ph) {
                if (@('title', 'ctrTitle') -contains $ph.GetAttribute('type')) { $hasTitle = $true }
                if ($ph.GetAttribute('type') -eq 'subTitle') { $hasSub = $true }
                $ix = $ph.GetAttribute('idx'); if ($ix -match '^\d+$') { $used.Add([int]$ix) }
            } elseif ((Shape-Text $sh).Trim()) { $loose.Add($sh) }
        }
        if ($loose.Count -eq 0) { return $out }
        $layout = Rel-Target $part 'slideLayout'; $layHasSub = $false
        if ($layout) {
            $ld = Get-Doc $layout; if ($ld) {
                $ln = Get-Nsm $ld
                foreach ($p2 in @($ld.SelectNodes('//p:sp/p:nvSpPr/p:nvPr/p:ph', $ln))) {
                    $ix = $p2.GetAttribute('idx'); if ($ix -match '^\d+$') { $used.Add([int]$ix) }
                    if ($p2.GetAttribute('type') -eq 'subTitle') { $layHasSub = $true }
                }
            }
        }
        $next = 10; foreach ($u in $used) { if ($u -ge $next) { $next = $u + 1 } }
        # the topmost loose box is the title when the slide has none and it sits in the top fifth of the slide
        $sldH = 6858000.0; $mH = [regex]::Match($pres, '<p:sldSz\s[^>]*?cy="(\d+)"'); if ($mH.Success) { $sldH = [double]$mH.Groups[1].Value }
        $order = @($loose | Sort-Object { $b = Get-Bounds $_ $part; if ($b) { [double]$b[1] } else { 1e12 } })
        $subDone = $false
        foreach ($sh in $order) {
            $text = (Shape-Text $sh).Trim()
            $b = Get-Bounds $sh $part
            $role = 'body'
            if ($text -match '^(\u00A9|\(c\)|credit|source\s*:|photo|image\s+credit)' -or $text -match '^\s*.?\s*\u00A9') { $role = 'credit' }
            elseif (-not $hasTitle -and $b -and ([double]$b[1] -lt ($sldH * 0.2))) { $role = 'title' }
            elseif ($slideIdx -eq 0 -and $layHasSub -and -not $subDone -and -not $hasSub -and $loose.Count -eq 1) { $role = 'subtitle' }
            $nvPr = $sh.SelectSingleNode('p:nvSpPr/p:nvPr', $n); $ph = $d.CreateElement('p', 'ph', $nsP)
            if ($role -eq 'title') { $ph.SetAttribute('type', 'title'); $hasTitle = $true }
            elseif ($role -eq 'subtitle') { $ph.SetAttribute('type', 'subTitle'); $ph.SetAttribute('idx', '1'); $subDone = $true }
            else { $ph.SetAttribute('type', 'body'); $ph.SetAttribute('sz', 'quarter'); $ph.SetAttribute('idx', [string]$next); $next++ }
            [void]$nvPr.PrependChild($ph); $script:convNodes.Add($sh)
            $cnv = $sh.SelectSingleNode('p:nvSpPr/p:cNvSpPr', $n)
            if ($cnv) {
                if ($cnv.GetAttribute('txBox')) { [void]$cnv.Attributes.RemoveNamedItem('txBox') }
                if (-not $cnv.SelectSingleNode('a:spLocks', $n)) { $lk = $d.CreateElement('a', 'spLocks', $nsA); $lk.SetAttribute('noGrp', '1'); [void]$cnv.AppendChild($lk) }
            }
            if ($role -eq 'title') {
                $nv = $sh.SelectSingleNode('p:nvSpPr/p:cNvPr', $n); if ($nv -and $nv.GetAttribute('name') -notmatch '^Title') { $nv.SetAttribute('name', 'Title 1') }
            }
            if ($role -ne 'title') {
                foreach ($para in @($sh.SelectNodes('p:txBody/a:p', $n))) {
                    $pPr = $para.SelectSingleNode('a:pPr', $n)
                    if (-not $pPr) { $pPr = $d.CreateElement('a', 'pPr', $nsA); [void]$para.PrependChild($pPr) }
                    if (-not $pPr.HasAttribute('marL')) { $pPr.SetAttribute('marL', '0') }
                    if (-not $pPr.HasAttribute('indent')) { $pPr.SetAttribute('indent', '0') }
                    if (-not $pPr.SelectSingleNode('a:spcBef', $n)) {
                        $sb = $d.CreateElement('a', 'spcBef', $nsA); $sp = $d.CreateElement('a', 'spcPts', $nsA); $sp.SetAttribute('val', '0'); [void]$sb.AppendChild($sp)
                        $after = $pPr.SelectSingleNode('a:lnSpc', $n)
                        if ($after) { [void]$pPr.InsertAfter($sb, $after) } else { [void]$pPr.PrependChild($sb) }
                    }
                    if (-not ($pPr.SelectSingleNode('a:buNone', $n) -or $pPr.SelectSingleNode('a:buChar', $n) -or $pPr.SelectSingleNode('a:buAutoNum', $n))) {
                        $bu = $d.CreateElement('a', 'buNone', $nsA)
                        $anchor = $null; foreach ($nm in @('buFont', 'buClr', 'buSzPct', 'buSzPts', 'buSzTx', 'spcAft', 'spcBef', 'lnSpc')) { $anchor = $pPr.SelectSingleNode('a:' + $nm, $n); if ($anchor) { break } }
                        if ($anchor) { [void]$pPr.InsertAfter($bu, $anchor) } else { [void]$pPr.PrependChild($bu) }
                    }
                    foreach ($rp in @($para.SelectNodes('a:r/a:rPr', $n))) { if (-not $rp.HasAttribute('sz')) { $rp.SetAttribute('sz', '1800') } }
                }
            }
            if ($role -eq 'credit') { $out.Credit++ } elseif ($role -eq 'title') { $out.Title++ } elseif ($role -eq 'subtitle') { $out.Subtitle++ } else { $out.Body++ }
        }
        return $out
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
        elseif ($f.Cond -match '(?i)reading order may not match|loose text box') {
            # handled after the other findings: the slide gets the SLID Macro reading order and shape names in one go
            $status = 'DEFER'
        }
        } catch { $status = 'Not fixed - needs a person'; $why = 'The tool could not process this finding: ' + $_.Exception.Message }
        $res[$f.Row] = @{ Status = $status; Why = $why; Sl = $sl; Part = $part }
    }
    # reading order + names (SLID Macro rules) once per slide, then every deferred finding on that slide gets the result
    $slideDone = @{}
    $looseParts = @{}; foreach ($f in $findings) { if ($res[$f.Row].Status -eq 'DEFER' -and $f.Cond -match '(?i)loose text box' -and $res[$f.Row].Part) { $looseParts[$res[$f.Row].Part] = $true } }
    foreach ($f in $findings) {
        if ($res[$f.Row].Status -ne 'DEFER') { continue }
        $part = $res[$f.Row].Part; $why = ''; $status = 'Not fixed - needs a person'
        try {
            if (-not $part) { $why = 'The slide named in this finding was not found in the deck.' }
            else {
                if (-not $slideDone.ContainsKey($part)) {
                    $idx = [int]$res[$f.Row].Sl - 1
                    $cv = $null
                    if ($looseParts.ContainsKey($part)) { $cv = Convert-Loose $part $idx; if ($cv.Credit + $cv.Title + $cv.Subtitle + $cv.Body) { Mark-Dirty $part } }
                    $slideDone[$part] = Apply-SlideOrder $part $idx $true
                    $slideDone[$part] | Add-Member -NotePropertyName Conv -NotePropertyValue $cv
                    if ($slideDone[$part].OrderChanged -or $slideDone[$part].Renamed) { Mark-Dirty $part }
                }
                $r = $slideDone[$part]; $isLoose = ($f.Cond -match '(?i)loose text box')
                $cvn = 0; if ($r.Conv) { $cvn = $r.Conv.Credit + $r.Conv.Title + $r.Conv.Subtitle + $r.Conv.Body }
                if ($isLoose -and $cvn) {
                    $status = 'Fixed'; $bits = @()
                    if ($r.Conv.Credit) { $bits += ('a text placeholder for ' + $r.Conv.Credit + ' credit line(s)') }
                    if ($r.Conv.Title) { $bits += 'the title placeholder for the slide title' }
                    if ($r.Conv.Subtitle) { $bits += 'the subtitle placeholder for the subtitle' }
                    if ($r.Conv.Body) { $bits += ('a text placeholder for ' + $r.Conv.Body + ' text box(es)') }
                    $why = 'Gave ' + ($bits -join ', ') + ', named TextBox in the Selection Pane and placed in the reading order. Text, position, size and look are unchanged.'
                } elseif ($r.OrderChanged -or $r.Renamed) {
                    $status = 'Fixed'
                    $why = 'Put the shapes in the reading order SLID Macro uses (title first, then column by column or top to bottom, left to right; anything sitting behind a picture stays behind it)'
                    if ($r.Renamed) { $why += ' and named the text boxes TextBox in the Selection Pane (TextBox 1, TextBox 2 when there are several)' }
                    $why += '. Credit lines are text boxes too. Text, position, size and look are unchanged.'
                } else {
                    $status = 'False positive (likely)'
                    if ($isLoose) { $why = 'The text boxes are already in the correct reading order and named TextBox in the Selection Pane. Credit lines normally stay as text boxes, so nothing was changed.' }
                    else { $why = 'The reading order already follows the visual order (title first, then the layout of the slide).' }
                }
            }
        } catch { $why = 'The tool could not process this finding: ' + $_.Exception.Message }
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
