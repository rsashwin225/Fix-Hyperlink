# ALT TEXT TOOL SCRIPT
# PPT Alt-Text Filler (S4Carlisle) - PowerShell version, no Python needed.
#
#   1. Pick the PowerPoint file, then the publisher's Alt-Text Excel log.
#   2. The tool matches every picture to a row of the log (same matching as the earlier Python tool).
#   3. A review page opens in your browser: each picture with a thumbnail, the suggested Element ID and the alt
#      text that will be written. Green = confident, amber = check the picture, red = pick the Element ID yourself.
#      You can also type your own alt text, mark a picture decorative, or add the long description.
#   4. Press "Apply" - a new copy <name>_with_alt.pptx is saved next to the original. The original is never changed.
#
# Self-updates from GitHub through AltText_Update.ps1 (see the version line below).
$ALT_VERSION = '2.0'   # shown at the start of the window; the updater compares this number with GitHub

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NS_SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$DECOR_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'
$DECOR_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative'

$STOP = @{}
foreach ($w in ('a an the of in on at to for and or with by from is are was were be as this that these those it its into over under shows show showing displays display depicts depict illustrates illustration photo picture image graphic labeled reads read text figure table titled title'.Split(' '))) { $STOP[$w] = $true }
$ABBR = [ordered]@{ 'ad' = 'additional duty'; 'ems' = 'emergency medical services'; 'emt' = 'emergency medical technician' }
$AUTO_THRESHOLD = 0.55   # match score needed to pre-fill
$MARGIN = 0.12           # best must beat 2nd best by this much

# ----------------------------------------------------------------- small helpers
function New-XmlDoc([string]$text) {
    $d = New-Object Xml.XmlDocument
    $d.PreserveWhitespace = $true
    $d.LoadXml($text)
    return $d
}
function New-Nsmgr([Xml.XmlDocument]$doc) {
    $m = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    [void]$m.AddNamespace('p', $NS_P); [void]$m.AddNamespace('a', $NS_A)
    [void]$m.AddNamespace('r', $NS_R); [void]$m.AddNamespace('rel', $NS_REL); [void]$m.AddNamespace('x', $NS_SS)
    return ,$m
}
function Read-ZipBytes($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { return $null }
    $s = $e.Open(); $m = New-Object IO.MemoryStream
    try { [void]$s.CopyTo($m) } finally { $s.Dispose() }
    return ,$m.ToArray()
}
function Read-ZipText($zip, [string]$name) {
    $b = Read-ZipBytes $zip $name
    if ($null -eq $b) { return $null }
    $t = [Text.Encoding]::UTF8.GetString($b)
    return $t.TrimStart([char]0xFEFF)
}
function Get-Rels($zip, [string]$part) {
    $dir = ''; $leaf = $part
    $i = $part.LastIndexOf('/')
    if ($i -ge 0) { $dir = $part.Substring(0, $i + 1); $leaf = $part.Substring($i + 1) }
    $text = Read-ZipText $zip ($dir + '_rels/' + $leaf + '.rels')
    $map = @{}
    if ($null -eq $text) { return $map }
    $doc = New-XmlDoc $text
    foreach ($rel in $doc.DocumentElement.ChildNodes) {
        if ($rel.NodeType -ne 'Element') { continue }
        $map[$rel.GetAttribute('Id')] = @{ target = $rel.GetAttribute('Target'); type = $rel.GetAttribute('Type'); mode = $rel.GetAttribute('TargetMode') }
    }
    return $map
}
function Resolve-Part([string]$fromPart, [string]$target) {
    $dir = ''
    $i = $fromPart.LastIndexOf('/')
    if ($i -ge 0) { $dir = $fromPart.Substring(0, $i) }
    $combined = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "$dir/$target" }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($seg in ($combined -replace '\\', '/').Split('/')) {
        if ($seg -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) } }
        elseif ($seg -eq '.' -or $seg -eq '') { continue }
        else { $parts.Add($seg) }
    }
    return ($parts -join '/')
}
function Get-Chapter([string]$path) {
    $m = [regex]::Match([IO.Path]::GetFileName($path), '(?:PPT|CH|Chapter)[_\- ]?0*(\d{1,3})', 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value.PadLeft(2, '0') }
    return $null
}

# ----------------------------------------------------------------- Excel log (xlsx read without Excel)
function Get-ColIndex([string]$cellRef) {
    $letters = ($cellRef -replace '[0-9]', '').ToUpperInvariant()
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}
function Read-Sheets([string]$xlsx) {
    $zip = [IO.Compression.ZipFile]::OpenRead($xlsx)
    try {
        $shared = New-Object System.Collections.Generic.List[string]
        $sst = Read-ZipText $zip 'xl/sharedStrings.xml'
        if ($sst) {
            $d = New-XmlDoc $sst; $nsm = New-Nsmgr $d
            foreach ($si in $d.SelectNodes('//x:si', $nsm)) {
                $sb = New-Object Text.StringBuilder
                foreach ($t in $si.SelectNodes('.//x:t[not(ancestor::x:rPh)]', $nsm)) { [void]$sb.Append($t.InnerText) }
                $shared.Add($sb.ToString())
            }
        }
        $wb = New-XmlDoc (Read-ZipText $zip 'xl/workbook.xml'); $wnsm = New-Nsmgr $wb
        $rels = Get-Rels $zip 'xl/workbook.xml'
        $sheets = @()
        foreach ($sh in $wb.SelectNodes('//x:sheets/x:sheet', $wnsm)) {
            $rid = $sh.GetAttribute('id', $NS_R)
            if (-not $rels.ContainsKey($rid)) { continue }
            $part = Resolve-Part 'xl/workbook.xml' $rels[$rid].target
            $text = Read-ZipText $zip $part
            if ($null -eq $text) { continue }
            $sd = New-XmlDoc $text; $snsm = New-Nsmgr $sd
            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($row in $sd.SelectNodes('//x:sheetData/x:row', $snsm)) {
                $cells = @{}
                foreach ($c in $row.SelectNodes('x:c', $snsm)) {
                    $idx = Get-ColIndex $c.GetAttribute('r')
                    $type = $c.GetAttribute('t')
                    $val = $null
                    if ($type -eq 's') {
                        $v = $c.SelectSingleNode('x:v', $snsm)
                        if ($v -and $v.InnerText -match '^\d+$') { $val = $shared[[int]$v.InnerText] }
                    } elseif ($type -eq 'inlineStr') {
                        $sb = New-Object Text.StringBuilder
                        foreach ($t in $c.SelectNodes('.//x:t', $snsm)) { [void]$sb.Append($t.InnerText) }
                        $val = $sb.ToString()
                    } else {
                        $v = $c.SelectSingleNode('x:v', $snsm)
                        if ($v) {
                            $val = $v.InnerText
                            if ($type -eq 'b') { $val = $(if ($val -eq '1') { 'True' } else { 'False' }) }
                            elseif ($type -ne 'str' -and $val -match '^-?\d+\.0+$') { $val = $val -replace '\.0+$', '' }
                        }
                    }
                    if ($null -ne $val) { $cells[$idx] = $val }
                }
                $rows.Add(@{ n = [int]$row.GetAttribute('r'); cells = $cells })
            }
            $sheets += ,@{ name = $sh.GetAttribute('name'); rows = $rows }
        }
        return ,$sheets
    } finally { $zip.Dispose() }
}

function Import-Log([string]$xlsx) {
    $sheets = Read-Sheets $xlsx
    $best = $null; $bestN = 0
    foreach ($sh in $sheets) {
        $hdrRow = $sh.rows | Where-Object { $_.n -eq 1 } | Select-Object -First 1
        if ($null -eq $hdrRow) { continue }
        $hit = $false
        foreach ($k in $hdrRow.cells.Keys) { if (([string]$hdrRow.cells[$k]).Trim().ToLowerInvariant().StartsWith('alt text - short')) { $hit = $true } }
        $maxRow = 0; foreach ($r in $sh.rows) { if ($r.n -gt $maxRow) { $maxRow = $r.n } }
        if ($hit -and $maxRow -gt $bestN -and $sh.name.ToLowerInvariant() -ne 'examples') { $best = $sh; $bestN = $maxRow }
    }
    if ($null -eq $best) { throw "Could not find a sheet with an 'Alt Text - Short' column." }
    $hdrRow = $best.rows | Where-Object { $_.n -eq 1 } | Select-Object -First 1
    $hdrNames = New-Object System.Collections.Generic.List[string]; $hdrIdx = @{}
    foreach ($k in ($hdrRow.cells.Keys | Sort-Object)) {
        $h = ([string]$hdrRow.cells[$k]).Trim().ToLowerInvariant()
        if (-not $hdrIdx.ContainsKey($h)) { $hdrNames.Add($h) }
        $hdrIdx[$h] = [int]$k
    }
    function Find-Col([string[]]$names) {
        foreach ($n in $names) { foreach ($h in $hdrNames) { if ($h.StartsWith($n)) { return $hdrIdx[$h] } } }
        return -1
    }
    $ci = @{
        prefix = (Find-Col @('chapter prefix')); ch = (Find-Col @('chapter number')); el = (Find-Col @('element number'))
        usage = (Find-Col @('usage')); dec = (Find-Col @('decorative')); short = (Find-Col @('alt text - short'))
        long = (Find-Col @('alt text - long')); fig = (Find-Col @('figure naming')); page = (Find-Col @('page no'))
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in ($best.rows | Where-Object { $_.n -ge 2 } | Sort-Object { $_.n })) {
        $g = { param($k) if ($ci[$k] -ge 0 -and $r.cells.ContainsKey($ci[$k])) { ([string]$r.cells[$ci[$k]]).Trim() } else { '' } }
        $short = & $g 'short'; $el = & $g 'el'
        if (-not ($short -or $el)) { continue }
        $ch = & $g 'ch'
        if ($ch -match '^\d+$') { $ch = $ch.PadLeft(2, '0') }
        $rows.Add([pscustomobject]@{
            prefix = (& $g 'prefix'); ch = $ch; el = $el; usage = (& $g 'usage')
            dec = (@('yes', 'y') -contains (& $g 'dec').ToLowerInvariant())
            short = $short; long = (& $g 'long'); fig = (& $g 'fig'); page = (& $g 'page'); key = ''
        })
    }
    return ,$rows
}

# ----------------------------------------------------------------- PPTX reading
$script:BoxCache = @{}
function Get-LayoutMaster($zip, [string]$slidePart) {
    $rels = Get-Rels $zip $slidePart
    $layout = $null
    foreach ($v in $rels.Values) { if ($v.type -match 'slideLayout$') { $layout = Resolve-Part $slidePart $v.target; break } }
    $master = $null
    if ($layout) {
        $lr = Get-Rels $zip $layout
        foreach ($v in $lr.Values) { if ($v.type -match 'slideMaster$') { $master = Resolve-Part $layout $v.target; break } }
    }
    return @($layout, $master)
}
function Get-XfrmBox($node, $nsm) {
    $xf = $node.SelectSingleNode('p:spPr/a:xfrm', $nsm)
    if ($null -eq $xf) { $xf = $node.SelectSingleNode('p:grpSpPr/a:xfrm', $nsm) }
    if ($null -eq $xf) { return $null }
    $off = $xf.SelectSingleNode('a:off', $nsm); $ext = $xf.SelectSingleNode('a:ext', $nsm)
    if ($null -eq $off -or $null -eq $ext) { return $null }
    return @{ x = [double]$off.GetAttribute('x'); y = [double]$off.GetAttribute('y'); cx = [double]$ext.GetAttribute('cx'); cy = [double]$ext.GetAttribute('cy') }
}
function Find-PhBox($zip, [string]$part, [string]$type, [string]$idx) {
    if (-not $part) { return $null }
    $key = "$part|$type|$idx"
    if ($script:BoxCache.ContainsKey($key)) { return $script:BoxCache[$key] }
    $text = Read-ZipText $zip $part
    $box = $null
    if ($text) {
        $d = New-XmlDoc $text; $nsm = New-Nsmgr $d
        foreach ($sp in $d.SelectNodes('//p:sp', $nsm)) {
            $ph = $sp.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $nsm)
            if ($null -eq $ph) { continue }
            $pi = $ph.GetAttribute('idx'); $pt = $ph.GetAttribute('type')
            if (($idx -and $pi -eq $idx) -or ((-not $idx) -and $type -and $pt -eq $type)) { $box = Get-XfrmBox $sp $nsm; if ($box) { break } }
        }
    }
    $script:BoxCache[$key] = $box
    return $box
}
function Get-ShapeBox($zip, [string]$slidePart, $node, $nsm) {
    $b = Get-XfrmBox $node $nsm
    if ($b) { return $b }
    $ph = $node.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $nsm)
    if ($null -eq $ph) { $ph = $node.SelectSingleNode('p:nvPicPr/p:nvPr/p:ph', $nsm) }
    if ($null -eq $ph) { return $null }
    $lm = Get-LayoutMaster $zip $slidePart
    $type = $ph.GetAttribute('type'); $idx = $ph.GetAttribute('idx')
    $b = Find-PhBox $zip $lm[0] $type $idx
    if ($b) { return $b }
    if (-not $type) { $type = 'body' }
    return (Find-PhBox $zip $lm[1] $type '')
}
function Get-ParaText($node, $nsm) {
    $paras = @()
    foreach ($p in $node.SelectNodes('.//a:p', $nsm)) {
        $sb = New-Object Text.StringBuilder
        foreach ($c in $p.ChildNodes) {
            if ($c.LocalName -eq 'r' -or $c.LocalName -eq 'fld') { foreach ($t in $c.SelectNodes('a:t', $nsm)) { [void]$sb.Append($t.InnerText) } }
            elseif ($c.LocalName -eq 'br') { [void]$sb.Append("`n") }
        }
        $paras += $sb.ToString()
    }
    return ($paras -join "`n").Trim()
}
function Get-Tokens([string]$text) {
    $t = ([string]$text).ToLowerInvariant()
    foreach ($k in $ABBR.Keys) { $t = [regex]::Replace($t, '\b' + $k + '\b', $ABBR[$k]) }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($t, '[a-z0-9]+')) {
        $w = $m.Value
        if ($STOP.ContainsKey($w)) { continue }
        if ($w.Length -gt 3 -and $w.EndsWith('s')) { $w = $w.Substring(0, $w.Length - 1) }
        $out.Add($w)
    }
    return ,$out
}

function Import-Deck([string]$pptx) {
    $zip = [IO.Compression.ZipFile]::OpenRead($pptx)
    try {
        $pres = New-XmlDoc (Read-ZipText $zip 'ppt/presentation.xml'); $pnsm = New-Nsmgr $pres
        $prels = Get-Rels $zip 'ppt/presentation.xml'
        $slideParts = @()
        foreach ($sid in $pres.SelectNodes('//p:sldIdLst/p:sldId', $pnsm)) {
            $slideParts += (Resolve-Part 'ppt/presentation.xml' $prels[$sid.GetAttribute('id', $NS_R)].target)
        }
        $pics = New-Object System.Collections.Generic.List[object]
        $docs = @{}
        $n = 0
        foreach ($part in $slideParts) {
            $n++
            $doc = New-XmlDoc (Read-ZipText $zip $part); $nsm = New-Nsmgr $doc
            $docs[$part] = $doc
            $rels = Get-Rels $zip $part
            $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
            # title + other text shapes (top level only, like the earlier tool)
            $title = ''; $texts = New-Object System.Collections.Generic.List[object]
            foreach ($e in $spTree.ChildNodes) {
                if ($e.NodeType -ne 'Element' -or $e.LocalName -ne 'sp') { continue }
                if ($null -eq $e.SelectSingleNode('p:txBody', $nsm)) { continue }
                $t = Get-ParaText $e $nsm
                if (-not $t) { continue }
                $ph = $e.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $nsm)
                $isTitle = $ph -and (@('title', 'ctrTitle') -contains $ph.GetAttribute('type'))
                if ($isTitle) { if (-not $title) { $title = $t } }
                elseif ($t -notmatch '^(?i)(\u00A9|courtesy|top |bottom |left |right |source)') { $texts.Add(@{ node = $e; text = $t }) }
            }
            $allText = (($texts | ForEach-Object { $_.text }) -join ' ')
            $stack = New-Object System.Collections.Generic.Stack[object]
            $order = New-Object System.Collections.Generic.List[object]
            function Walk($parent) {
                foreach ($e in $parent.ChildNodes) {
                    if ($e.NodeType -ne 'Element') { continue }
                    if ($e.LocalName -eq 'pic') { $order.Add($e) }
                    elseif ($e.LocalName -eq 'grpSp') { Walk $e }
                }
            }
            Walk $spTree
            foreach ($pic in $order) {
                $pr = $pic.SelectSingleNode('p:nvPicPr/p:cNvPr', $nsm)
                $near = ''
                $pb = Get-ShapeBox $zip $part $pic $nsm
                if ($pb) {
                    $cx = $pb.x + $pb.cx / 2; $cy = $pb.y + $pb.cy / 2; $bd = $null
                    foreach ($tx in $texts) {
                        $tb = Get-ShapeBox $zip $part $tx.node $nsm
                        if ($null -eq $tb) { continue }
                        $dd = [math]::Sqrt([math]::Pow($tb.x + $tb.cx / 2 - $cx, 2) + [math]::Pow($tb.y + $tb.cy / 2 - $cy, 2))
                        if ($null -eq $bd -or $dd -lt $bd) { $near = $tx.text; $bd = $dd }
                    }
                }
                $blob = $null; $mime = ''
                $blip = $pic.SelectSingleNode('p:blipFill/a:blip', $nsm)
                if ($blip) {
                    $rid = $blip.GetAttribute('embed', $NS_R)
                    if ($rid -and $rels.ContainsKey($rid)) {
                        $mp = Resolve-Part $part $rels[$rid].target
                        $blob = Read-ZipBytes $zip $mp
                        $ext = [IO.Path]::GetExtension($mp).ToLowerInvariant()
                        $mime = switch ($ext) { '.png' { 'image/png' } '.jpg' { 'image/jpeg' } '.jpeg' { 'image/jpeg' } '.gif' { 'image/gif' } '.bmp' { 'image/bmp' } '.svg' { 'image/svg+xml' } '.webp' { 'image/webp' } default { '' } }
                    }
                }
                $decor = ($null -ne $pic.SelectSingleNode('.//*[local-name()="decorative"]'))
                $pics.Add([pscustomobject]@{
                    slide = $n; part = $part; node = $pic; id = [int]$pr.GetAttribute('id'); name = $pr.GetAttribute('name')
                    alt = ([string]$pr.GetAttribute('descr')).Trim(); decor = $decor; title = $title; near = $near
                    allText = $allText; blob = $blob; mime = $mime
                })
            }
        }
        return @{ pics = $pics; docs = $docs; slides = $slideParts.Count }
    } finally { $zip.Dispose() }
}

# ----------------------------------------------------------------- matching (same method as the earlier tool)
function Invoke-Match($pics, $rows) {
    $cands = @($rows | Where-Object { $_.short })
    $seen = @{}
    foreach ($r in $cands) {
        $k = if ($r.el) { $r.el } else { '?' }
        if ($seen.ContainsKey($k)) { $seen[$k]++ } else { $seen[$k] = 1 }
        $r.key = if ($seen[$k] -eq 1) { $k } else { "$k #$($seen[$k])" }
    }
    $docs = @()
    foreach ($r in $cands) {
        $lg = $r.long; if ($lg.Length -gt 300) { $lg = $lg.Substring(0, 300) }
        $h = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($w in (Get-Tokens ($r.short + ' ' + $lg + ' ' + $r.el))) { [void]$h.Add($w) }
        $docs += ,$h
    }
    $n = $docs.Count
    $df = @{}
    foreach ($d in $docs) { foreach ($w in $d) { if ($df.ContainsKey($w)) { $df[$w]++ } else { $df[$w] = 1 } } }
    $idf = @{}
    foreach ($w in $df.Keys) { $idf[$w] = [math]::Log(($n + 1) / ($df[$w] + 0.5)) }
    $pairs = New-Object System.Collections.Generic.List[object]
    for ($pi = 0; $pi -lt $pics.Count; $pi++) {
        $p = $pics[$pi]
        $weights = [ordered]@{}
        foreach ($pair in @(@($p.near, 3), @($p.title, 1.5), @($p.allText, 1), @($p.alt, 4))) {
            $set = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($w in (Get-Tokens $pair[0])) { [void]$set.Add($w) }
            foreach ($w in $set) { if (-not $weights.Contains($w) -or $weights[$w] -lt $pair[1]) { $weights[$w] = [double]$pair[1] } }
        }
        $tot = 0.0
        foreach ($w in $weights.Keys) { $tot += $weights[$w] * $(if ($df.ContainsKey($w)) { $idf[$w] } else { 0.6 }) }
        if ($tot -eq 0) { $tot = 1.0 }
        $sc = New-Object 'double[]' $n
        for ($ri = 0; $ri -lt $n; $ri++) {
            $hit = 0.0
            foreach ($w in $weights.Keys) { if ($docs[$ri].Contains($w)) { $hit += $weights[$w] * $idf[$w] } }
            $sc[$ri] = $hit / $tot
        }
        $order = @(0..($n - 1) | Sort-Object @{ Expression = { - $sc[$_] } }, @{ Expression = { $_ } })
        if ($n -eq 0) { $order = @() }
        $best = if ($order.Count -gt 0) { $sc[$order[0]] } else { 0.0 }
        $second = if ($order.Count -gt 1) { $sc[$order[1]] } else { 0.0 }
        $top = [math]::Min(4, $order.Count)
        for ($k = 0; $k -lt $top; $k++) {
            $ri = $order[$k]
            $pairs.Add([pscustomobject]@{ sc = $sc[$ri]; pi = $pi; ri = $ri; margin = $(if ($ri -eq $order[0]) { $best - $second } else { 0.0 }) })
        }
    }
    $sorted = @($pairs | Sort-Object @{ Expression = { $_.sc }; Descending = $true }, @{ Expression = { $_.pi }; Descending = $true }, @{ Expression = { $_.ri }; Descending = $true }, @{ Expression = { $_.margin }; Descending = $true })
    $usedP = @{}; $usedR = @{}; $res = @{}
    foreach ($e in $sorted) {
        if ($usedP.ContainsKey($e.pi) -or $usedR.ContainsKey($e.ri)) { continue }
        $res[$e.pi] = @{ row = $cands[$e.ri]; score = $e.sc; margin = $e.margin }
        $usedP[$e.pi] = $true; $usedR[$e.ri] = $true
    }
    return @{ res = $res; cands = $cands }
}

# ----------------------------------------------------------------- writing the alt text
function Set-Decorative($pic, $nsm, [Xml.XmlDocument]$doc) {
    $pr = $pic.SelectSingleNode('p:nvPicPr/p:cNvPr', $nsm)
    [void]$pr.SetAttribute('descr', '')
    Remove-DecorativeMark $pic $nsm
    $extLst = $pr.SelectSingleNode('a:extLst', $nsm)
    if ($null -eq $extLst) { $extLst = $doc.CreateElement('a', 'extLst', $NS_A); [void]$pr.AppendChild($extLst) }
    $ext = $doc.CreateElement('a', 'ext', $NS_A); [void]$ext.SetAttribute('uri', $DECOR_URI)
    $dec = $doc.CreateElement('adec', 'decorative', $DECOR_NS); [void]$dec.SetAttribute('val', '1')
    [void]$ext.AppendChild($dec); [void]$extLst.AppendChild($ext)
}
function Remove-DecorativeMark($pic, $nsm) {
    $pr = $pic.SelectSingleNode('p:nvPicPr/p:cNvPr', $nsm)
    $extLst = $pr.SelectSingleNode('a:extLst', $nsm)
    if ($null -eq $extLst) { return }
    foreach ($ext in @($extLst.SelectNodes('a:ext', $nsm))) {
        if ($ext.GetAttribute('uri') -eq $DECOR_URI) { [void]$extLst.RemoveChild($ext) }
    }
    if (-not $extLst.HasChildNodes) { [void]$pr.RemoveChild($extLst) }
}

function Save-WithAlt([string]$pptxIn, [string]$pptxOut, $pics, $docs, $edits) {
    Copy-Item -LiteralPath $pptxIn -Destination $pptxOut -Force
    $done = 0; $decor = 0; $skipped = 0; $touched = @{}
    foreach ($ed in $edits) {
        $i = [int]$ed.i
        if ($i -lt 0 -or $i -ge $pics.Count) { continue }
        $p = $pics[$i]
        $doc = $docs[$p.part]; $nsm = New-Nsmgr $doc
        $pr = $p.node.SelectSingleNode('p:nvPicPr/p:cNvPr', $nsm)
        $txt = ([string]$ed.text -replace '\s+', ' ').Trim()
        if ($ed.dec) { Set-Decorative $p.node $nsm $doc; $decor++; $touched[$p.part] = $true }
        elseif ($txt) {
            Remove-DecorativeMark $p.node $nsm
            [void]$pr.SetAttribute('descr', $txt); $done++; $touched[$p.part] = $true
        } else { $skipped++ }
    }
    $zip = [IO.Compression.ZipFile]::Open($pptxOut, 'Update')
    try {
        foreach ($part in $touched.Keys) {
            $e = $zip.GetEntry($part); if ($e) { $e.Delete() }
            $ne = $zip.CreateEntry($part)
            $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
            try { $sw.Write($docs[$part].OuterXml) } finally { $sw.Dispose() }
        }
    } finally { $zip.Dispose() }
    $chk = [IO.Compression.ZipFile]::OpenRead($pptxOut); $chk.Dispose()
    return @{ done = $done; decor = $decor; skipped = $skipped }
}

# ----------------------------------------------------------------- review page
function Build-Page($items, $list, [string]$deckName, [string]$token) {
    $itemJson = '[' + (($items | ForEach-Object { $_ | ConvertTo-Json -Compress }) -join ',') + ']'
    $listJson = '[' + (($list | ForEach-Object { $_ | ConvertTo-Json -Compress }) -join ',') + ']'
    $itemJson = $itemJson.Replace('</', '<\/'); $listJson = $listJson.Replace('</', '<\/')
    $html = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Alt text review</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--g:#e2efda;--a:#fff2cc;--r:#f8cbad;--ink:#1b1f24;--mut:#5b6573;--line:#d5dbe3;--pri:#1f4e78}
*{box-sizing:border-box}body{margin:0;font:15px/1.45 "Segoe UI",system-ui,sans-serif;color:var(--ink);background:#f4f6f9}
header{position:sticky;top:0;z-index:5;background:var(--pri);color:#fff;padding:10px 20px;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
header h1{font-size:18px;margin:0;font-weight:600}header .sub{opacity:.85;font-size:13px}
.bar{margin-left:auto;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
button{font:inherit;border:1px solid var(--line);background:#fff;border-radius:6px;padding:6px 12px;cursor:pointer}
button.pri{background:#2e7d32;color:#fff;border-color:#2e7d32;font-weight:600}
button.f.on{background:#e7eefc;border-color:#7a9ad6}
main{max-width:1180px;margin:16px auto;padding:0 16px 90px}
.card{display:grid;grid-template-columns:230px 1fr;gap:16px;background:#fff;border:1px solid var(--line);border-left:8px solid var(--r);border-radius:10px;padding:14px;margin-bottom:14px}
.card.green{border-left-color:#70ad47}.card.amber{border-left-color:#ffc000}.card.red{border-left-color:#e07a5f}
.thumb{background:#f0f2f5;border-radius:8px;display:flex;align-items:center;justify-content:center;min-height:130px;overflow:hidden}
.thumb img{max-width:100%;max-height:200px;display:block}
.meta{color:var(--mut);font-size:13px}.meta b{color:var(--ink)}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin-top:8px}
select,textarea,input[type=text]{font:inherit;border:1px solid var(--line);border-radius:6px;padding:6px 8px;width:100%}
select{max-width:520px}textarea{min-height:54px;resize:vertical}
label{font-size:13px}.res{margin-top:8px;padding:8px 10px;border-radius:6px;background:#f6f8fb;border:1px dashed var(--line);white-space:pre-wrap}
.res.empty{color:#9a3412;background:#fff4ec}.tag{display:inline-block;font-size:12px;padding:1px 8px;border-radius:10px;margin-right:6px}
.tg{background:var(--g)}.ta{background:var(--a)}.tr{background:var(--r)}
#done{display:none;position:fixed;left:0;right:0;bottom:0;background:#e8f5e9;border-top:2px solid #2e7d32;padding:14px 20px;z-index:9}
.hide{display:none!important}
@media(max-width:800px){.card{grid-template-columns:1fr}}
</style></head><body>
<header><div><h1>Alt text review</h1><div class="sub">__DECK__</div></div>
<div class="bar"><span id="count"></span>
<button class="f on" data-f="all">All</button><button class="f" data-f="red">Needs a pick</button><button class="f" data-f="amber">Check</button><button class="f" data-f="green">Confident</button>
<button class="pri" id="apply">Apply and save copy</button></div></header>
<main id="main"></main>
<div id="done"></div>
<script>
var ITEMS=__ITEMS__, LIST=__LIST__, TOKEN="__TOKEN__";
var byKey={};LIST.forEach(function(r){byKey[r.key]=r;});
var main=document.getElementById('main');
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
var opts='<option value="">- pick an Element ID -</option>'+LIST.map(function(r){var t=r.short.length>70?r.short.slice(0,70)+'...':r.short;return '<option value="'+esc(r.key)+'">'+esc(r.key)+' - '+esc(t)+'</option>';}).join('');
function build(){
 ITEMS.forEach(function(it){
  var c=document.createElement('div');c.className='card '+it.level;c.id='c'+it.i;c.dataset.level=it.level;
  var tag=it.level==='green'?'<span class="tag tg">Confident</span>':it.level==='amber'?'<span class="tag ta">Check the picture</span>':'<span class="tag tr">Pick the Element ID</span>';
  var sugg=it.sugg?'Suggested: <b>'+esc(it.sugg)+'</b> (score '+it.score+') '+(it.eid?'':'<a href="#" data-use="'+it.i+'">use it</a>'):'';
  c.innerHTML='<div class="thumb">'+(it.img?'<img src="/img/'+it.i+'?t='+TOKEN+'" alt="" onerror="this.parentNode.innerHTML=\'(no preview for this picture type)\'">':'(no preview)')+'</div>'+
  '<div><div>'+tag+'<b>Slide '+it.slide+'</b> &middot; '+esc(it.name)+'</div>'+
  '<div class="meta">'+(it.title?'<b>Title:</b> '+esc(it.title)+'<br>':'')+(it.near?'<b>Nearby text:</b> '+esc(it.near.slice(0,200))+'<br>':'')+(it.alt?'<b>Existing alt text:</b> '+esc(it.alt)+'<br>':'<b>Existing alt text:</b> (none)<br>')+sugg+'</div>'+
  '<div class="row"><select id="e'+it.i+'">'+opts+'</select></div>'+
  '<div class="row"><textarea id="o'+it.i+'" placeholder="Or type your own alt text here (this overrides the Element ID)"></textarea></div>'+
  '<div class="row"><label><input type="checkbox" id="d'+it.i+'"> Decorative image</label><label><input type="checkbox" id="a'+it.i+'"> Also add the long description</label><span class="meta" id="n'+it.i+'"></span></div>'+
  '<div class="res" id="r'+it.i+'"></div></div>';
  main.appendChild(c);
  document.getElementById('e'+it.i).value=it.eid||'';
  document.getElementById('d'+it.i).checked=!!it.decor;
  ['e','o','d','a'].forEach(function(p){var el=document.getElementById(p+it.i);el.addEventListener('input',function(){upd(it.i);});el.addEventListener('change',function(){upd(it.i);});});
  upd(it.i);
 });
 main.addEventListener('click',function(ev){var u=ev.target.getAttribute&&ev.target.getAttribute('data-use');if(u!==null&&u!==undefined){ev.preventDefault();var it=ITEMS[+u];document.getElementById('e'+u).value=it.sugg;upd(+u);}});
}
function resultOf(i){
 if(!document.getElementById('d'+i))return {dec:false,text:'x',eid:''};
 var dec=document.getElementById('d'+i).checked,own=document.getElementById('o'+i).value.trim(),k=document.getElementById('e'+i).value,app=document.getElementById('a'+i).checked,txt=own;
 if(!txt&&byKey[k]){txt=byKey[k].short;if(app&&byKey[k].long)txt=(txt+' '+byKey[k].long).trim();}
 return {dec:dec,text:txt.replace(/\s+/g,' ').trim(),eid:k};
}
function upd(i){
 var r=resultOf(i),el=document.getElementById('r'+i),n=document.getElementById('n'+i);
 if(r.dec){el.textContent='(marked decorative - screen readers skip it)';el.className='res';}
 else if(r.text){el.textContent=r.text;el.className='res';}
 else{el.textContent='(nothing will be written - the picture keeps its current alt text)';el.className='res empty';}
 n.textContent=(byKey[r.eid]&&byKey[r.eid].dec)?'The log marks this one as decorative.':'';
 var left=ITEMS.filter(function(it){var x=resultOf(it.i);return !x.dec&&!x.text;}).length;
 document.getElementById('count').textContent=ITEMS.length+' pictures - '+left+' still empty';
}
document.querySelectorAll('button.f').forEach(function(b){b.addEventListener('click',function(){
 document.querySelectorAll('button.f').forEach(function(x){x.classList.remove('on');});b.classList.add('on');
 var f=b.dataset.f;document.querySelectorAll('.card').forEach(function(c){c.classList.toggle('hide',f!=='all'&&c.dataset.level!==f);});});});
document.getElementById('apply').addEventListener('click',function(){
 var edits=ITEMS.map(function(it){var r=resultOf(it.i);return {i:it.i,dec:r.dec,text:r.text};});
 var left=edits.filter(function(e){return !e.dec&&!e.text;}).length;
 if(left&&!confirm(left+' picture(s) will be left as they are (no new alt text). Save anyway?'))return;
 var b=this;b.disabled=true;b.textContent='Saving...';
 fetch('/apply',{method:'POST',headers:{'Content-Type':'application/json','X-Token':TOKEN},body:JSON.stringify({edits:edits})})
 .then(function(r){return r.json();}).then(function(j){
  b.disabled=false;b.textContent='Apply and save copy';var d=document.getElementById('done');d.style.display='block';
  if(j.error){d.style.background='#fdecea';d.style.borderColor='#c62828';d.textContent='Could not save: '+j.error;return;}
  d.innerHTML='<b>Saved.</b> Alt text written: '+j.done+' &middot; marked decorative: '+j.decor+' &middot; left as they were: '+j.skipped+'<br>New file: '+esc(j.out)+' &nbsp; <button id="open">Show in folder</button> <button id="fin">Finish and close</button>';
  document.getElementById('open').onclick=function(){fetch('/open',{method:'POST',headers:{'X-Token':TOKEN}});};
  document.getElementById('fin').onclick=function(){fetch('/quit',{method:'POST',headers:{'X-Token':TOKEN}}).then(function(){document.body.innerHTML='<p style="padding:30px;font-size:18px">Done. You can close this tab and the command window.</p>';});};
 }).catch(function(e){b.disabled=false;b.textContent='Apply and save copy';alert('Could not reach the tool: '+e);});
});
build();upd(0);
</script></body></html>
'@
    return $html.Replace('__ITEMS__', $itemJson).Replace('__LIST__', $listJson).Replace('__TOKEN__', $token).Replace('__DECK__', [Net.WebUtility]::HtmlEncode($deckName))
}

function Start-ReviewServer([string]$html, [string]$token, $pics, [string]$pptx, $docs, [string]$out) {
    $port = 0; $listener = $null
    foreach ($try in 1..30) {
        $port = Get-Random -Minimum 49200 -Maximum 65000
        $l = New-Object Net.HttpListener
        $l.Prefixes.Add("http://127.0.0.1:$port/")
        try { $l.Start(); $listener = $l; break } catch { try { $l.Close() } catch { } }
    }
    if ($null -eq $listener) { throw 'Could not start the local review page (no free port).' }
    $url = "http://127.0.0.1:$port/"
    if ($env:ALT_PORTFILE) { Set-Content -LiteralPath $env:ALT_PORTFILE -Value "$port|$token" -Encoding ASCII } else { Start-Process $url }
    Write-Host ''
    Write-Host 'The review page is open in your browser.' -ForegroundColor Green
    Write-Host ('If it did not open, go to: ' + $url)
    Write-Host 'Check the pictures, then press "Apply and save copy". Keep this window open until you finish.'
    $enc = New-Object Text.UTF8Encoding($false)
    $quit = $false
    try {
        while (-not $quit -and $listener.IsListening) {
            $ctx = $listener.GetContext()
            $req = $ctx.Request; $resp = $ctx.Response
            try {
                $path = $req.Url.AbsolutePath
                $okTok = ($req.Headers['X-Token'] -eq $token) -or ($req.QueryString['t'] -eq $token)
                if ($req.HttpMethod -eq 'GET' -and $path -eq '/') {
                    $b = $enc.GetBytes($html); $resp.ContentType = 'text/html; charset=utf-8'; $resp.OutputStream.Write($b, 0, $b.Length)
                } elseif ($req.HttpMethod -eq 'GET' -and $path -like '/img/*' -and $okTok) {
                    $i = [int]($path.Substring(5))
                    $p = $pics[$i]
                    if ($p.blob -and $p.mime) { $resp.ContentType = $p.mime; $resp.OutputStream.Write($p.blob, 0, $p.blob.Length) } else { $resp.StatusCode = 404 }
                } elseif ($req.HttpMethod -eq 'POST' -and $okTok -and $path -eq '/apply') {
                    $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8); $body = $sr.ReadToEnd(); $sr.Close()
                    try {
                        $data = $body | ConvertFrom-Json
                        $edits = @(); foreach ($e in @($data.edits)) { $edits += ,$e }
                        $r = Save-WithAlt $pptx $out $pics $docs $edits
                        $json = '{"done":' + $r.done + ',"decor":' + $r.decor + ',"skipped":' + $r.skipped + ',"out":' + (ConvertTo-Json ([string]$out) -Compress) + '}'
                        Write-Host ('Saved: ' + $out + '  (written ' + $r.done + ', decorative ' + $r.decor + ', left as they were ' + $r.skipped + ')') -ForegroundColor Green
                    } catch { $json = '{"error":' + (ConvertTo-Json ([string]$_.Exception.Message) -Compress) + '}' }
                    $b = $enc.GetBytes($json); $resp.ContentType = 'application/json'; $resp.OutputStream.Write($b, 0, $b.Length)
                } elseif ($req.HttpMethod -eq 'POST' -and $okTok -and $path -eq '/open') {
                    if (Test-Path -LiteralPath $out) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $out + '"') }
                } elseif ($req.HttpMethod -eq 'POST' -and $okTok -and $path -eq '/quit') {
                    $quit = $true
                } else { $resp.StatusCode = 404 }
            } catch { try { $resp.StatusCode = 500 } catch { } }
            finally { try { $resp.Close() } catch { } }
        }
    } finally { try { $listener.Stop(); $listener.Close() } catch { } }
}

# ----------------------------------------------------------------- main
Write-Host ('PPT Alt-Text Filler - running version ' + $ALT_VERSION) -ForegroundColor Cyan
Write-Host ''
try {
    $pptx = $env:ALT_PPTX; $xlsx = $env:ALT_XLSX
    if (-not $pptx -and $args.Count -ge 1) { $pptx = [string]$args[0] }
    if (-not $xlsx -and $args.Count -ge 2) { $xlsx = [string]$args[1] }
    function Pick-One([string]$title, [string]$filter) {
        Add-Type -AssemblyName System.Windows.Forms
        $owner = New-Object System.Windows.Forms.Form; $owner.TopMost = $true
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = $title; $dlg.Filter = $filter
        if ($dlg.ShowDialog($owner) -eq 'OK') { return $dlg.FileName } else { return $null }
    }
    if (-not $pptx) { Write-Host 'Choose the PowerPoint file...'; $pptx = Pick-One 'Select the PowerPoint file (.pptx)' 'PowerPoint (*.pptx)|*.pptx' }
    if (-not $pptx) { Write-Host 'No file selected.'; return }
    if (-not $xlsx) { Write-Host 'Choose the Alt-Text Excel file...'; $xlsx = Pick-One 'Select the Alt-Text Excel log (.xlsx)' 'Excel (*.xlsx)|*.xlsx' }
    if (-not $xlsx) { Write-Host 'No Excel file selected.'; return }
    $pptx = [IO.Path]::GetFullPath($pptx.Trim().Trim('"')); $xlsx = [IO.Path]::GetFullPath($xlsx.Trim().Trim('"'))

    $ch = $env:ALT_CHAPTER
    if (-not $ch) {
        $ch = Get-Chapter $pptx
        if (-not $env:ALT_NOPROMPT) {
            $x = (Read-Host ('Chapter number to use [' + $(if ($ch) { $ch } else { 'all' }) + '] (Enter to accept, or type e.g. 05, or "all")')).Trim()
            if ($x) { if ($x -match '^\d+$') { $ch = $x.PadLeft(2, '0') } elseif ($x.ToLowerInvariant() -eq 'all') { $ch = $null } }
        }
    } elseif ($ch.ToLowerInvariant() -eq 'all') { $ch = $null }

    Write-Host 'Reading the Excel log...'
    $rows = Import-Log $xlsx
    if ($ch) {
        $sub = @($rows | Where-Object { $_.ch -eq $ch -or $_.ch.TrimStart('0') -eq $ch.TrimStart('0') })
        Write-Host ("Chapter $ch" + ': ' + $sub.Count + ' Excel rows (of ' + $rows.Count + ' total)')
        if ($sub.Count -gt 0) { $rows = $sub }
    }
    Write-Host 'Reading the PowerPoint...'
    $deck = Import-Deck $pptx
    $pics = $deck.pics
    Write-Host ('Found ' + $pics.Count + ' pictures in ' + $deck.slides + ' slides.')
    if ($pics.Count -eq 0) { Write-Host 'No pictures in this deck - nothing to do.'; return }

    $m = Invoke-Match $pics $rows
    $items = New-Object System.Collections.Generic.List[object]
    $auto = 0
    for ($i = 0; $i -lt $pics.Count; $i++) {
        $p = $pics[$i]; $r = $m.res[$i]
        $isAuto = ($null -ne $r) -and ($r.score -ge $AUTO_THRESHOLD) -and ($r.margin -ge $MARGIN)
        if ($isAuto) { $auto++ }
        $level = if ($isAuto -and $r.score -ge 0.8) { 'green' } elseif ($isAuto) { 'amber' } else { 'red' }
        $items.Add([pscustomobject]@{
            i = $i; slide = $p.slide; id = $p.id; name = $p.name; alt = $p.alt; decor = $p.decor; title = $p.title; near = $p.near
            sugg = $(if ($r) { $r.row.key } else { '' }); score = $(if ($r) { [math]::Round($r.score, 2) } else { 0 })
            eid = $(if ($isAuto) { $r.row.key } else { '' }); level = $level; img = [bool]($p.blob -and $p.mime)
        })
    }
    Write-Host ("Auto-matched $auto/" + $pics.Count + '; the rest need a manual look.')
    $list = @($m.cands | ForEach-Object { [pscustomobject]@{ key = $_.key; page = $_.page; usage = $_.usage; short = $_.short; long = $_.long; dec = [bool]$_.dec } })
    $token = [guid]::NewGuid().ToString('N')
    $html = Build-Page $items $list ([IO.Path]::GetFileName($pptx)) $token
    $out = [IO.Path]::Combine([IO.Path]::GetDirectoryName($pptx), [IO.Path]::GetFileNameWithoutExtension($pptx) + '_with_alt.pptx')
    if ($env:ALT_DUMP) { Set-Content -LiteralPath $env:ALT_DUMP -Value ($items | ConvertTo-Json -Depth 4) -Encoding UTF8 }
    if ($env:ALT_NOSERVE) { Set-Content -LiteralPath ($out + '.review.html') -Value $html -Encoding UTF8; return }
    Start-ReviewServer $html $token $pics $pptx $deck.docs $out
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ('  (version ' + $ALT_VERSION + ', at line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + ($_.InvocationInfo.Line + '').Trim() + ')') -ForegroundColor DarkGray
}
