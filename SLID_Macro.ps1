# SLID Macro - numbers repeated slide titles (1 of 4), renames shapes, puts sideways photo credits under their picture (SDRS files only), fixes language tags.
# Saves a remediated copy next to the original. (Self-updates from GitHub via SLID_Update.ps1.)
# SLID MACRO SCRIPT
$SLID_VERSION = '1.8'   # shown at the start of the window; the updater compares this number with GitHub
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
function Fix-Deck {
param([string]$Path,[string]$OutPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$src = (Resolve-Path -LiteralPath $Path).Path
if (-not $OutPath) {
    $OutPath = Join-Path (Split-Path -Parent $src) ([IO.Path]::GetFileNameWithoutExtension($src) + ' Remediated.pptx')
}
if ($OutPath -eq $src) { throw 'OutPath must differ from Path so the original is kept.' }
Copy-Item -LiteralPath $src -Destination $OutPath -Force

$zip = [IO.Compression.ZipFile]::Open($OutPath, 'Update')
$report = New-Object System.Collections.Generic.List[string]

function Read-Part($name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    $x = New-Object Xml.XmlDocument
    $x.PreserveWhitespace = $true
    $x.LoadXml($sr.ReadToEnd())
    $sr.Close()
    return $x
}
function Write-Part($name, $xml) {
    $e = $zip.GetEntry($name)
    if ($e) { $e.Delete() }
    $ne = $zip.CreateEntry($name)
    $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
    $xml.Save($sw)
    $sw.Close()
}
function New-Ns($xml) {
    $m = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $m.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $m.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    $m.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $m.AddNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships')
    return ,$m
}
function Resolve-Target($basePart, $target) {
    $baseDir = [IO.Path]::GetDirectoryName($basePart.Replace('/', '\')).Replace('\', '/')
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($s in ($baseDir + '/' + $target).Split('/')) {
        if ($s -eq '' -or $s -eq '.') { continue }
        if ($s -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) } } else { $parts.Add($s) }
    }
    return ($parts -join '/')
}
function Get-RelTarget($partName, $typeSuffix) {
    $relName = [IO.Path]::GetDirectoryName($partName.Replace('/', '\')).Replace('\', '/') + '/_rels/' + [IO.Path]::GetFileName($partName) + '.rels'
    $rx = Read-Part $relName
    if (-not $rx) { return $null }
    $m = New-Ns $rx
    foreach ($r in $rx.SelectNodes('//pr:Relationship', $m)) {
        if ($r.GetAttribute('Type').EndsWith($typeSuffix)) { return (Resolve-Target $partName $r.GetAttribute('Target')) }
    }
    return $null
}

try {
    # slide order from presentation.xml
    $pres = Read-Part 'ppt/presentation.xml'
    $pns = New-Ns $pres
    $prels = Read-Part 'ppt/_rels/presentation.xml.rels'
    $prns = New-Ns $prels
    $relMap = @{}
    foreach ($r in $prels.SelectNodes('//pr:Relationship', $prns)) { $relMap[$r.GetAttribute('Id')] = (Resolve-Target 'ppt/presentation.xml' $r.GetAttribute('Target')) }
    $slideParts = New-Object System.Collections.Generic.List[string]
    foreach ($sid in $pres.SelectNodes('//p:sldIdLst/p:sldId', $pns)) {
        $rid = $sid.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $slideParts.Add($relMap[$rid])
    }

    $docs = @{}
    $nss = @{}
    $before = @{}
    foreach ($sp in $slideParts) {
        $docs[$sp] = Read-Part $sp
        $nss[$sp] = New-Ns $docs[$sp]
        $before[$sp] = (@($docs[$sp].SelectNodes('//a:t', $nss[$sp]) | ForEach-Object { $_.InnerText }) | Sort-Object) -join '|'
    }

    function Get-TitleShape($sp) {
        foreach ($s in $docs[$sp].SelectNodes('//p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph[@type="title" or @type="ctrTitle"]]', $nss[$sp])) {
            if (($s.SelectNodes('.//a:t', $nss[$sp]) | ForEach-Object { $_.InnerText }) -join '') { return $s }
        }
        return $null
    }
    # text of a shape: runs joined as they are, line breaks and paragraphs become one space
    function Get-ShapeText($shape, $m) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($p in $shape.SelectNodes('p:txBody/a:p', $m)) {
            $sb = New-Object System.Text.StringBuilder
            foreach ($c in $p.ChildNodes) {
                if ($c.LocalName -eq 'br') { [void]$sb.Append(' ') }
                elseif ($c.LocalName -eq 'r' -or $c.LocalName -eq 'fld') {
                    $t = $c.SelectSingleNode('a:t', $m)
                    if ($t) { [void]$sb.Append($t.InnerText) }
                }
            }
            $parts.Add($sb.ToString())
        }
        $t = ($parts -join ' ').Replace([string][char]160, ' ')
        return ([regex]::Replace($t, '\s+', ' ')).Trim()
    }
    function Get-TitleKey($shape, $m) { return (Get-ShapeText $shape $m).ToLowerInvariant() }

    # 1. repeated titles
    $keys = @{}
    $shapes = @{}
    $counts = @{}
    foreach ($sp in $slideParts) {
        $ts = Get-TitleShape $sp
        if ($ts) {
            $shapes[$sp] = $ts
            $k = Get-TitleKey $ts $nss[$sp]
            $keys[$sp] = $k
            if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 }
        }
    }
    $seen = @{}
    $numbered = 0
    foreach ($sp in $slideParts) {
        if (-not $shapes.ContainsKey($sp)) { continue }
        $k = $keys[$sp]
        if ($counts[$k] -le 1) { continue }
        $ts = $shapes[$sp]
        $all = @($ts.SelectNodes('.//a:t', $nss[$sp]))
        $full = (($all | ForEach-Object { $_.InnerText }) -join '').Trim()
        if ($full -match '\(\d{1,2} of \d{1,2}\)$') { continue }
        if ($seen.ContainsKey($k)) { $seen[$k]++ } else { $seen[$k] = 1 }
        $suffix = '(' + $seen[$k] + ' of ' + $counts[$k] + ')'
        $last = $all[$all.Count - 1]
        if (-not $last.InnerText.EndsWith(' ')) { $suffix = ' ' + $suffix }
        $last.InnerText = $last.InnerText + $suffix
        $numbered++
        $report.Add("Slide " + ($slideParts.IndexOf($sp) + 1) + ": " + $full + $suffix)
    }
    $report.Add("Titles numbered: $numbered")

    # 2. rename shapes
    function Get-LayoutOff($sp, $shape) {
        $m = $nss[$sp]
        $ph = $shape.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $m)
        if (-not $ph) { $ph = $shape.SelectSingleNode('p:nvPicPr/p:nvPr/p:ph', $m) }
        if (-not $ph) { return $null }
        $idx = $ph.GetAttribute('idx'); $type = $ph.GetAttribute('type')
        $layout = Get-RelTarget $sp '/slideLayout'
        $chain = @($layout)
        if ($layout) { $master = Get-RelTarget $layout '/slideMaster'; if ($master) { $chain += $master } }
        foreach ($part in $chain) {
            if (-not $part) { continue }
            $ld = Read-Part $part
            $lm = New-Ns $ld
            $cand = $null
            if ($idx) { $cand = $ld.SelectSingleNode('//p:sp[p:nvSpPr/p:nvPr/p:ph[@idx="' + $idx + '"]]', $lm) }
            if (-not $cand -and $type) { $cand = $ld.SelectSingleNode('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="' + $type + '"]]', $lm) }
            if (-not $cand -and -not $type -and -not $idx) { $cand = $ld.SelectSingleNode('//p:sp[p:nvSpPr/p:nvPr/p:ph[not(@type)]]', $lm) }
            if ($cand) {
                $off = $cand.SelectSingleNode('p:spPr/a:xfrm/a:off', $lm)
                if ($off) {
                    $ext = $cand.SelectSingleNode('p:spPr/a:xfrm/a:ext', $lm)
                    $cw = 0.0
                    if ($ext) { $cw = [double]$ext.GetAttribute('cx') }
                    return @([double]$off.GetAttribute('y'), [double]$off.GetAttribute('x'), $cw)
                }
            }
        }
        return $null
    }

    $renamed = 0
    $reordered = 0
    $slideWpt = 960.0
    $sldSz = $pres.SelectSingleNode('//p:sldSz', $pns)
    if ($sldSz) { $slideWpt = [double]$sldSz.GetAttribute('cx') / 12700 }
    $subtitleText = ''
    # a credit line is a plain text box that starts like a photo credit
    $creditRx = '^\s*(' + [char]169 + '|\(c\)|Courtesy|Source|Photo|Image|Copyright|Used with)'
    # 1b. sideways photo credits: a credit line ("(c) ... Photographed by ...") turned 90 / 270 degrees beside a picture is
    #     made upright and placed centred right under that picture; a text box that would then collide with it
    #     (the caption) moves down just enough. The reading order below then becomes: title, picture, credit, caption.
    $slideHemu = 6858000.0
    if ($sldSz) { $slideHemu = [double]$sldSz.GetAttribute('cy') }
    $creditMoved = 0; $captionMoved = 0
    function Get-Xf($node, $m) {
        $x = $node.SelectSingleNode('p:spPr/a:xfrm', $m)
        if (-not $x) { $x = $node.SelectSingleNode('p:grpSpPr/a:xfrm', $m) }
        if (-not $x) { $x = $node.SelectSingleNode('p:xfrm', $m) }
        return $x
    }
    # only for SDRS files (the file name has SDRS in it, e.g. 9781284298260_SDRS_CH02.pptx); other decks are left as they are
    $isSdrs = ([IO.Path]::GetFileName($src) -match '(?i)(^|[_\-\s\.])SDRS([_\-\s\.]|$)')
    foreach ($sp in $(if ($isSdrs) { $slideParts } else { @() })) {
        $m = $nss[$sp]
        $tree = $docs[$sp].SelectSingleNode('//p:cSld/p:spTree', $m)
        $pics = New-Object System.Collections.Generic.List[object]
        $creds = New-Object System.Collections.Generic.List[object]
        $texts = New-Object System.Collections.Generic.List[object]
        foreach ($c in $tree.ChildNodes) {
            if ($c.NodeType -ne 'Element' -or @('sp', 'pic', 'grpSp') -notcontains $c.LocalName) { continue }
            $xf = Get-Xf $c $m; if (-not $xf) { continue }
            $off = $xf.SelectSingleNode('a:off', $m); $ext = $xf.SelectSingleNode('a:ext', $m); if (-not $off -or -not $ext) { continue }
            $r = [pscustomobject]@{ Node = $c; Xf = $xf; Off = $off; Ext = $ext; L = [double]$off.GetAttribute('x'); T = [double]$off.GetAttribute('y'); W = [double]$ext.GetAttribute('cx'); H = [double]$ext.GetAttribute('cy') }
            if ($c.LocalName -eq 'pic' -or ($c.LocalName -eq 'grpSp' -and $c.SelectSingleNode('.//p:pic', $m))) { $pics.Add($r); continue }
            if ($c.LocalName -ne 'sp') { continue }
            $ph = $c.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $m)
            $txt = (($c.SelectNodes('.//a:t', $m) | ForEach-Object { $_.InnerText }) -join '')
            if ($txt.Length -eq 0) { continue }
            $rot = $xf.GetAttribute('rot')
            if (-not $ph -and $txt -match $creditRx -and ($rot -eq '5400000' -or $rot -eq '16200000')) { $creds.Add($r) }
            elseif (-not ($ph -and @('title', 'ctrTitle', 'subTitle') -contains $ph.GetAttribute('type'))) { $texts.Add($r) }
        }
        # which picture each credit stands beside: the sideways box has swapped width and height. Every picture gets ONE credit
        # (nearest first), so two credits can never land on the same picture; a credit left over takes its nearest picture.
        $cands = New-Object System.Collections.Generic.List[object]
        foreach ($cr in $creds) {
            $cxm = $cr.L + $cr.W / 2; $cym = $cr.T + $cr.H / 2
            $vl = $cxm - $cr.H / 2; $vr = $cxm + $cr.H / 2; $vt = $cym - $cr.W / 2; $vb = $cym + $cr.W / 2
            foreach ($pc in $pics) {
                $ov = [Math]::Min($vb, $pc.T + $pc.H) - [Math]::Max($vt, $pc.T)
                if ($ov -lt ($vb - $vt) * 0.5) { continue }
                $gap = 0.0
                if ($vr -lt $pc.L) { $gap = $pc.L - $vr } elseif ($vl -gt ($pc.L + $pc.W)) { $gap = $vl - ($pc.L + $pc.W) }
                if ($gap -gt 548640) { continue }
                $edge = [Math]::Min([Math]::Abs($vr - $pc.L), [Math]::Abs($vl - ($pc.L + $pc.W)))
                $cands.Add([pscustomobject]@{ Cr = $cr; Pc = $pc; Gap = $gap; Edge = $edge })
            }
        }
        $pairOf = @{}; $picUsed = @{}
        foreach ($cd in @($cands | Sort-Object Edge)) {
            $ck = $creds.IndexOf($cd.Cr); $pk = $pics.IndexOf($cd.Pc)
            if ($pairOf.ContainsKey($ck) -or $picUsed.ContainsKey($pk)) { continue }
            $pairOf[$ck] = $cd.Pc; $picUsed[$pk] = $true
        }
        foreach ($cd in @($cands | Sort-Object Gap, Edge)) { $ck = $creds.IndexOf($cd.Cr); if (-not $pairOf.ContainsKey($ck)) { $pairOf[$ck] = $cd.Pc } }
        foreach ($cr in $creds) {
            $best = $pairOf[$creds.IndexOf($cr)]
            if (-not $best) { continue }
            $nw = [Math]::Min($cr.W, $best.W)
            $nx = $best.L + ($best.W - $nw) / 2
            $ny = $best.T + $best.H + 91440
            if ($ny + $cr.H -gt $slideHemu) { $report.Add("Slide " + ($slideParts.IndexOf($sp) + 1) + ": a sideways credit line was left as it is (no room under the picture)"); continue }
            $cr.Xf.RemoveAttribute('rot')
            $cr.Off.SetAttribute('x', [string][long][Math]::Round($nx)); $cr.Off.SetAttribute('y', [string][long][Math]::Round($ny))
            $cr.Ext.SetAttribute('cx', [string][long][Math]::Round($nw))
            $nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            foreach ($para in $cr.Node.SelectNodes('p:txBody/a:p', $m)) {
                $pPr = $para.SelectSingleNode('a:pPr', $m)
                if (-not $pPr) { $pPr = $docs[$sp].CreateElement('a', 'pPr', $nsA); $null = $para.PrependChild($pPr) }
                $pPr.SetAttribute('algn', 'ctr')
            }
            $creditMoved++
            $cr.L = $nx; $cr.T = $ny; $cr.W = $nw
            # a text box right under the picture (the caption) moves down so that it does not sit on the credit
            $crBottom = $ny + $cr.H
            foreach ($tx in $texts) {
                if ($tx.Node -eq $cr.Node) { continue }
                $hov = [Math]::Min($tx.L + $tx.W, $nx + $nw) - [Math]::Max($tx.L, $nx)
                if ($hov -le 0 -or $tx.T -lt ($best.T + $best.H - 20000)) { continue }
                $want = $crBottom - 45720
                if ($tx.T -ge $want) { continue }
                $shift = $want - $tx.T
                $tx.Off.SetAttribute('y', [string][long][Math]::Round($tx.T + $shift))
                $over = ($tx.T + $shift + $tx.H) - $slideHemu
                if ($over -gt 0 -and -not $tx.Node.SelectSingleNode('p:txBody/a:bodyPr/a:spAutoFit', $m)) {
                    $tx.Ext.SetAttribute('cy', [string][long][Math]::Round([Math]::Max($tx.H - $over, $tx.H * 0.6)))
                }
                $tx.T = $tx.T + $shift; $captionMoved++
            }
        }
    }
    if ($isSdrs) { $report.Add("Sideways credit lines placed under their picture: $creditMoved (captions moved down: $captionMoved)") }
    else { $report.Add("Sideways credit lines: not changed (this step is only done for SDRS files)") }

    foreach ($sp in $slideParts) {
        $m = $nss[$sp]
        $tree = $docs[$sp].SelectSingleNode('//p:cSld/p:spTree', $m)
        $items = New-Object System.Collections.Generic.List[object]
        $order = 0
        foreach ($c in $tree.ChildNodes) {
            if ($c.NodeType -ne 'Element') { continue }
            $order++
            $kind = ''; $nv = $null; $isCredit = $false
            # a shape PowerPoint keeps inside a compatibility wrapper (mc:AlternateContent): it is renamed too, in the wrapper's two copies
            $top = $c; $nv2 = $null
            if ($c.LocalName -eq 'AlternateContent') {
                $inner = $null; $innerFb = $null
                foreach ($ch in $c.ChildNodes) {
                    if ($ch.NodeType -ne 'Element') { continue }
                    foreach ($g in $ch.ChildNodes) {
                        if ($g.NodeType -eq 'Element' -and @('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp') -contains $g.LocalName) {
                            if ($ch.LocalName -eq 'Choice' -and -not $inner) { $inner = $g } elseif ($ch.LocalName -eq 'Fallback' -and -not $innerFb) { $innerFb = $g }
                            break
                        }
                    }
                }
                if (-not $inner) { continue }
                if ($innerFb) { $nv2 = $innerFb.SelectSingleNode('*[1]/p:cNvPr', $m) }
                $c = $inner
            }
            if (@('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp') -notcontains $c.LocalName) { continue }
            switch ($c.LocalName) {
                'sp' {
                    $nv = $c.SelectSingleNode('p:nvSpPr/p:cNvPr', $m)
                    $ph = $c.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $m)
                    $pt = if ($ph) { $ph.GetAttribute('type') } else { '' }
                    $shapeText = (($c.SelectNodes('.//a:t', $m) | ForEach-Object { $_.InnerText }) -join '')
                    if ($pt -eq 'title' -or $pt -eq 'ctrTitle') { $kind = 'Title' }
                    elseif ($pt -eq 'pic') { $kind = 'Picture' }
                    elseif ($shapeText.Length -gt 0) {
                        if ($pt -eq 'subTitle') { $kind = 'Subtitle' } else { $kind = 'TextBox' }
                        if (-not $ph -and $shapeText -match $creditRx) { $isCredit = $true }
                    }
                }
                'pic' { $nv = $c.SelectSingleNode('p:nvPicPr/p:cNvPr', $m); $kind = 'Picture' }
                'grpSp' {
                    $nv = $c.SelectSingleNode('p:nvGrpSpPr/p:cNvPr', $m)
                    if ($c.SelectSingleNode('.//p:pic', $m)) { $kind = 'Picture' }
                }
            }
            # reading order group: 0 title, 1 text and other content, 2 pictures, 3 credit lines
            $rk = 1
            if ($kind -eq 'Title') { $rk = 0 } elseif ($kind -eq 'Picture') { $rk = 2 } elseif ($isCredit) { $rk = 3 }
            $off = $c.SelectSingleNode('p:spPr/a:xfrm/a:off', $m)
            if (-not $off) { $off = $c.SelectSingleNode('p:grpSpPr/a:xfrm/a:off', $m) }
            if (-not $off) { $off = $c.SelectSingleNode('p:xfrm/a:off', $m) }
            $w = 0.0; $h = 0.0
            if ($off) {
                $y = [double]$off.GetAttribute('y'); $x = [double]$off.GetAttribute('x')
                $ext = $off.ParentNode.SelectSingleNode('a:ext', $m)
                if ($ext) { $w = [double]$ext.GetAttribute('cx'); $h = [double]$ext.GetAttribute('cy') }
            }
            else {
                $lo = Get-LayoutOff $sp $c
                if ($lo) { $y = $lo[0]; $x = $lo[1]; $w = $lo[2] } else { $y = 1e12; $x = 0 }
            }
            $items.Add([pscustomobject]@{ Kind = $kind; Nv = $nv; Nv2 = $nv2; Top = $top; Node = $c; RK = $rk; Y = [Math]::Round($y / 12700); X = $x; CX = (($x + $w / 2) / 12700); Order = $order; L = $x; T = $y; W = $w; H = $h })
        }
        # first slide: its one text object is the subtitle, even when it is not a subtitle placeholder
        if ($slideParts.IndexOf($sp) -eq 0 -and -not @($items | Where-Object { $_.Kind -eq 'Subtitle' }).Count) {
            $tbs = @($items | Where-Object { $_.Kind -eq 'TextBox' })
            if ($tbs.Count -eq 1) { $tbs[0].Kind = 'Subtitle' }
        }
        if ($slideParts.IndexOf($sp) -eq 0) {
            $subItem = @($items | Where-Object { $_.Kind -eq 'Subtitle' } | Sort-Object Y, X, Order) | Select-Object -First 1
            if ($subItem) { $subtitleText = Get-ShapeText $subItem.Node $m }
        }

        # reading order = order of the shapes in the slide (first is read first)
        # visual order: title first; then column by column when the slide is laid out in columns
        # (each column top to bottom), otherwise row by row (each row left to right)
        $sorted = New-Object System.Collections.Generic.List[object]
        foreach ($it in @($items | Where-Object { $_.RK -eq 0 } | Sort-Object Y, X, Order)) { $sorted.Add($it) }
        $rest = @($items | Where-Object { $_.RK -ne 0 })
        # columns are found from the centre of each shape: a new column starts when the next
        # centre is more than 15 percent of the slide width away from the previous one
        $placed = @($rest | Where-Object { $_.Y -lt 1e9 } | Sort-Object CX, Y, Order)
        $unplaced = @($rest | Where-Object { $_.Y -ge 1e9 } | Sort-Object Order)
        $clusters = New-Object System.Collections.Generic.List[object]
        $cur = $null
        foreach ($it in $placed) {
            if ($null -eq $cur -or ($it.CX - $cur.Last) -gt ($slideWpt * 0.15)) {
                $cur = [pscustomobject]@{ Last = $it.CX; Items = (New-Object System.Collections.Generic.List[object]) }
                $clusters.Add($cur)
            }
            $cur.Items.Add($it)
            $cur.Last = $it.CX
        }
        $multi = @($clusters | Where-Object { $_.Items.Count -ge 2 }).Count
        if ($clusters.Count -ge 2 -and $multi -ge 2) {
            foreach ($cl in $clusters) { foreach ($it in @($cl.Items | Sort-Object Y, X, Order)) { $sorted.Add($it) } }
        } else {
            $rowList = New-Object System.Collections.Generic.List[object]
            $rowItems = $null; $rowStart = 0
            foreach ($it in @($placed | Sort-Object Y, X, Order)) {
                if ($null -eq $rowItems -or ($it.Y - $rowStart) -gt 18) {
                    $rowItems = New-Object System.Collections.Generic.List[object]
                    $rowStart = $it.Y
                    $rowList.Add($rowItems)
                }
                $rowItems.Add($it)
            }
            foreach ($r in $rowList) { foreach ($it in @($r | Sort-Object X, Order)) { $sorted.Add($it) } }
        }
        foreach ($it in $unplaced) { $sorted.Add($it) }

        # a shape that sits mostly behind a picture stays behind it (the order also decides what is drawn on top),
        # e.g. an "<Insert Fig ... here>" note hidden under the photo: it is placed just before that picture
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
                    if (($ow * $oh) / ($s1.W * $s1.H) -gt 0.5) {
                        $sorted.RemoveAt($i); $sorted.Insert($j, $s1); $moved = $true; break
                    }
                }
            }
        }
        $sorted = $sorted.ToArray()

        # names are numbered in reading order: Picture 1, TextBox 1, TextBox 2, Picture 2 ...
        $totals = @{}
        foreach ($it in $sorted) { if ($it.Kind) { if ($totals.ContainsKey($it.Kind)) { $totals[$it.Kind]++ } else { $totals[$it.Kind] = 1 } } }
        $running = @{}
        foreach ($it in $sorted) {
            if (-not $it.Kind -or -not $it.Nv) { continue }
            if ($running.ContainsKey($it.Kind)) { $running[$it.Kind]++ } else { $running[$it.Kind] = 1 }
            $nm = if ($totals[$it.Kind] -eq 1) { $it.Kind } else { $it.Kind + ' ' + $running[$it.Kind] }
            $it.Nv.SetAttribute('name', $nm)
            if ($it.Nv2) { $it.Nv2.SetAttribute('name', $nm) }
            $renamed++
        }
        $same = $true
        for ($i = 0; $i -lt $sorted.Count; $i++) { if ($sorted[$i].Order -ne $items[$i].Order) { $same = $false; break } }
        if (-not $same) {
            $anchor = $tree.SelectSingleNode('p:extLst', $m)
            foreach ($it in $sorted) {
                $null = $tree.RemoveChild($it.Top)
                if ($anchor) { $null = $tree.InsertBefore($it.Top, $anchor) } else { $null = $tree.AppendChild($it.Top) }
            }
            $reordered++
        }
    }
    $report.Add("Shapes renamed: $renamed")
    $report.Add("Slides with reading order changed: $reordered (order: title, then column by column or row by row as laid out)")

    # integrity: slide text may only grow by the counter
    $bad = 0
    $n = 0
    foreach ($sp in $slideParts) {
        $n++
        # same pieces of text as before, in any order, apart from the counter
        $now = (@($docs[$sp].SelectNodes('//a:t', $nss[$sp]) | ForEach-Object { [regex]::Replace($_.InnerText, '\s?\(\d{1,2} of \d{1,2}\)$', '') }) | Sort-Object) -join '|'
        $was = (@($before[$sp].Split('|') | ForEach-Object { [regex]::Replace($_, '\s?\(\d{1,2} of \d{1,2}\)$', '') }) | Sort-Object) -join '|'
        if ($now.Replace(' ', '') -ne $was.Replace(' ', '')) { $bad++; $report.Add("WARNING text changed on slide $n") }
    }
    $report.Add("Integrity check: $bad slides with unexpected text change")

    foreach ($sp in $slideParts) { Write-Part $sp $docs[$sp] }

    # 3. document title property = title of slide 1, a space, then its subtitle (plain space, no punctuation)
    $docTitle = ''
    $firstPart = $slideParts[0]
    if ($shapes.ContainsKey($firstPart)) {
        $docTitle = Get-ShapeText $shapes[$firstPart] $nss[$firstPart]
        if ($subtitleText) {
            # a "CHAPTER n" label goes first, then the chapter name
            if ($subtitleText -match '^(?i)\s*chapter\s+\d+') { $docTitle = $subtitleText + ' ' + $docTitle }
            else { $docTitle = $docTitle + ' ' + $subtitleText }
        }
    }
    if ($docTitle) {
        $core = Read-Part 'docProps/core.xml'
        if ($core) {
            $dcNs = 'http://purl.org/dc/elements/1.1/'
            $cm = New-Object System.Xml.XmlNamespaceManager($core.NameTable)
            $cm.AddNamespace('dc', $dcNs)
            $tn = $core.SelectSingleNode('//dc:title', $cm)
            if (-not $tn) {
                $tn = $core.CreateElement('dc', 'title', $dcNs)
                $null = $core.DocumentElement.PrependChild($tn)
            }
            $tn.InnerText = $docTitle
            Write-Part 'docProps/core.xml' $core
            $report.Add("Document title property set to: $docTitle")
        }
    } else {
        $report.Add('Document title property not changed (no title on slide 1)')
    }

    # 4. language tags: every language except none becomes English (US)
    $latin = [Text.Encoding]::GetEncoding(28591)
    $langRx = New-Object Text.RegularExpressions.Regex('(\s(?:lang|altLang)=")([^"]+)(")')
    $script:langChangedPart = 0
    foreach ($entry in @($zip.Entries)) {
        if (-not $entry.FullName.EndsWith('.xml')) { continue }
        $ms = New-Object IO.MemoryStream
        $es = $entry.Open(); $es.CopyTo($ms); $es.Dispose()
        $textL = $latin.GetString($ms.ToArray())
        $count = 0
        $newText = $langRx.Replace($textL, {
            param($mm)
            $v = $mm.Groups[2].Value
            if ($v -ieq 'en-US' -or $v -ieq 'x-none') { return $mm.Value }
            $script:langChangedPart++
            return $mm.Groups[1].Value + 'en-US' + $mm.Groups[3].Value
        })
        if ($newText -ne $textL) {
            $name = $entry.FullName
            $entry.Delete()
            $ne = $zip.CreateEntry($name)
            $os = $ne.Open()
            $bytes = $latin.GetBytes($newText)
            $os.Write($bytes, 0, $bytes.Length); $os.Dispose()
        }
    }
    $report.Add("Language tags changed to English (US): $script:langChangedPart")
}
finally {
    $zip.Dispose()
}

$repPath = Join-Path (Split-Path -Parent $OutPath) ([IO.Path]::GetFileNameWithoutExtension($OutPath) + ' report.txt')
$report | Set-Content -LiteralPath $repPath -Encoding UTF8
$report | ForEach-Object { Write-Host $_ }
Write-Host "Saved: $OutPath"
Add-WorkLog 'SLID Macro' $src $OutPath
}

Write-Host "SLID Macro - running version $SLID_VERSION" -ForegroundColor Cyan
Write-Host ''
if ($env:SLID_FILES) {
    $files = @($env:SLID_FILES.Split(';') | Where-Object { $_ })
} elseif ($env:FIXDECK_FILE) {
    $files = @($env:FIXDECK_FILE)
} else {
    Add-Type -AssemblyName System.Windows.Forms
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Select the PowerPoint file(s) to fix'
    $dlg.Filter = 'PowerPoint presentations (*.pptx)|*.pptx'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog($owner) -eq 'OK') { $files = $dlg.FileNames } else { $files = @() }
}
if ($files.Count -eq 0) { Write-Host 'No file selected.' }
foreach ($f in $files) {
    Write-Host ''
    Write-Host ("Fixing: " + $f)
    try { Fix-Deck -Path $f } catch { Write-Host ("FAILED: " + $_.Exception.Message) -ForegroundColor Red }
}
