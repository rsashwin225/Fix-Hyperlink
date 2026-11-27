# ALT TEXT TOOL SCRIPT
# PPT Alt-Text Filler (S4Carlisle). Self-updates from GitHub through AltText_Update.ps1.
$ALT_VERSION = '3.6'   # shown at the start of the window; the updater compares this number with GitHub
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
Write-Host ('PPT Alt-Text Filler - running version ' + $ALT_VERSION) -ForegroundColor Cyan
Write-Host ''
# AltText_Matcher.ps1
# Fills picture alt text in a .pptx from the "For Import" sheet of the alt text Excel file.
#
# Step 1  Make the review page: shows every picture with a dropdown of the Excel alt texts.
#         You pick the matching row for each picture, then press Save mapping.
# Step 2  Apply: reads the saved mapping and writes the Excel alt text into the pictures.
#
# Alt text written = Alt Text - Short, plus Alt Text - Long when the row has one.
# Pictures that already have alt text are kept as they are. Rows marked Decorative are
# marked decorative. Nothing else in the deck is changed, and the original file is never
# touched: the result is saved as a copy.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$NS_P   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NS_M   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NS_DEC = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
$DEC_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'

function Pick-File([string]$title, [string]$filter) {
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $title
    $dlg.Filter = $filter
    $dlg.InitialDirectory = Join-Path $env:USERPROFILE 'Downloads'
    if ($dlg.ShowDialog($owner) -eq 'OK') { return $dlg.FileName }
    return $null
}

function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
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

function Col-Letters([string]$ref) { return ([regex]::Match($ref, '^[A-Z]+')).Value }

function Norm-Chapter([string]$v) {
    $t = ($v + '').Trim()
    $n = 0
    if ([int]::TryParse($t, [ref]$n)) { return [string]$n }
    return $t.ToUpperInvariant()
}

# a figure / table / equation number written on the slide: "Figure 14.F03", "Table 14.T01", "Figure 14.EQ05"
$FIG_RX = '(?i)\b(Fig(?:ure)?|Table|Box|Eq(?:uation)?)\.?\s*(\d{1,3})\.([A-Za-z]{1,3}\s?\d{1,3})(?![A-Za-z0-9])'
function Norm-ElemId([string]$v) {
    $t = (($v + '') -replace '\s', '').ToUpperInvariant()
    $m = [regex]::Match($t, '^([A-Z]*)0*(\d+)$')
    if ($m.Success) { return ($m.Groups[1].Value + $m.Groups[2].Value) }
    return $t
}
function Get-FigRefs([string]$text) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($m in [regex]::Matches($text, $FIG_RX)) {
        $ch = Norm-Chapter $m.Groups[2].Value
        $id = Norm-ElemId $m.Groups[3].Value
        $dup = $false; foreach ($x in $out) { if ($x.Chapter -eq $ch -and $x.Id -eq $id) { $dup = $true } }
        if (-not $dup) { $out.Add([pscustomobject]@{ Chapter = $ch; Id = $id; Text = ($m.Groups[1].Value + ' ' + $m.Groups[2].Value + '.' + ($m.Groups[3].Value -replace '\s', '')) }) }
    }
    return @($out.ToArray())
}

# ---------- Excel ----------
function Get-AltRows([string]$xlsx) {
    $zip = [IO.Compression.ZipFile]::OpenRead($xlsx)
    try {
        $wb = New-Object Xml.XmlDocument; $wb.LoadXml((Read-ZipText $zip 'xl/workbook.xml'))
        $wns = New-Object Xml.XmlNamespaceManager($wb.NameTable)
        $wns.AddNamespace('m', $NS_M); $wns.AddNamespace('r', $NS_R)
        $rels = New-Object Xml.XmlDocument; $rels.LoadXml((Read-ZipText $zip 'xl/_rels/workbook.xml.rels'))
        $relMap = @{}
        foreach ($r in $rels.DocumentElement.ChildNodes) { if ($r.NodeType -eq 'Element') { $relMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target') } }

        $strings = New-Object System.Collections.Generic.List[string]
        $sstText = Read-ZipText $zip 'xl/sharedStrings.xml'
        if ($sstText) {
            $sd = New-Object Xml.XmlDocument; $sd.LoadXml($sstText)
            $sns = New-Object Xml.XmlNamespaceManager($sd.NameTable); $sns.AddNamespace('m', $NS_M)
            foreach ($si in $sd.DocumentElement.ChildNodes) {
                if ($si.NodeType -ne 'Element') { continue }
                $strings.Add((($si.SelectNodes('.//m:t', $sns) | ForEach-Object { $_.InnerText }) -join ''))
            }
        }

        $sheetList = @()
        foreach ($sh in $wb.SelectNodes('//m:sheets/m:sheet', $wns)) {
            $rid = $sh.GetAttribute('id', $NS_R)
            $sheetList += [pscustomobject]@{ Name = $sh.GetAttribute('name'); Part = (Resolve-Part 'xl/workbook.xml' $relMap[$rid]) }
        }

        $best = $null
        $ordered = @($sheetList | Where-Object { $_.Name -match 'import' }) + @($sheetList | Where-Object { $_.Name -notmatch 'import' })
        foreach ($sh in $ordered) {
            $grid = Read-Sheet $zip $sh.Part $strings
            foreach ($rn in ($grid.Keys | Sort-Object)) {
                $vals = @($grid[$rn].Values | ForEach-Object { ([string]$_).ToLowerInvariant() -replace '[^a-z0-9]', '' })
                if ($vals -contains 'elementnumber' -and ($vals -contains 'alttextshort' -or $vals -contains 'shortalttext')) { $best = @{ Grid = $grid; Header = $rn }; break }
            }
            if ($best) { break }
        }
        if (-not $best) { throw 'The Excel file has no sheet with the columns Element Number and Short Alt Text (or Alt Text - Short).' }

        $hdr = $best.Grid[$best.Header]
        $col = @{}
        # the column titles are matched by their words, so 'Alt Text - Short' and 'Short Alt Text' (and Long) both work
        $canon = @{ 'elementnumber' = 'Element Number'; 'alttextshort' = 'Alt Text - Short'; 'shortalttext' = 'Alt Text - Short'; 'alttextlong' = 'Alt Text - Long'; 'longalttext' = 'Alt Text - Long'; 'chapternumber' = 'Chapter Number'; 'usage' = 'Usage'; 'decorative' = 'Decorative' }
        foreach ($k in $hdr.Keys) {
            $nm = ([string]$hdr[$k]).ToLowerInvariant() -replace '[^a-z0-9]', ''
            if ($canon.ContainsKey($nm) -and -not $col.ContainsKey($canon[$nm])) { $col[$canon[$nm]] = $k }
            $col[$hdr[$k].Trim()] = $k
        }
        $get = { param($row, $name) if ($col.ContainsKey($name) -and $row.ContainsKey($col[$name])) { return [string]$row[$col[$name]] } else { return '' } }

        $out = New-Object System.Collections.Generic.List[object]
        foreach ($rn in ($best.Grid.Keys | Sort-Object)) {
            if ($rn -le $best.Header) { continue }
            $row = $best.Grid[$rn]
            $id = (& $get $row 'Element Number').Trim()
            if (-not $id) { continue }
            $out.Add([pscustomobject]@{
                Row = $rn
                Chapter = (Norm-Chapter (& $get $row 'Chapter Number'))
                Id = $id
                Usage = (& $get $row 'Usage').Trim()
                Decorative = ((& $get $row 'Decorative').Trim().ToUpperInvariant() -eq 'YES')
                Short = (& $get $row 'Alt Text - Short').Trim()
                Long = (& $get $row 'Alt Text - Long').Trim()
            })
        }
        return ,$out.ToArray()
    } finally { $zip.Dispose() }
}

function Read-Sheet($zip, [string]$part, $strings) {
    $grid = @{}
    $text = Read-ZipText $zip $part
    if (-not $text) { return $grid }
    $d = New-Object Xml.XmlDocument; $d.LoadXml($text)
    $m = New-Object Xml.XmlNamespaceManager($d.NameTable); $m.AddNamespace('m', $NS_M)
    foreach ($r in $d.SelectNodes('//m:sheetData/m:row', $m)) {
        $rn = [int]$r.GetAttribute('r')
        $cells = @{}
        foreach ($c in $r.SelectNodes('m:c', $m)) {
            $t = $c.GetAttribute('t')
            $v = ''
            if ($t -eq 'inlineStr') { $v = (($c.SelectNodes('.//m:t', $m) | ForEach-Object { $_.InnerText }) -join '') }
            else {
                $vn = $c.SelectSingleNode('m:v', $m)
                if ($vn) {
                    if ($t -eq 's') { $v = $strings[[int]$vn.InnerText] } else { $v = $vn.InnerText }
                }
            }
            if ($v -ne '') { $cells[(Col-Letters $c.GetAttribute('r'))] = $v }
        }
        if ($cells.Count) { $grid[$rn] = $cells }
    }
    return $grid
}

# ---------- PowerPoint ----------
function Get-SlideParts($zip) {
    $pres = New-Object Xml.XmlDocument; $pres.LoadXml((Read-ZipText $zip 'ppt/presentation.xml'))
    $pm = New-Object Xml.XmlNamespaceManager($pres.NameTable); $pm.AddNamespace('p', $NS_P); $pm.AddNamespace('r', $NS_R)
    $rels = New-Object Xml.XmlDocument; $rels.LoadXml((Read-ZipText $zip 'ppt/_rels/presentation.xml.rels'))
    $map = @{}
    foreach ($r in $rels.DocumentElement.ChildNodes) { if ($r.NodeType -eq 'Element') { $map[$r.GetAttribute('Id')] = (Resolve-Part 'ppt/presentation.xml' $r.GetAttribute('Target')) } }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($s in $pres.SelectNodes('//p:sldIdLst/p:sldId', $pm)) { $list.Add($map[$s.GetAttribute('id', $NS_R)]) }
    return ,$list.ToArray()
}

function New-SlideNs($doc) {
    $m = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    $m.AddNamespace('a', $NS_A); $m.AddNamespace('p', $NS_P); $m.AddNamespace('r', $NS_R)
    return ,$m
}

function Get-NodeText($node, $m) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($para in $node.SelectNodes('.//a:p', $m)) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($c in $para.ChildNodes) {
            if ($c.LocalName -eq 'br') { [void]$sb.Append(' ') }
            elseif ($c.LocalName -eq 'r' -or $c.LocalName -eq 'fld') {
                $t = $c.SelectSingleNode('a:t', $m)
                if ($t) { [void]$sb.Append($t.InnerText) }
            }
        }
        $parts.Add($sb.ToString())
    }
    return ([regex]::Replace(($parts -join ' '), '\s+', ' ')).Trim()
}

function Test-AiGenerated([string]$alt) {
    if (-not $alt) { return $false }
    return [bool]([regex]::IsMatch($alt, '(?i)(ai[\s-]*generat|generated\s+(by|using)\s+ai|automatically\s+generated|ai[\s-]*content\s+may\s+be\s+incorrect)'))
}

function Get-DeckPictures([string]$pptx) {
    $zip = [IO.Compression.ZipFile]::OpenRead($pptx)
    try {
        $parts = Get-SlideParts $zip
        $out = New-Object System.Collections.Generic.List[object]
        $creditRx = '(?i)(' + [char]169 + '|courtesy|shutterstock|istock|getty|alamy|copyright)'
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.LoadXml((Read-ZipText $zip $parts[$i]))
            $m = New-SlideNs $d
            $title = ''
            $ts = $d.SelectSingleNode('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="title" or @type="ctrTitle"]]', $m)
            if ($ts) { $title = Get-NodeText $ts $m }
            $slideText = (($d.SelectNodes('//a:t', $m) | ForEach-Object { $_.InnerText }) -join ' ')
            # short texts on the slide, cut off before any photo credit
            $geo = New-Object System.Collections.Generic.List[object]
            $shorts = New-Object System.Collections.Generic.List[string]
            foreach ($sp in $d.SelectNodes('//p:cSld/p:spTree/p:sp', $m)) {
                if ($sp.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph[@type="title" or @type="ctrTitle"]', $m)) { continue }
                $gt = Get-NodeText $sp $m
                $mm = [regex]::Match($gt, $creditRx)
                if ($mm.Success) { $gt = $gt.Substring(0, $mm.Index).Trim() }
                if ($gt.Length -eq 0 -or $gt.Length -gt 200) { continue }
                if ($gt.Length -le 80 -and -not $shorts.Contains($gt)) { $shorts.Add($gt) }
                if ($gt.Length -gt 80) { continue }
                $go = $sp.SelectSingleNode('p:spPr/a:xfrm/a:off', $m); $ge = $sp.SelectSingleNode('p:spPr/a:xfrm/a:ext', $m)
                if (-not $go -or -not $ge) { continue }
                $geo.Add([pscustomobject]@{ Text = $gt; X = [double]$go.GetAttribute('x'); Y = [double]$go.GetAttribute('y'); W = [double]$ge.GetAttribute('cx'); H = [double]$ge.GetAttribute('cy') })
            }
            foreach ($p in $d.SelectNodes('//p:pic', $m)) {
                $nv = $p.SelectSingleNode('p:nvPicPr/p:cNvPr', $m)
                $isVideo = [bool]($p.SelectSingleNode('.//a:videoFile|.//a:audioFile', $m))
                $label = ''
                $po = $p.SelectSingleNode('p:spPr/a:xfrm/a:off', $m); $pe = $p.SelectSingleNode('p:spPr/a:xfrm/a:ext', $m)
                if ($po -and $pe) {
                    $px = [double]$po.GetAttribute('x'); $py = [double]$po.GetAttribute('y'); $pw = [double]$pe.GetAttribute('cx'); $ph = [double]$pe.GetAttribute('cy')
                    $bestD = [double]::MaxValue
                    foreach ($g in $geo) {
                        $ox = [Math]::Min($px + $pw, $g.X + $g.W) - [Math]::Max($px, $g.X)
                        $oy = [Math]::Min($py + $ph, $g.Y + $g.H) - [Math]::Max($py, $g.Y)
                        $minW = [Math]::Min($pw, $g.W); $minH = [Math]::Min($ph, $g.H)
                        $dist = -1.0
                        if ($minW -gt 0 -and ($ox / $minW) -ge 0.3 -and $g.Y -ge ($py + $ph * 0.85)) {
                            $gap = [Math]::Max(0, $g.Y - ($py + $ph))
                            if ($gap -le 1000000) { $dist = $gap + 0.5 * [Math]::Abs(($g.X + $g.W / 2) - ($px + $pw / 2)) }
                        } elseif ($minH -gt 0 -and ($oy / $minH) -ge 0.3) {
                            if ($g.X -ge ($px + $pw)) { $gap = $g.X - ($px + $pw) } elseif (($g.X + $g.W) -le $px) { $gap = $px - ($g.X + $g.W) } else { $gap = 0 }
                            if ($gap -le 800000) { $dist = 400000 + $gap }
                        }
                        if ($dist -ge 0 -and $dist -lt $bestD) { $bestD = $dist; $label = $g.Text }
                    }
                }
                # the figure number of this picture: in the caption next to it, else the only figure number on the slide
                $figRef = $null
                if ($label) { $fr = @(Get-FigRefs $label); if ($fr.Count -ge 1) { $figRef = $fr[0] } }
                if (-not $figRef) { $fr = @(Get-FigRefs $slideText); if ($fr.Count -eq 1) { $figRef = $fr[0] } }
                $embed = ''
                $blipN = $p.SelectSingleNode('p:blipFill/a:blip', $m)
                if ($blipN) { $embed = $blipN.GetAttribute('embed', $NS_R) }
                $out.Add([pscustomobject]@{
                    Part = $parts[$i]; Embed = $embed;
                    Slide = $i + 1; Title = $title.Trim(); SlideText = $slideText; ShapeId = $nv.GetAttribute('id'); Name = $nv.GetAttribute('name')
                    Alt = $nv.GetAttribute('descr'); Label = $label; Shorts = $shorts.ToArray(); IsVideo = $isVideo
                    IsAi = (Test-AiGenerated $nv.GetAttribute('descr')); FigRef = $figRef
                })
            }
        }
        return ,$out.ToArray()
    } finally { $zip.Dispose() }
}

function Find-ShapeById($shapes, [int]$id) {
    foreach ($s in $shapes) {
        if ($s.Id -eq $id) { return $s }
        if ($s.Type -eq 6) { $r = Find-ShapeById $s.GroupItems $id; if ($r) { return $r } }
    }
    return $null
}

function Get-Thumbnails([string]$pptx, $pics) {
    $thumbs = @{}; $slideThumbs = @{}
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('altmatch_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    $app = $null; $pres = $null
    $wasRunning = [bool](Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    try {
        $app = New-Object -ComObject PowerPoint.Application
        $pres = $app.Presentations.Open($pptx, $true, $false, $false)
        foreach ($p in $pics) {
            try {
                $slide = $pres.Slides.Item($p.Slide)
                $sh = Find-ShapeById $slide.Shapes ([int]$p.ShapeId)
                if (-not $sh) { continue }
                $f = Join-Path $tmp ('p' + $p.Slide + '_' + $p.ShapeId + '.png')
                $sh.Export($f, 2)
                if (Test-Path -LiteralPath $f) { $thumbs["$($p.Slide)|$($p.ShapeId)"] = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f)) }
            } catch { }
        }
        # One image of the whole slide, as it actually looks in PowerPoint -
        # exported once per slide even when a slide has more than one picture.
        foreach ($slideNo in @($pics | Select-Object -ExpandProperty Slide -Unique)) {
            try {
                $slide = $pres.Slides.Item($slideNo)
                $fs = Join-Path $tmp ('s' + $slideNo + '.png')
                $slide.Export($fs, 'PNG', 1280, 720)
                if (Test-Path -LiteralPath $fs) { $slideThumbs["$slideNo"] = [Convert]::ToBase64String([IO.File]::ReadAllBytes($fs)) }
            } catch { }
        }
    } catch {
        Write-Host ('Pictures could not be previewed (PowerPoint was not available): ' + $_.Exception.Message)
    } finally {
        if ($pres) { try { $pres.Close() } catch { } }
        if ($app -and -not $wasRunning) { try { $app.Quit() } catch { } }
        try { Remove-Item -LiteralPath $tmp -Recurse -Force } catch { }
    }
    return @{ Pics = $thumbs; Slides = $slideThumbs }
}

# ---------- word comparison for the suggestions ----------
$stop = @('a','an','the','of','in','on','at','to','and','or','with','by','is','are','as','for','from','showing','photo','illustration','shows','that','its','their','table','titled','its','it','this','be','read','reads','reading')
function Get-Stem([string]$w) {
    $w = ($w.ToLowerInvariant() -replace '[^a-z0-9]', '')
    if ($w.Length -gt 6 -and $w.EndsWith('ers')) { $w = $w.Substring(0, $w.Length - 3) }
    elseif ($w.Length -gt 5 -and $w.EndsWith('er')) { $w = $w.Substring(0, $w.Length - 2) }
    elseif ($w.Length -gt 5 -and $w.EndsWith('ing')) { $w = $w.Substring(0, $w.Length - 3) }
    elseif ($w.Length -gt 5 -and $w.EndsWith('ed')) { $w = $w.Substring(0, $w.Length - 2) }
    elseif ($w.Length -gt 3 -and $w.EndsWith('s')) { $w = $w.Substring(0, $w.Length - 1) }
    if ($w.Length -ge 2 -and $w[$w.Length - 1] -eq $w[$w.Length - 2] -and $w[$w.Length - 1] -match '[a-z]' -and $w -notmatch '(ll|ss|ff|zz)$') { $w = $w.Substring(0, $w.Length - 1) }
    if ($w.Length -gt 4 -and $w.EndsWith('e')) { $w = $w.Substring(0, $w.Length - 1) }
    return $w
}
$generic = @('ambulance','vehicle','emergency','road','driver','person','people','picture','image','figure','photo','illustration') | ForEach-Object { Get-Stem $_ }
function Get-Tokens([string]$s) {
    $x = ($s.ToLowerInvariant() -replace '\bad\b', 'additional duty') -replace '[^a-z0-9 ]', ' '
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($w in $x.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) {
        if ($stop -contains $w) { continue }
        $st = Get-Stem $w
        if ($st.Length -gt 0 -and -not $out.Contains($st)) { $out.Add($st) }
    }
    return @($out.ToArray())
}
function Get-Score($a, $b) {
    if ($a.Count -eq 0 -or $b.Count -eq 0) { return 0.0 }
    $inter = @($a | Where-Object { $b -contains $_ }).Count
    $uni = @($a + $b | Select-Object -Unique).Count
    return [double]$inter / [double]$uni
}

function Html($s) { return [System.Net.WebUtility]::HtmlEncode([string]$s) }

# =====================================================================
# Step 1: review page
# =====================================================================
function Make-ReviewPage([string]$pptx, [string]$xlsx, [string]$chapter) {
    Write-Host 'Reading the Excel file...'
    $rows = Get-AltRows $xlsx
    $chap = Norm-Chapter $chapter
    $withText = @($rows | Where-Object { $_.Short -or $_.Long -or $_.Decorative })
    $opts = @($withText | Where-Object { $_.Chapter -eq $chap })
    if ($withText.Count -eq 0) { throw 'The Excel file has no rows with alt text.' }
    if ($opts.Count -eq 0) {
        $have = (@($withText | ForEach-Object { $_.Chapter } | Select-Object -Unique) -join ', ')
        Write-Host ("This Excel file has no rows for chapter " + $chapter + ". Chapters in the file: " + $have) -ForegroundColor Yellow
        Write-Host '  Check that it is the right Excel file for this book. The page still opens: every row of the file can be searched, and a figure number on a slide (for example Figure 14.F03) is matched in any chapter.' -ForegroundColor Yellow
    }
    # rows of the other chapters: only for Browse all rows and the figure-number match (their key carries the chapter)
    $others = @($withText | Where-Object { $_.Chapter -ne $chap })
    foreach ($o in $others) {
        $cl = if ($o.Chapter -match '^\d+$') { 'Ch ' + $o.Chapter } else { [string]$o.Chapter }
        $o | Add-Member -NotePropertyName Key -NotePropertyValue ($cl + ' ' + $o.Id) -Force
    }
    foreach ($o in $opts) { $o | Add-Member -NotePropertyName Key -NotePropertyValue $o.Id -Force }
    Write-Host ("Excel rows for chapter " + $chapter + ": " + $opts.Count + " (the whole file has " + $withText.Count + " rows, all can be searched)")

    Write-Host 'Reading the pictures in the deck...'
    $allPics = Get-DeckPictures $pptx
    $pics = @($allPics | Where-Object { -not $_.IsVideo })
    Write-Host ("Pictures in the deck: " + $pics.Count)
    Write-Host 'Making previews (PowerPoint opens the file quietly)...'
    $thumbResult = Get-Thumbnails $pptx $pics
    $thumbs = $thumbResult.Pics
    $slideThumbs = $thumbResult.Slides
    # Where PowerPoint could not make a preview, show the picture stored in the file (PNG, JPEG, GIF) instead.
    $missingPrev = @($pics | Where-Object { -not $thumbs.ContainsKey("$($_.Slide)|$($_.ShapeId)") -and $_.Embed })
    if ($missingPrev.Count -gt 0) {
        try {
            $zp = [IO.Compression.ZipFile]::OpenRead($pptx)
            try {
                foreach ($p in $missingPrev) {
                    try {
                        $relText = Read-ZipText $zp (([IO.Path]::GetDirectoryName($p.Part.Replace('/', '\')).Replace('\', '/')) + '/_rels/' + ($p.Part.Split('/')[-1]) + '.rels')
                        if (-not $relText) { continue }
                        $rd = New-Object Xml.XmlDocument; $rd.LoadXml($relText)
                        $tgt = $null
                        foreach ($rl in $rd.DocumentElement.ChildNodes) { if ($rl.NodeType -eq 'Element' -and $rl.GetAttribute('Id') -eq $p.Embed) { $tgt = $rl.GetAttribute('Target') } }
                        if (-not $tgt) { continue }
                        $mp = Resolve-Part $p.Part $tgt
                        if ($mp -notmatch '(?i)\.(png|jpe?g|gif)$') { continue }
                        $ent = $zp.GetEntry($mp)
                        if (-not $ent -or $ent.Length -gt 3000000) { continue }
                        $st = $ent.Open(); $ms = New-Object IO.MemoryStream; $st.CopyTo($ms); $st.Dispose()
                        $thumbs["$($p.Slide)|$($p.ShapeId)"] = [Convert]::ToBase64String($ms.ToArray())
                    } catch { }
                }
            } finally { $zp.Dispose() }
        } catch { }
    }

    $optShort = @{}; $optLong = @{}
    foreach ($o in $opts) {
        $tx = $o.Short; if (-not $tx) { $tx = $o.Long }
        $optShort[$o.Id] = @(Get-Tokens $tx)
        $optLong[$o.Id] = @(Get-Tokens $o.Long)
    }
    $js = { param($s) (ConvertTo-Json -InputObject ([string]$s) -Compress).Replace('<', ([string][char]92 + 'u003c')) }

    $optJson = New-Object System.Collections.Generic.List[string]
    foreach ($o in $opts) {
        $full = $o.Short
        if ($o.Long) { $full = ($full + ' ' + $o.Long).Trim() }
        $dec = 'false'; if ($o.Decorative) { $dec = 'true' }
        $optJson.Add('{"id":' + (& $js $o.Id) + ',"ch":' + (& $js $o.Chapter) + ',"ex":' + (& $js $o.Id) + ',"full":' + (& $js $full) + ',"dec":' + $dec + '}')
    }
    $othJson = New-Object System.Collections.Generic.List[string]
    foreach ($o in $others) {
        $full = $o.Short
        if ($o.Long) { $full = ($full + ' ' + $o.Long).Trim() }
        $dec = 'false'; if ($o.Decorative) { $dec = 'true' }
        $othJson.Add('{"id":' + (& $js $o.Key) + ',"ch":' + (& $js $o.Chapter) + ',"ex":' + (& $js $o.Id) + ',"full":' + (& $js $full) + ',"dec":' + $dec + '}')
    }

    $picJson = New-Object System.Collections.Generic.List[string]
    $cards = New-Object System.Text.StringBuilder
    $n = 0
    foreach ($p in $pics) {
        # phrases on the slide that the Excel rows are compared with
        $phr = New-Object System.Collections.Generic.List[object]
        $titleClean = ([regex]::Replace([string]$p.Title, '\s*\(\d+ of \d+\)\s*$', '')).Trim()
        $titleTok = @(Get-Tokens $titleClean | Where-Object { $generic -notcontains $_ })
        if ($titleTok.Count) { $phr.Add([pscustomobject]@{ Text = $titleClean; Tok = $titleTok; Kind = 'title' }) }
        if ($p.Label) {
            $lt = @(Get-Tokens $p.Label | Where-Object { $generic -notcontains $_ })
            if ($lt.Count) { $phr.Add([pscustomobject]@{ Text = $p.Label; Tok = $lt; Kind = 'label' }) }
        }
        foreach ($sx in $p.Shorts) {
            if ($p.Label) { break }
            $st2 = @(Get-Tokens $sx | Where-Object { $generic -notcontains $_ })
            if ($st2.Count) { $phr.Add([pscustomobject]@{ Text = $sx; Tok = $st2; Kind = 'text' }) }
        }
        # the rest of the slide's own text (bullets, body copy, anything not
        # already covered above) - a weaker, lower-priority signal, since it
        # is a much bigger and less targeted block of words than the title,
        # label or a short caption right next to the picture.
        if ($p.SlideText) {
            $slideTok = @(Get-Tokens $p.SlideText | Where-Object { $generic -notcontains $_ })
            if ($slideTok.Count) { $phr.Add([pscustomobject]@{ Text = $p.SlideText; Tok = $slideTok; Kind = 'slide' }) }
        }
        $altTok = @(); if ($p.Alt) { $altTok = @(Get-Tokens $p.Alt) }

        $found = New-Object System.Collections.Generic.List[object]
        $order = 0
        foreach ($o in $opts) {
            $order++
            $sT = $optShort[$o.Id]; $lT = $optLong[$o.Id]
            $best = $null
            foreach ($ph in $phr) {
                $nn = $ph.Tok.Count
                $ms = @($ph.Tok | Where-Object { $sT -contains $_ })
                $ml = @($ph.Tok | Where-Object { ($lT -contains $_) -and ($sT -notcontains $_) })
                $maxLen = 0; foreach ($w in $ms) { if ($w.Length -gt $maxLen) { $maxLen = $w.Length } }
                $need = [Math]::Min(2, $nn)
                $okShort = ($ms.Count -ge $need) -or ($ms.Count -ge 1 -and $maxLen -ge 8)
                $cShort = 0.0; if ($okShort) { $cShort = $ms.Count / $nn }
                $cLong = 0.0
                if ($ml.Count -gt 0 -and (($ms.Count + $ml.Count) -ge $need)) { $cLong = 0.7 * ($ms.Count + $ml.Count) / $nn }
                $cWeak = 0.0; if (-not $okShort -and $ms.Count -ge 1 -and $maxLen -ge 4) { $cWeak = 0.4 }
                $score = [Math]::Max([Math]::Max($cShort, $cLong), $cWeak)
                $strong = $false
                if ($cShort -gt 0) {
                    if ($nn -ge 2 -and $cShort -ge 0.999) { $score = 1.0; $strong = $true }
                    elseif ($maxLen -ge 8 -and $cShort -ge 0.5 -and $nn -le 2) { $strong = $true }
                }
                if ($ph.Kind -eq 'label' -and $ms.Count -ge 2) {
                    $hyb = 0.5 * ($ms.Count / $sT.Count) + 0.5 * ($ms.Count / ($sT.Count + $nn - $ms.Count))
                    if ($hyb -gt $score) { $score = $hyb }
                    if ($hyb -ge 0.45) { $strong = $true }
                }
                if ($ph.Kind -eq 'slide' -and $ms.Count -ge 3) {
                    # The whole slide's text is long and noisy next to a short
                    # Excel description, so dividing by its own huge word
                    # count (as cShort does) crushes even a genuinely strong
                    # overlap. Score by how much of the CANDIDATE's own words
                    # are corroborated somewhere on the slide instead.
                    $recall = $ms.Count / $sT.Count
                    $slideScore = 0.3 + 0.5 * $recall
                    if ($slideScore -gt $score) { $score = $slideScore }
                    if ($recall -ge 0.5 -and $ms.Count -ge 4) { $strong = $true }
                }
                if ($ph.Kind -eq 'title' -and $p.Label -and -not $strong) { continue }
                $better = ($null -eq $best) -or ($strong -and -not $best.Strong) -or (($strong -eq $best.Strong) -and ($score -gt $best.Score))
                if ($score -ge 0.4 -and $better) {
                    $best = [pscustomobject]@{ Score = $score; Strong = $strong; Words = @($ms + $ml); From = $ph.Text; InLong = ($cShort -lt $cLong -and $ml.Count -gt 0); Kind = $ph.Kind }
                }
            }
            if ($altTok.Count) {
                $sa = Get-Score $altTok $sT
                $saStrong = ($sa -ge 0.6)
                $betterA = ($null -eq $best) -or ($saStrong -and -not $best.Strong) -or (($saStrong -eq $best.Strong) -and ($sa -gt $best.Score))
                if ($sa -ge 0.34 -and $betterA) {
                    $best = [pscustomobject]@{ Score = $sa; Strong = ($sa -ge 0.6); Words = @($altTok | Where-Object { $sT -contains $_ }); From = 'its current alt text'; InLong = $false; Kind = 'alt' }
                }
            }
            if ($best -and ($best.Strong -or $o.Id -notmatch '^(CO|C)\d+$')) { $found.Add([pscustomobject]@{ Id = $o.Id; Order = $order; Best = $best }) }
        }
        # A weak match that only comes from words shared with the slide's
        # own title is treated as less trustworthy than a weak match that
        # comes from the picture's label, a short caption, or the slide's
        # own body text - title wording (especially a running chapter
        # title) often repeats across many unrelated rows, so by itself it
        # is a poor way to tell them apart.
        $top = @($found | Sort-Object @{ Expression = { if ($_.Best.Strong) { 0 } else { 1 } } }, @{ Expression = { if ((-not $_.Best.Strong) -and $_.Best.Kind -eq 'title') { 1 } else { 0 } } }, @{ Expression = { -$_.Best.Score } }, Order | Select-Object -First 5)
        if (-not $top.Count) {
            # nothing is close: offer the rows that share the most words, marked weak
            $weakList = New-Object System.Collections.Generic.List[object]
            $order2 = 0
            foreach ($o in $opts) {
                $order2++
                $allTok = New-Object System.Collections.Generic.List[string]
                foreach ($ph in $phr) { foreach ($w in $ph.Tok) { if (-not $allTok.Contains($w)) { $allTok.Add($w) } } }
                if ($o.Id -match '^(CO|C)\d+$') { continue }
                $mw = @($allTok | Where-Object { $_.Length -ge 5 -and (($optShort[$o.Id] -contains $_) -or ($optLong[$o.Id] -contains $_)) })
                if ($mw.Count -ge 1) { $weakList.Add([pscustomobject]@{ Id = $o.Id; Order = $order2; Best = [pscustomobject]@{ Score = $mw.Count / [Math]::Max(1, $allTok.Count); Strong = $false; Words = $mw; From = 'words on the slide'; InLong = $false } }) }
            }
            $top = @($weakList | Sort-Object @{ Expression = { -$_.Best.Score } }, Order | Select-Object -First 3)
        }
        $candJson = @()
        # the figure number on the slide ("Figure 14.F03") finds its Excel row first - in any chapter of the file
        $figJs = '""'; $figMiss = 'false'
        if ($p.FigRef) {
            $figJs = (& $js $p.FigRef.Text)
            $hit = $null
            foreach ($o in @($opts) + @($others)) { if ($o.Chapter -eq $p.FigRef.Chapter -and (Norm-ElemId $o.Id) -eq $p.FigRef.Id) { $hit = $o; break } }
            if ($hit) {
                $candJson += ('{"id":' + (& $js $hit.Key) + ',"st":"s","w":[],"from":' + (& $js ('the figure number ' + $p.FigRef.Text)) + ',"sc":2,"fig":true,"long":false}')
                $top = @($top | Where-Object { $_.Id -ne $hit.Key })
            } else { $figMiss = 'true' }
        }
        foreach ($c in $top) {
            $wj = ($c.Best.Words | ForEach-Object { & $js $_ }) -join ','
            $stj = 'w'; if ($c.Best.Strong) { $stj = 's' }
            $lg = 'false'; if ($c.Best.InLong) { $lg = 'true' }
            $candJson += ('{"id":' + (& $js $c.Id) + ',"st":"' + $stj + '","w":[' + $wj + '],"from":' + (& $js $c.Best.From) + ',"sc":' + ([Math]::Round([double]$c.Best.Score, 2).ToString([Globalization.CultureInfo]::InvariantCulture)) + ',"long":' + $lg + '}')
        }
        $swStems = @(Get-Tokens $titleClean | Where-Object { $generic -notcontains $_ })
        $swShow = (($titleClean.ToLowerInvariant() -replace '[^a-z0-9 ]', ' ').Split(' ', [StringSplitOptions]::RemoveEmptyEntries) | Where-Object { ($stop -notcontains $_) -and ($generic -notcontains $_) }) -join ' '
        $aiJs = 'false'; if ($p.IsAi) { $aiJs = 'true' }
        $picJson.Add('{"i":' + $n + ',"slide":' + $p.Slide + ',"sid":' + (& $js $p.ShapeId) + ',"name":' + (& $js $p.Name) + ',"title":' + (& $js $p.Title) + ',"alt":' + (& $js $p.Alt) + ',"ai":' + $aiJs + ',"label":' + (& $js $p.Label) + ',"fig":' + $figJs + ',"figmiss":' + $figMiss + ',"sw":' + (& $js ($swStems -join ' ')) + ',"swt":' + (& $js $swShow) + ',"cand":[' + ($candJson -join ',') + ']}')

        $img = '<div class="noimg">No preview</div>'
        $key = "$($p.Slide)|$($p.ShapeId)"
        $hasPicThumb = $thumbs.ContainsKey($key)
        $hasSlideThumb = $slideThumbs.ContainsKey("$($p.Slide)")
        if ($hasPicThumb) { $img = '<img src="data:image/png;base64,' + $thumbs[$key] + '" alt="">' }
        $zh = ''; if ($hasPicThumb -or $hasSlideThumb) { $zh = '<span class="zh">Click to zoom</span>' }
        [void]$cards.Append('<section class="card" data-i="' + $n + '"><div class="th">' + $img + $zh + '</div><div class="main"></div></section>')
        $n++
    }

    $deckName = [IO.Path]::GetFileNameWithoutExtension($pptx)
    $html = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Alt text review</title>
<style>
:root{--bg:#f4f5f7;--card:#fff;--text:#1b1f23;--muted:#59636e;--line:#d0d7de;--accent:#0b5cad;--accent-bg:#ddeeff;--ok:#1a7f37;--ok-bg:#e6f6ea;--warn:#8a5a00;--warn-bg:#fff3d1;--err:#a40e26;--soft:#eef1f4}
@media (prefers-color-scheme: dark){:root{--bg:#161b22;--card:#1f2630;--text:#e6edf3;--muted:#9aa4af;--line:#38414d;--accent:#58a6ff;--accent-bg:#14304f;--ok:#56d364;--ok-bg:#12301b;--warn:#e3b341;--warn-bg:#3a2d0b;--err:#ff7b72;--soft:#262e3a}}
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--text);padding-bottom:96px}
.wrap{max-width:1000px;margin:0 auto;padding:0 16px}
header{position:sticky;top:0;z-index:5;background:var(--bg);padding:14px 0 10px;border-bottom:1px solid var(--line)}
.top{display:flex;gap:16px;align-items:center;flex-wrap:wrap}
.top h1{margin:0;font-size:18px;font-weight:600}.sub{font-size:13px;color:var(--muted)}
.prog{width:270px}.prog .t{display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px;gap:8px}
.track{height:8px;border-radius:4px;background:var(--line)}.fillbar{height:8px;border-radius:4px;background:var(--ok);width:0;transition:width .25s}
.tools{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;align-items:center}
.chip{border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:999px;padding:5px 12px;font-size:13px;cursor:pointer}
.chip.on{background:var(--accent-bg);color:var(--accent);border-color:var(--accent)}
.grow{flex:1}.search{padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-size:14px;width:230px}
.card{display:flex;gap:16px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;margin-top:12px}
.card.ok{border-color:var(--ok)}.card.own{border-color:var(--accent)}.card.uns{border-color:var(--warn)}
.th{width:260px;flex:none;text-align:center;cursor:zoom-in;align-self:flex-start}.th img{max-width:260px;max-height:210px;border:1px solid var(--line);border-radius:6px;background:#fff}
.zh{display:block;font-size:12px;color:var(--accent);margin-top:6px}
.noimg{height:120px;background:var(--soft);color:var(--muted);display:flex;align-items:center;justify-content:center;border-radius:6px;font-size:13px}
.main{flex:1;min-width:0}
.head{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.head b{font-size:15px;font-weight:600}
.pill{font-size:12px;padding:2px 10px;border-radius:999px}
.pill.need,.pill.uns{background:var(--warn-bg);color:var(--warn)}.pill.has{background:var(--accent-bg);color:var(--accent)}.pill.ok{background:var(--ok-bg);color:var(--ok)}.pill.own{background:var(--accent-bg);color:var(--accent)}.pill.deco{background:var(--soft);color:var(--muted)}
.ctx{font-size:13px;color:var(--muted);margin-top:2px}
.lab{font-size:13px;margin-top:6px}.lab b{font-weight:600}
.has{font-size:13px;color:var(--muted);margin-top:6px}
.lbl{font-size:12px;color:var(--muted);margin-top:10px}
.sgw{margin-top:6px}
.sg,.opt{display:block;width:100%;text-align:left;font:inherit;font-size:13px;padding:8px 10px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--text);cursor:pointer;line-height:1.45}
.sg:hover,.sg:focus,.opt:hover,.opt:focus{background:var(--accent-bg);border-color:var(--accent);outline:none}
.sg b,.opt b{color:var(--accent);margin-right:6px}.sg i,.opt i{font-style:normal;color:var(--muted);font-size:12px;margin-left:6px}
.more{font-size:12px;color:var(--accent);margin-left:4px}
mark{background:var(--warn-bg);color:var(--warn);border-radius:3px;padding:0 2px}
.why{display:flex;gap:6px;flex-wrap:wrap;margin-top:4px}
.tag{font-size:11px;padding:1px 8px;border-radius:999px;background:var(--soft);color:var(--muted)}
.why .use{margin:0 0 0 auto;padding:2px 10px;font-size:12px}
.tag.s{background:var(--ok-bg);color:var(--ok)}.tag.w{background:var(--warn-bg);color:var(--warn)}.tag.o{background:var(--accent-bg);color:var(--accent)}
.sg.used,.opt.used{border-color:var(--warn);background:var(--warn-bg)}
.usedbadge{float:right;display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;color:var(--warn);background:var(--card);border:1px solid var(--warn);border-radius:999px;padding:2px 8px;margin-left:8px}
.usedbadge svg{width:11px;height:11px;fill:none;stroke:currentColor;stroke-width:3}
.lbsearchrow{padding:10px 14px;border-bottom:1px solid var(--line);background:var(--soft)}
.lbsearchrow input{width:100%;padding:8px 10px;font-size:13px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-family:inherit}
.toggle{display:flex;gap:2px;background:var(--soft);border-radius:999px;padding:2px}
.toggle button{border:none;background:transparent;font:inherit;font-size:13px;padding:6px 12px;border-radius:999px;cursor:pointer;color:var(--muted)}
.toggle button.on{background:var(--card);color:var(--accent);font-weight:600}
.small{font:inherit;font-size:13px;margin-top:10px;padding:6px 12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--text);cursor:pointer}
.rowsbox{margin-top:8px;border:1px solid var(--line);border-radius:8px;padding:8px;background:var(--soft)}
.flt,.notein{width:100%;padding:8px 10px;font-size:13px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-family:inherit}
.rows{margin-top:8px;max-height:340px;overflow:auto;display:flex;flex-direction:column;gap:6px}
.owntext{width:100%;min-height:62px;margin-top:6px;padding:8px 10px;font-size:13px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-family:inherit;resize:vertical}
.acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
.acts .small{margin-top:0}
.err{font-size:13px;color:var(--err);margin-top:4px}
.chosen,.state{margin-top:10px;font-size:13px;border-radius:8px;padding:10px 12px;line-height:1.5}
.chosen{background:var(--ok-bg);color:var(--ok)}.chosen a{color:var(--ok);margin-left:8px}
.state.own{background:var(--accent-bg);color:var(--accent)}.state.own a{color:var(--accent);margin-left:8px}
.state.uns{background:var(--warn-bg);color:var(--warn)}
.state.deco{background:var(--soft);color:var(--muted)}.state.deco a{color:var(--accent);margin-left:8px}
.note{font-size:12px;color:var(--muted);margin-top:6px}
.empty{padding:40px;text-align:center;color:var(--muted)}
#bar{position:fixed;left:0;right:0;bottom:0;background:var(--card);border-top:1px solid var(--line);padding:12px 16px;z-index:9}
#bar .in{max-width:1000px;margin:0 auto;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
#bar label{display:flex;gap:8px;align-items:center;font-size:13px;color:var(--muted);flex:1;min-width:220px}
#msg{font-size:13px;color:var(--muted)}
button.btn{font:inherit;font-size:14px;padding:9px 16px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--text);cursor:pointer}
button.main-btn{background:var(--text);color:var(--card);border-color:var(--text)}
button:disabled{opacity:.5;cursor:default}
#cf{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:30;display:flex;align-items:center;justify-content:center;padding:16px}
#cf[hidden]{display:none}
.cfbox{background:var(--card);color:var(--text);border-radius:12px;max-width:620px;width:100%;padding:18px 20px;box-shadow:0 10px 30px rgba(0,0,0,.35)}
.cfbox h2{font-size:17px;margin:0 0 4px}.cfsub{font-size:13px;color:var(--muted);margin:0 0 10px}
.cflbl{font-size:12px;color:var(--muted);margin-top:10px}
.cfold,.cfnew{font-size:13px;line-height:1.5;border-radius:8px;padding:8px 10px;margin-top:4px;max-height:160px;overflow:auto}
.cfold{background:var(--soft);text-decoration:line-through;text-decoration-color:var(--muted)}.cfnew{background:var(--ok-bg);color:var(--ok)}
.cfrow{display:flex;gap:10px;margin-top:16px;flex-wrap:wrap}
#panel{position:fixed;left:16px;right:16px;bottom:84px;max-width:968px;margin:0 auto;max-height:60vh;overflow:auto;background:var(--card);border:1px solid var(--ok);border-radius:12px;padding:16px 18px;z-index:10;box-shadow:0 6px 24px rgba(0,0,0,.25)}
#panel.err{border-color:var(--err)}#panel h2{font-size:16px;margin:0 0 8px}#panel ul{margin:4px 0 10px 18px;padding:0;font-size:13px}#panel .row{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap}
#lb{position:fixed;inset:0;background:rgba(0,0,0,.82);z-index:20;display:flex;align-items:center;justify-content:center;padding:16px}
#lb[hidden]{display:none}
.lbbox{background:var(--card);border-radius:12px;width:min(1100px,100%);max-height:100%;display:flex;flex-direction:column;overflow:hidden}
.lbtop{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 14px;border-bottom:1px solid var(--line)}
.lbtop b{font-size:15px;font-weight:600}.lbtop .s{font-size:13px;color:var(--muted)}
.lbtop button{font:inherit;font-size:13px;padding:6px 12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--text);cursor:pointer}
#lbzoom{font-size:13px;min-width:44px;text-align:center}
.lbimg{overflow:auto;flex:1;min-height:200px;background:var(--soft);padding:12px;text-align:center}
.lbimg img{background:#fff;max-width:none}
.lbpick{padding:10px 14px;border-top:1px solid var(--line);max-height:34vh;overflow:auto}
@media (max-width:700px){.card{flex-direction:column}.th{width:auto}}
</style></head><body>
<header><div class="wrap">
<div class="top"><div class="grow"><h1>Alt text review</h1><div class="sub">__DECK__ &middot; chapter __CHAPTER__</div></div>
<div class="prog"><div class="t"><b id="cnt"></b><span id="pct"></span></div><div class="track"><div class="fillbar" id="fill"></div></div></div></div>
<div class="tools">
<button class="chip on" data-f="all">All <span id="n-all"></span></button>
<button class="chip" data-f="need">Needs a match <span id="n-need"></span></button>
<button class="chip" data-f="done">Matched <span id="n-done"></span></button>
<button class="chip" data-f="own">Your own text <span id="n-own"></span></button>
<button class="chip" data-f="uns">Not sure <span id="n-uns"></span></button>
<button class="chip" data-f="has">Already has alt text <span id="n-has"></span></button>
<span class="grow"></span><input class="search" id="q" type="search" placeholder="Search slide, name or text" aria-label="Search pictures">
</div></div></header>
<div class="wrap" id="list">__CARDS__<div class="empty" id="none" hidden>Nothing here. Try another filter.</div></div>
<div id="lb" hidden><div class="lbbox" role="dialog" aria-label="Zoomed picture">
<div class="lbtop"><b id="lbtitle"></b><span class="s" id="lbsub"></span><span class="grow"></span>
<div class="toggle" role="group" aria-label="View"><button id="lbviewpic" class="on" aria-pressed="true">Picture</button><button id="lbviewslide" aria-pressed="false">Whole slide</button></div>
<button id="lbminus" aria-label="Zoom out">Zoom out</button><span id="lbzoom">100%</span><button id="lbplus" aria-label="Zoom in">Zoom in</button>
<button id="lbprev">Previous</button><button id="lbnext">Next</button><button id="lbclose">Close</button></div>
<div class="lbimg" id="lbimgbox"><img id="lbimg" alt=""></div>
<div class="lbsearchrow" id="lbsearchrow" hidden><input id="lbsearch" type="search" placeholder="Search the whole Excel file, for example crash or Figure 14.F03" aria-label="Search the Excel rows"></div>
<div class="lbpick" id="lbpick"></div></div></div>
<div id="panel" hidden></div>
<div id="cf" hidden><div class="cfbox" role="dialog" aria-modal="true" aria-labelledby="cft">
<h2 id="cft">Replace the alt text?</h2><p class="cfsub" id="cfsub"></p>
<div class="cflbl">Alt text it has now</div><div class="cfold" id="cfold"></div>
<div class="cflbl">New alt text</div><div class="cfnew" id="cfnew"></div>
<div class="cfrow"><button class="btn main-btn" id="cfyes">Replace it</button><button class="btn" id="cfno">Keep the current alt text</button></div></div></div>
<div id="bar"><div class="in"><span class="grow" id="msg"></span>
<button class="btn" id="save">Save mapping</button><button class="btn main-btn" id="apply">Add alt text to the PowerPoint</button></div></div>
<script>
var OPTS = __OPTS__;
var OTHERS = __OTHERS__;
var PICS = __PICS__;
var SLIDETHUMBS = __SLIDETHUMBS__;
var CHAPTER = "__CHAPTER__", DECK = "__DECK__", TOKEN = "__TOKEN__";
var byId = {}; OPTS.forEach(function(o, k){ o.k = k; o.ord = k; o.text = (o.dec ? "(marked decorative in the Excel file) " : "") + o.full; byId[o.id] = o; });
OTHERS.forEach(function(o, k){ o.k = -1; o.ord = OPTS.length + k; o.text = (o.dec ? "(marked decorative in the Excel file) " : "") + o.full; byId[o.id] = o; });
var ALLOPTS = OPTS.concat(OTHERS);
var S = {}; PICS.forEach(function(p){ S[p.i] = {mode: "none", id: "", text: "", note: "", err: ""}; });
var filt = "all", query = "", moreSug = {}, openBrowse = {}, fltText = {}, expanded = {}, lbIdx = -1, lbScale = 1, lbBase = 0, lbQuery = "", lbView = "pic";
var cards = [].slice.call(document.querySelectorAll(".card"));
function esc(s){ return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function cut(s, n){ s = String(s); return s.length > n ? s.slice(0, n).replace(/\s+\S*$/, "") + "..." : s; }
function reset(i){ S[i] = {mode: "none", id: "", text: S[i].text || "", note: "", err: ""}; }
function usedMap(){ var m = {}; PICS.forEach(function(p){ var s = S[p.i]; if(s.mode === "excel"){ (m[s.id] = m[s.id] || []).push(p.slide); } }); return m; }
function lastKFor(idx){ var k = -1; for(var j = 0; j < idx; j++){ var s = S[PICS[j].i]; if(s.mode === "excel" && byId[s.id].k >= 0) k = byId[s.id].k; } return k; }
function nextKFor(idx){ for(var j = idx + 1; j < PICS.length; j++){ var s = S[PICS[j].i]; if(s.mode === "excel" && byId[s.id].k >= 0) return byId[s.id].k; } return -1; }
function stem(w){
  w = String(w).toLowerCase().replace(/[^a-z0-9]/g, "");
  if(w.length > 6 && /ers$/.test(w)) w = w.slice(0, -3);
  else if(w.length > 5 && /er$/.test(w)) w = w.slice(0, -2);
  else if(w.length > 5 && /ing$/.test(w)) w = w.slice(0, -3);
  else if(w.length > 5 && /ed$/.test(w)) w = w.slice(0, -2);
  else if(w.length > 3 && /s$/.test(w)) w = w.slice(0, -1);
  if(/([a-z])\1$/.test(w) && !/(ll|ss|ff|zz)$/.test(w)) w = w.slice(0, -1);
  if(w.length > 4 && /e$/.test(w)) w = w.slice(0, -1);
  return w;
}
function hl(text, stems){
  var set = {}; (stems || []).forEach(function(x){ set[x] = 1; });
  return String(text).split(" ").map(function(w){ var out = esc(w); return set[stem(w)] ? "<mark>" + out + "</mark>" : out; }).join(" ");
}
function suggest(p, idx, used){
  var sg = [], seen = {}, lastK = lastKFor(idx), nextK = nextKFor(idx), win = {}, winList = [];
  if(nextK >= 0){ for(var k = lastK + 1; k < nextK; k++){ var oid = OPTS[k].id; if(!used[oid]){ win[oid] = 1; winList.push(oid); } } }
  var useWin = winList.length > 0 && winList.length <= 6;
  (p.cand || []).forEach(function(c){ if(byId[c.id] && !seen[c.id]){ seen[c.id] = 1; sg.push({id: c.id, kind: c.st === "s" ? "strong" : "weak", sc: c.sc || 0, words: c.w || [], from: c.from || "", inLong: !!c.long, fig: !!c.fig, between: useWin && !!win[c.id]}); } });
  if(useWin){ var addB = 0; winList.forEach(function(id){ if(!seen[id] && addB < 3 && !/^(CO|C)\d+$/i.test(id)){ seen[id] = 1; sg.push({id: id, kind: "between", words: [], from: "", inLong: false, between: true}); addB++; } }); }
  if(nextK < 0){ var add = 0; for(var k2 = lastK + 1; k2 < OPTS.length && add < 2; k2++){ var o2 = OPTS[k2].id; if(!used[o2] && !seen[o2] && !/^(CO|C)\d+$/i.test(o2)){ seen[o2] = 1; sg.push({id: o2, kind: "next", words: [], from: "", inLong: false, between: false}); add++; } } }
  var base = {strong: 0, weak: 100, between: 200, next: 300};
  sg.forEach(function(s, n){ s.rank = base[s.kind] + (used[s.id] ? 500 : 0) + (s.between && s.kind !== "between" ? -50 : 0) + n - (s.fig ? 2000 : 0); });
  sg.sort(function(a, b){ return a.rank - b.rank; });
  return sg.slice(0, 7);
}
function usedBadge(u){ return u ? '<span class="usedbadge"><svg viewBox="0 0 24 24"><path d="M5 13l4 4L19 7"/></svg>Used on slide ' + esc(u.join(", ")) + '</span>' : ""; }
function sugItem(p, s, used, winner){
  var txt = byId[s.id].text, key = p.i + "|" + s.id, long = txt.length > 170, shown = (long && !expanded[key]) ? cut(txt, 170) : txt, u = used[s.id], tags = "";
  if(s.fig) tags += '<span class="tag s">figure number match</span>';
  else if(s.kind === "strong") tags += '<span class="tag s">strong match</span>';
  else if(s.kind === "weak") tags += '<span class="tag w">weak match</span>';
  else if(s.kind === "between") tags += '<span class="tag o">fits between your picks</span>';
  else tags += '<span class="tag n">next in order</span>';
  if(s.between && (s.kind === "strong" || s.kind === "weak")) tags += '<span class="tag o">fits between your picks</span>';
  if(s.from && (s.kind === "strong" || s.kind === "weak")) tags += '<span class="tag n">matches: ' + esc(cut(s.from, 40)) + '</span>';
  if(s.inLong) tags += '<span class="tag n">found in the long description</span>';
  var h = '<div class="sgw"><button class="sg' + (u ? ' used' : '') + '" data-i="' + p.i + '" data-id="' + esc(s.id) + '">' + usedBadge(u) + '<b>' + esc(s.id) + '</b>' + hl(shown, s.words) + '</button><div class="why">' + tags;
  if(winner) h += '<button class="small use" data-i="' + p.i + '" data-id="' + esc(s.id) + '">Use this</button>';
  h += '</div>';
  if(long) h += '<a href="#" class="more" data-more="' + esc(key) + '">' + (expanded[key] ? 'Show less' : 'Show more') + '</a>';
  return h + '</div>';
}
function sugHtml(p, idx, used){
  var sg = suggest(p, idx, used), h = "";
  if(!sg.length) return '<div class="note">Nothing in the Excel file shares words with this slide. Use Browse all rows or write your own text.</div>';
  var strongL = sg.filter(function(s){ return s.kind === "strong"; }).slice(0, 3), primary, secondary;
  if(strongL.length){ primary = strongL; secondary = sg.filter(function(s){ return primary.indexOf(s) < 0; }); }
  else { primary = sg.slice(0, 3); secondary = sg.slice(3); }
  var winner = (primary[0].kind === "strong" && (primary.length === 1 || primary[1].kind !== "strong" || (primary[0].sc - primary[1].sc) >= 0.15)) ? primary[0].id : "";
  h += '<div class="lbl">' + (p.alt ? 'Pick a match only if you want to change the current alt text' : 'Suggested for this picture') + '</div>';
  primary.forEach(function(s){ h += sugItem(p, s, used, s.id === winner); });
  if(secondary.length){
    h += '<a href="#" class="more" data-moresug="' + p.i + '">' + (moreSug[p.i] ? 'Hide the weaker suggestions' : 'Show ' + secondary.length + ' weaker suggestion' + (secondary.length > 1 ? 's' : '')) + '</a>';
    if(moreSug[p.i]) secondary.forEach(function(s){ h += sugItem(p, s, used, false); });
  }
  return h;
}function terms(q){
  var s = String(q).toLowerCase(), out = [];
  // a figure number such as "Figure 14.F03" or "14.f03" finds that row in its chapter (the word Figure is dropped)
  s = s.replace(/\b(?:fig(?:ure)?|table|box|eq(?:uation)?)\.?\s*(\d{1,3})\.([a-z]{1,3}\s?\d{1,3})/g, function(m, ch, id){ out.push("@" + parseInt(ch, 10) + "." + id.replace(/\s/g, "").replace(/^([a-z]*)0*(\d+)$/, "$1$2")); return " "; });
  s.split(/\s+/).forEach(function(x){ if(x) out.push(x); });
  return out;
}
function rowKey(o){ return "@" + (/^\d+$/.test(o.ch) ? parseInt(o.ch, 10) : String(o.ch).toLowerCase()) + "." + String(o.ex).toLowerCase().replace(/\s/g, "").replace(/^([a-z]*)0*(\d+)$/, "$1$2"); }
function rowsInner(p, used, q){
  var ts = terms(q), need = Math.max(1, Math.ceil(ts.length / 2)), list = [];
  ALLOPTS.forEach(function(o){
    var t = (o.id + " " + o.ch + " " + o.text).toLowerCase(), rk = rowKey(o), c = 0;
    ts.forEach(function(x){ if(x.charAt(0) === "@") { if(rk === x) c += 3; } else if(t.indexOf(x) >= 0) c++; });
    if(!ts.length || c >= need) list.push({o: o, c: c, k: o.ord});
  });
  if(ts.length) list.sort(function(a, b){ return (b.c - a.c) || (a.k - b.k); });
  if(!list.length) return '<div class="note">No rows contain those words.</div>';
  return list.map(function(r){ var o = r.o, u = used[o.id]; return '<button class="opt' + (u ? ' used' : '') + '" data-i="' + p.i + '" data-id="' + esc(o.id) + '">' + usedBadge(u) + '<b>' + esc(o.id) + '</b>' + esc(cut(o.text, 260)) + '</button>'; }).join("");
}
function rowsHtml(p, used){
  var h = '<div class="acts"><button class="small" data-browse="' + p.i + '">' + (openBrowse[p.i] ? 'Hide all rows' : 'Browse all rows') + '</button>';
  if(p.sw) h += '<button class="small" data-sw="' + p.i + '">Search these words: ' + esc(p.swt || p.sw) + '</button>';
  h += '</div>';
  if(openBrowse[p.i]){
    h += '<div class="rowsbox"><input class="flt" data-flt="' + p.i + '" type="search" value="' + esc(fltText[p.i] || "") + '" placeholder="Search the whole Excel file, for example crash or Figure 14.F03" aria-label="Search the Excel rows"><div class="rows">' + rowsInner(p, used, fltText[p.i] || "") + '</div></div>';
  }
  return h;
}
function ownBoxHtml(p){
  var s = S[p.i], h = '<div class="lbl">' + (p.alt ? 'Or write new alt text to replace it' : 'Or write your own alt text') + '</div>';
  h += '<textarea class="owntext" data-own="' + p.i + '" placeholder="Describe what the picture shows" aria-label="Your own alt text">' + esc(s.text) + '</textarea>';
  if(s.err) h += '<div class="err">' + esc(s.err) + '</div>';
  h += '<div class="acts"><button class="small" data-useown="' + p.i + '">Use this text</button><button class="small" data-uns="' + p.i + '">Not sure</button><button class="small" data-deco="' + p.i + '">Mark as decorative</button></div>';
  return h;
}
function keepNote(p, rep){ return p.alt ? '<div class="note">You chose to replace the alt text it had: ' + esc(cut(p.alt, 120)) + '</div>' : ''; }
function stateHtml(p, rep){
  var s = S[p.i], h = "";
  if(s.mode === "excel") h = '<div class="chosen"><b>' + esc(s.id) + '</b> &middot; ' + esc(byId[s.id].text) + '<a href="#" data-clear="' + p.i + '">Change</a><a href="#" data-edit="' + p.i + '">Edit this text</a></div>' + keepNote(p, rep);
  else if(s.mode === "own") h = '<div class="state own">Your text: ' + esc(s.text) + '<a href="#" data-clear="' + p.i + '">Change</a></div>' + keepNote(p, rep);
  else if(s.mode === "unsure") h = '<div class="state uns">Marked to check later. Nothing will be written for this picture.</div><input class="notein" data-note="' + p.i + '" type="text" value="' + esc(s.note) + '" placeholder="Add a note, for example ask the author" aria-label="Note" style="margin-top:8px"><div class="acts"><button class="small" data-clear="' + p.i + '">Undo</button></div>';
  else if(s.mode === "deco") h = '<div class="state deco">Marked decorative. It gets no alt text and screen readers skip it.<a href="#" data-clear="' + p.i + '">Undo</a></div>' + keepNote(p, rep);
  return h;
}
var PILL = {none: ["need", "Needs a match"], has: ["has", "Has alt text"], excel: ["ok", "Matched"], own: ["own", "Your own text"], unsure: ["uns", "Not sure"], deco: ["deco", "Decorative"]};
function render(){
  var used = usedMap(), c = {none: 0, has: 0, excel: 0, own: 0, unsure: 0, deco: 0}, rep = true;
  PICS.forEach(function(p, idx){
    var card = cards[idx], main = card.querySelector(".main"), s = S[p.i], m = s.mode, key = (m === "none" && p.alt && !p.ai) ? "has" : m, h = "";
    c[key]++;
    h += '<div class="head"><b>Slide ' + p.slide + ' &middot; ' + esc(p.name) + '</b><span class="pill ' + PILL[key][0] + '">' + PILL[key][1] + '</span></div>';
    if(p.title) h += '<div class="ctx">' + esc(p.title) + '</div>';
    if(p.label) h += '<div class="lab">Label on the slide: <b>' + esc(p.label) + '</b></div>';
    if(p.fig) h += '<div class="lab">Figure number on the slide: <b>' + esc(p.fig) + '</b>' + (p.figmiss ? ' - not in the Excel file, so it could not be matched by number' : '') + '</div>';
    if(p.alt) h += '<div class="has">' + (p.ai ? 'AI-generated alt text already there (not yet reviewed): ' : 'Current alt text: ') + esc(p.alt) + '</div>';
    if(m === "none") h += sugHtml(p, idx, used) + rowsHtml(p, used) + ownBoxHtml(p);
    else h += stateHtml(p, rep);
    main.innerHTML = h;
    card.className = "card" + (m === "excel" || m === "deco" ? " ok" : m === "own" ? " own" : m === "unsure" ? " uns" : "");
    var vis = true, hay = (p.slide + " " + p.name + " " + p.title + " " + p.label + " " + p.alt).toLowerCase();
    if(filt === "need" && key !== "none") vis = false; if(filt === "done" && m !== "excel" && m !== "deco") vis = false;
    if(filt === "own" && m !== "own") vis = false; if(filt === "uns" && m !== "unsure") vis = false; if(filt === "has" && key !== "has") vis = false;
    if(query && hay.indexOf(query) < 0) vis = false;
    card.style.display = vis ? "" : "none";
    card.setAttribute("data-vis", vis ? "1" : "0");
  });
  var total = PICS.length, done = c.excel + c.own + c.deco + c.has, pc = total ? Math.round(done / total * 100) : 0;
  document.getElementById("cnt").textContent = done + " of " + total + " done" + (c.unsure ? " - " + c.unsure + " not sure" : "");
  document.getElementById("pct").textContent = pc + "%";
  document.getElementById("fill").style.width = pc + "%";
  document.getElementById("n-all").textContent = total; document.getElementById("n-need").textContent = c.none;
  document.getElementById("n-done").textContent = c.excel + c.deco; document.getElementById("n-own").textContent = c.own;
  document.getElementById("n-uns").textContent = c.unsure; document.getElementById("n-has").textContent = c.has;
  document.getElementById("none").hidden = cards.some(function(x){ return x.getAttribute("data-vis") === "1"; });
  if(lbIdx >= 0) renderLb();
}var cfYes = null;
function askReplace(i, newText, yes){
  var p = PICS[i];
  if(!p.alt){ yes(); return; }
  cfYes = yes;
  document.getElementById("cfsub").textContent = "Slide " + p.slide + " - " + p.name + " already has alt text.";
  document.getElementById("cfold").textContent = p.alt;
  document.getElementById("cfnew").textContent = newText;
  document.getElementById("cf").hidden = false;
  document.getElementById("cfyes").focus();
}
function closeAsk(ok){ var f = cfYes; cfYes = null; document.getElementById("cf").hidden = true; if(ok && f) f(); }
function doPick(i, id){ askReplace(i, byId[id].dec ? "(marked decorative in the Excel file - no alt text, screen readers skip the picture)" : byId[id].full, function(){ S[i] = {mode: "excel", id: id, text: "", note: "", err: ""}; render(); }); }
function lbImg(){ return document.getElementById("lbimg"); }
function applyScale(){ var im = lbImg(); if(lbBase){ im.style.width = Math.round(lbBase * lbScale) + "px"; } document.getElementById("lbzoom").textContent = Math.round(lbScale * 100) + "%"; }
function loadLbImage(){
  var p = PICS[lbIdx], im = lbImg(), box = document.getElementById("lbimgbox");
  lbScale = 1;
  document.getElementById("lbviewpic").classList.toggle("on", lbView === "pic");
  document.getElementById("lbviewpic").setAttribute("aria-pressed", lbView === "pic" ? "true" : "false");
  document.getElementById("lbviewslide").classList.toggle("on", lbView === "slide");
  document.getElementById("lbviewslide").setAttribute("aria-pressed", lbView === "slide" ? "true" : "false");
  if(lbView === "slide"){
    var data = SLIDETHUMBS[String(p.slide)];
    if(!data){ box.innerHTML = '<div class="noimg">No preview of the whole slide</div>'; return; }
    box.innerHTML = '<img id="lbimg" alt="">'; im = lbImg();
    im.onload = function(){ lbBase = Math.min(im.naturalWidth, Math.round(window.innerWidth * 0.85)); applyScale(); };
    im.src = "data:image/png;base64," + data; if(im.complete && im.naturalWidth){ im.onload(); }
  } else {
    var src = cards[lbIdx].querySelector(".th img");
    if(!src){ box.innerHTML = '<div class="noimg">No preview</div>'; return; }
    box.innerHTML = '<img id="lbimg" alt="">'; im = lbImg();
    im.onload = function(){ lbBase = Math.min(im.naturalWidth, Math.round(window.innerWidth * 0.85)); applyScale(); };
    im.src = src.src; if(im.complete && im.naturalWidth){ im.onload(); }
  }
}
function setLbView(v){ lbView = v; loadLbImage(); }
function openZoom(idx){
  if(!hasAnyPreview(idx)) return;
  lbIdx = idx; lbQuery = ""; lbView = cards[idx].querySelector(".th img") ? "pic" : "slide";
  loadLbImage();
  document.getElementById("lb").hidden = false; renderLb();
}
function renderLb(){
  var p = PICS[lbIdx], used = usedMap(), s = S[p.i], rep = true;
  document.getElementById("lbtitle").textContent = "Slide " + p.slide + " - " + p.name;
  document.getElementById("lbsub").textContent = p.label ? ("Label: " + p.label) : p.title;
  var row = document.getElementById("lbsearchrow");
  if(s.mode === "none"){
    row.hidden = false;
    var si = document.getElementById("lbsearch"); if(document.activeElement !== si) si.value = lbQuery;
    if(lbQuery.trim()){
      document.getElementById("lbpick").innerHTML = '<div class="rows">' + rowsInner(p, used, lbQuery) + '</div><div class="note">To write your own text or mark it not sure, close this window and use the card.</div>';
    } else {
      document.getElementById("lbpick").innerHTML = sugHtml(p, lbIdx, used) + '<div class="note">To write your own text or mark it not sure, close this window and use the card.</div>';
    }
  } else {
    row.hidden = true;
    document.getElementById("lbpick").innerHTML = stateHtml(p, rep);
  }
}
function closeZoom(){ lbIdx = -1; document.getElementById("lb").hidden = true; }
function hasAnyPreview(idx){ return !!(cards[idx].querySelector(".th img") || SLIDETHUMBS[String(PICS[idx].slide)]); }
function stepZoom(d){ var n = lbIdx + d; while(n >= 0 && n < PICS.length && !hasAnyPreview(n)) n += d; if(n >= 0 && n < PICS.length) openZoom(n); }
document.addEventListener("click", function(e){
  var t = e.target;
  if(t.id === "cfyes"){ closeAsk(true); return; }
  if(t.id === "cfno" || t.id === "cf"){ closeAsk(false); return; }
  var th = t.closest(".th"); if(th){ openZoom(parseInt(th.parentNode.getAttribute("data-i"), 10)); return; }
  var b = t.closest(".sg, .opt, .use"); if(b){ doPick(b.getAttribute("data-i"), b.getAttribute("data-id")); return; }
  var ms = t.closest("[data-moresug]"); if(ms){ e.preventDefault(); var i0 = ms.getAttribute("data-moresug"); moreSug[i0] = !moreSug[i0]; render(); return; }
  var m = t.closest("[data-more]"); if(m){ e.preventDefault(); var k = m.getAttribute("data-more"); expanded[k] = !expanded[k]; render(); return; }
  var sw = t.closest("[data-sw]"); if(sw){ var i1 = sw.getAttribute("data-sw"); openBrowse[i1] = true; fltText[i1] = PICS[i1].sw; render(); return; }
  var br = t.closest("[data-browse]"); if(br){ var i2 = br.getAttribute("data-browse"); openBrowse[i2] = !openBrowse[i2]; render(); return; }
  var uo = t.closest("[data-useown]"); if(uo){ var i3 = uo.getAttribute("data-useown"), tv = (S[i3].text || "").replace(/\s*[\r\n]+\s*/g, " ").trim(); if(!tv){ S[i3].err = "Type some text first"; render(); return; } askReplace(i3, tv, function(){ S[i3] = {mode: "own", id: "", text: tv, note: "", err: ""}; render(); }); return; }
  var un = t.closest("[data-uns]"); if(un){ var i4 = un.getAttribute("data-uns"); S[i4] = {mode: "unsure", id: "", text: S[i4].text || "", note: "", err: ""}; render(); return; }
  var dc = t.closest("[data-deco]"); if(dc){ var i5 = dc.getAttribute("data-deco"); askReplace(i5, "(decorative - no alt text, screen readers skip the picture)", function(){ S[i5] = {mode: "deco", id: "", text: S[i5].text || "", note: "", err: ""}; render(); }); return; }
  var ed = t.closest("[data-edit]"); if(ed){ e.preventDefault(); var i6 = ed.getAttribute("data-edit"); S[i6] = {mode: "none", id: "", text: byId[S[i6].id].full, note: "", err: ""}; render(); return; }
  var a = t.closest("[data-clear]"); if(a){ e.preventDefault(); var i7 = a.getAttribute("data-clear"); S[i7] = {mode: "none", id: "", text: (S[i7].mode === "own" ? S[i7].text : S[i7].text) || "", note: "", err: ""}; render(); return; }
  if(t.id === "lbclose" || t.id === "lb"){ closeZoom(); return; }
  if(t.id === "lbviewpic"){ setLbView("pic"); return; }
  if(t.id === "lbviewslide"){ setLbView("slide"); return; }
  if(t.id === "lbplus"){ lbScale = Math.min(6, lbScale * 1.25); applyScale(); return; }
  if(t.id === "lbminus"){ lbScale = Math.max(0.3, lbScale / 1.25); applyScale(); return; }
  if(t.id === "lbprev"){ stepZoom(-1); return; }
  if(t.id === "lbnext"){ stepZoom(1); return; }
});
document.addEventListener("keydown", function(e){
  if(!document.getElementById("cf").hidden){ if(e.key === "Escape"){ closeAsk(false); } return; }
  if(lbIdx < 0) return;
  if(e.key === "Escape"){ closeZoom(); }
  else if(e.key === "ArrowRight"){ stepZoom(1); }
  else if(e.key === "ArrowLeft"){ stepZoom(-1); }
  else if(e.key === "+" || e.key === "="){ lbScale = Math.min(6, lbScale * 1.25); applyScale(); }
  else if(e.key === "-"){ lbScale = Math.max(0.3, lbScale / 1.25); applyScale(); }
});
document.addEventListener("input", function(e){
  var el = e.target; if(!el.closest) return;
  if(el.id === "lbsearch"){
    lbQuery = el.value;
    renderLb();
    var si2 = document.getElementById("lbsearch"); si2.focus(); si2.setSelectionRange(si2.value.length, si2.value.length);
    return;
  }
  var f = el.closest(".flt");
  if(f){
    var i8 = f.getAttribute("data-flt"); fltText[i8] = f.value;
    f.parentNode.querySelector(".rows").innerHTML = rowsInner(PICS[i8], usedMap(), f.value);
    return;
  }
  var ow = el.closest(".owntext"); if(ow){ var i9 = ow.getAttribute("data-own"); S[i9].text = ow.value; S[i9].err = ""; return; }
  var nt = el.closest(".notein"); if(nt){ S[nt.getAttribute("data-note")].note = nt.value; }
});
[].slice.call(document.querySelectorAll(".chip")).forEach(function(c){ c.addEventListener("click", function(){ filt = c.getAttribute("data-f"); [].slice.call(document.querySelectorAll(".chip")).forEach(function(x){ x.classList.toggle("on", x === c); }); render(); }); });
document.getElementById("q").addEventListener("input", function(e){ query = e.target.value.toLowerCase(); render(); });
function list(items){ return items.length ? "<ul>" + items.map(function(x){ return "<li>" + esc(x) + "</li>"; }).join("") + "</ul>" : ""; }
function post(path, body){
  return fetch(path, {method: "POST", headers: {"Content-Type": "application/json", "X-Token": TOKEN}, body: JSON.stringify(body || {})})
    .then(function(r){ return r.json().then(function(j){ if(!r.ok){ throw new Error(j.error || r.statusText); } return j; }); });
}
function showPanel(html, isErr){ var p = document.getElementById("panel"); p.hidden = false; p.className = isErr ? "err" : ""; p.innerHTML = html; }
function rowsPayload(){
  return PICS.map(function(p){ var s = S[p.i]; var ro = s.mode === "excel" ? byId[s.id] : null; return {Chapter: ro ? ro.ch : CHAPTER, Slide: p.slide, ShapeId: p.sid, PictureName: p.name, ExcelId: ro ? ro.ex : "", Mode: s.mode, CustomText: s.mode === "own" ? s.text : "", Note: s.mode === "unsure" ? s.note : ""}; });
}
document.getElementById("apply").addEventListener("click", function(){
  if(location.protocol === "file:" || !TOKEN){
    showPanel("<h2>This page needs the CMD window</h2><p>This is only a saved copy. Run AltText Matcher.cmd, choose Yes, and use the page it opens (its address starts with http://localhost). You can still press Save mapping here.</p>", true); return;
  }
  var n = {excel: 0, own: 0, deco: 0, unsure: 0}; PICS.forEach(function(p){ var m = S[p.i].mode; if(n[m] !== undefined) n[m]++; });
  if(!(n.excel + n.own + n.deco)){ showPanel("<h2>Nothing is done yet</h2><p>Pick a match, write your own text or mark a picture decorative first.</p>", true); return; }
  var rep = true;
  if(!confirm("From Excel: " + n.excel + "\nYour own text: " + n.own + "\nDecorative: " + n.deco + "\nNot sure (nothing written): " + n.unsure + "\n\nThe alt text is written into a copy of the PowerPoint. The original is not changed.\n" + "Pictures you gave a choice get it, also when they already had alt text (you confirmed each one). Pictures without a choice keep their alt text." + "\n\nContinue?")) return;
  var btn = document.getElementById("apply"); btn.disabled = true; document.getElementById("msg").textContent = "Writing...";
  post("/apply", {mapping: rowsPayload(), replace: rep}).then(function(r){
    document.getElementById("msg").textContent = "Done";
    showPanel("<h2>Saved: " + esc(r.OutFile) + "</h2><p>From Excel: <b>" + r.Filled + "</b> &middot; your own text: <b>" + r.Own + "</b> &middot; marked decorative: <b>" + r.Decor + "</b> &middot; replaced: <b>" + r.Replaced + "</b> &middot; kept as they were: <b>" + r.Kept + "</b> &middot; problems: <b>" + r.Problems + "</b></p>" +
      (r.Kept ? "<p>Kept their existing alt text:</p>" + list(r.KeptList) : "") +
      (r.Issues.length ? "<p>Problems:</p>" + list(r.Issues) : "") +
      (r.NotSure.length ? "<p>Marked not sure, nothing written:</p>" + list(r.NotSure) : "") +
      (r.Unmatched.length ? "<p>Still without a match (write these by hand):</p>" + list(r.Unmatched) : "") +
      '<div class="row"><button class="btn main-btn" id="openf">Open the folder</button><button class="btn" id="finish">Finish and close the helper</button><button class="btn" id="hide">Hide this</button></div>', false);
    document.getElementById("openf").onclick = function(){ post("/open"); };
    document.getElementById("hide").onclick = function(){ document.getElementById("panel").hidden = true; };
    document.getElementById("finish").onclick = function(){ post("/quit").then(function(){ document.getElementById("msg").textContent = "The helper window has stopped. You can close this page."; document.getElementById("apply").disabled = true; }); };
  }).catch(function(e){
    document.getElementById("msg").textContent = "";
    showPanel("<h2>It did not work</h2><p>" + esc(e.message) + "</p><p>If the CMD window was closed, run AltText Matcher.cmd again. If the deck or its copy is open in PowerPoint, close it and press the button again.</p>", true);
  }).then(function(){ btn.disabled = false; });
});
function q(v){ return '"' + String(v).replace(/"/g, '""') + '"'; }
document.getElementById("save").addEventListener("click", function(){
  var rep = "TRUE";
  var lines = ["Chapter,Slide,ShapeId,PictureName,ExcelId,Replace,Mode,CustomText,Note"];
  rowsPayload().forEach(function(r){ lines.push([q(r.Chapter), r.Slide, r.ShapeId, q(r.PictureName), q(r.ExcelId), rep, q(r.Mode), q(String(r.CustomText).replace(/\s*[\r\n]+\s*/g, " ")), q(String(r.Note).replace(/\s*[\r\n]+\s*/g, " "))].join(",")); });
  var blob = new Blob([String.fromCharCode(65279) + lines.join("\r\n")], {type: "text/csv"});
  var a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = "AltText_Mapping " + DECK + ".csv"; document.body.appendChild(a); a.click(); document.body.removeChild(a);
});
render();
</script></body></html>
'@
    $html = $html.Replace('__DECK__', (Html $deckName)).Replace('__CHAPTER__', (Html $chap)).Replace('__CARDS__', $cards.ToString())
    $slideThumbJson = New-Object System.Collections.Generic.List[string]
    foreach ($k in $slideThumbs.Keys) { $slideThumbJson.Add((& $js $k) + ':' + (& $js $slideThumbs[$k])) }
    $html = $html.Replace('__OPTS__', ('[' + ($optJson -join ',') + ']')).Replace('__OTHERS__', ('[' + ($othJson -join ',') + ']')).Replace('__PICS__', ('[' + ($picJson -join ',') + ']')).Replace('__SLIDETHUMBS__', ('{' + ($slideThumbJson -join ',') + '}'))
    $outHtml = Join-Path (Split-Path -Parent $pptx) ($deckName + ' AltText Review.html')
    [IO.File]::WriteAllText($outHtml, $html.Replace('__TOKEN__', ''), (New-Object Text.UTF8Encoding($false)))
    Write-Host ''
    Write-Host ('Review page saved as a backup copy: ' + $outHtml)
    if ($env:ALTMATCH_NOSERVE) { return }
    $token = [guid]::NewGuid().ToString('N')
    $lookup = Get-Lookup $rows
    Start-ReviewServer $html.Replace('__TOKEN__', $token) $token $pptx $lookup
}

# =====================================================================
# Apply (used by Step 2 and by the button on the review page)
# =====================================================================
function Invoke-Apply([string]$pptx, $lookup, $map, [string]$outPath, [bool]$replace = $false) {
    Copy-Item -LiteralPath $pptx -Destination $outPath -Force
    $issues = New-Object System.Collections.Generic.List[string]
    $keptList = New-Object System.Collections.Generic.List[string]
    $unmatched = New-Object System.Collections.Generic.List[string]
    $notSure = New-Object System.Collections.Generic.List[string]
    $filled = 0; $own = 0; $kept = 0; $decor = 0; $missing = 0; $replaced = 0

    $zip = [IO.Compression.ZipFile]::Open($outPath, 'Update')
    try {
        $parts = Get-SlideParts $zip
        $docs = @{}
        foreach ($row in $map) {
            $slideNo = [int]$row.Slide
            if ($slideNo -lt 1 -or $slideNo -gt $parts.Count) { continue }
            if (([string]$row.ShapeId) -notmatch '^\d+$') { continue }
            if (-not $docs.ContainsKey($slideNo)) {
                $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.LoadXml((Read-ZipText $zip $parts[$slideNo - 1]))
                $docs[$slideNo] = @{ Doc = $d; Ns = (New-SlideNs $d); Changed = $false }
            }
            $entry = $docs[$slideNo]
            $pic = $entry.Doc.SelectSingleNode('//p:pic[p:nvPicPr/p:cNvPr/@id="' + $row.ShapeId + '"]', $entry.Ns)
            if (-not $pic) { $issues.Add("Slide $slideNo : picture id $($row.ShapeId) not found, skipped"); $missing++; continue }
            $nv = $pic.SelectSingleNode('p:nvPicPr/p:cNvPr', $entry.Ns)
            $where = "Slide $slideNo $($nv.GetAttribute('name'))"

            $mode = ''
            if ($row.PSObject.Properties['Mode'] -and $row.Mode) { $mode = ([string]$row.Mode).ToLowerInvariant() }
            elseif ($row.ExcelId) { $mode = 'excel' }
            else { $mode = 'none' }
            $custom = ''; if ($row.PSObject.Properties['CustomText']) { $custom = [string]$row.CustomText }
            $noteTxt = ''; if ($row.PSObject.Properties['Note']) { $noteTxt = ([string]$row.Note).Trim() }

            if ($mode -eq 'unsure') { if ($noteTxt) { $notSure.Add("$where ($noteTxt)") } else { $notSure.Add($where) }; continue }
            if ($mode -eq 'none') { $curAlt = $nv.GetAttribute('descr'); if (-not ($curAlt -and $curAlt.Trim().Length -gt 0)) { $unmatched.Add($where) }; continue }

            $alt = ''; $isDeco = $false; $src = ''
            if ($mode -eq 'excel') {
                $key = (Norm-Chapter $row.Chapter) + '|' + $row.ExcelId
                if (-not $lookup.ContainsKey($key)) { $issues.Add("$where : Excel row $($row.ExcelId) not found, skipped"); $missing++; continue }
                $x = $lookup[$key]
                $src = "Excel row $($row.ExcelId)"
                if ($x.Decorative) { $isDeco = $true }
                else { $alt = $x.Short; if ($x.Long) { $alt = ($alt + ' ' + $x.Long).Trim() } }
            }
            elseif ($mode -eq 'own') { $alt = $custom.Trim(); $src = 'your own text' }
            elseif ($mode -eq 'deco') { $isDeco = $true; $src = 'decorative' }
            else { $issues.Add("$where : unknown choice '$mode', skipped"); $missing++; continue }

            if (-not $isDeco -and -not $alt) { $issues.Add("$where : $src has no alt text, skipped"); $missing++; continue }

            $existing = $nv.GetAttribute('descr')
            $hasExisting = [bool]($existing -and $existing.Trim().Length -gt 0)
            if ($hasExisting -and -not $replace) { $kept++; $keptList.Add("$where ($src not used)"); continue }

            if ($isDeco) {
                $nv.SetAttribute('descr', '')
                $ext = $nv.SelectSingleNode('a:extLst', $entry.Ns)
                if (-not $ext) { $ext = $entry.Doc.CreateElement('a', 'extLst', $NS_A); $null = $nv.AppendChild($ext) }
                if (-not $ext.SelectSingleNode('a:ext[@uri="' + $DEC_URI + '"]', $entry.Ns)) {
                    $e = $entry.Doc.CreateElement('a', 'ext', $NS_A)
                    $e.SetAttribute('uri', $DEC_URI)
                    $dnode = $entry.Doc.CreateElement('adec', 'decorative', $NS_DEC)
                    $dnode.SetAttribute('val', '1')
                    $null = $e.AppendChild($dnode)
                    $null = $ext.AppendChild($e)
                }
                if ($hasExisting) { $replaced++ }
                $decor++
            } else {
                $nv.SetAttribute('descr', $alt)
                $oldDec = $nv.SelectSingleNode('a:extLst/a:ext[@uri="' + $DEC_URI + '"]', $entry.Ns)
                if ($oldDec) { $null = $oldDec.ParentNode.RemoveChild($oldDec) }
                if ($hasExisting) { $replaced++ }
                if ($mode -eq 'own') { $own++ } else { $filled++ }
            }
            $entry.Changed = $true
        }
        foreach ($k in $docs.Keys) {
            if ($docs[$k].Changed) {
                $name = $parts[$k - 1]
                $e = $zip.GetEntry($name); if ($e) { $e.Delete() }
                $ne = $zip.CreateEntry($name)
                $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
                $docs[$k].Doc.Save($sw)
                $sw.Close()
            }
        }
    } finally { $zip.Dispose() }

    $summary = "From Excel: $filled ; your own text: $own ; decorative: $decor ; replaced: $replaced ; kept as they were: $kept ; problems: $missing ; marked not sure: $($notSure.Count) ; still without a match: $($unmatched.Count)"
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add($summary)
    foreach ($k in $keptList) { $report.Add('Kept its existing alt text: ' + $k) }
    foreach ($i in $issues) { $report.Add($i) }
    if ($notSure.Count) { $report.Add('Marked not sure, nothing written: ' + ($notSure -join ' ; ')) }
    if ($unmatched.Count) { $report.Add('Still without a match (write these by hand): ' + ($unmatched -join ' ; ')) }
    Add-WorkLog 'Alt Text Tool' $pptx $outPath

    return [pscustomobject]@{
        Filled = $filled; Own = $own; Replaced = $replaced; Decor = $decor; Kept = $kept; Problems = $missing
        KeptList = $keptList.ToArray(); Issues = $issues.ToArray(); Unmatched = $unmatched.ToArray(); NotSure = $notSure.ToArray()
        OutPath = $outPath; OutFile = (Split-Path -Leaf $outPath); Summary = $summary
    }
}

function Get-Lookup($rows) {
    $lookup = @{}
    foreach ($r in $rows) { $lookup[($r.Chapter + '|' + $r.Id)] = $r }
    return $lookup
}

function Get-OutPath([string]$pptx) {
    return (Join-Path (Split-Path -Parent $pptx) ([IO.Path]::GetFileNameWithoutExtension($pptx) + ' AltText.pptx'))
}

# =====================================================================
# Step 2: apply a saved mapping file
# =====================================================================
function Apply-AltText([string]$pptx, [string]$xlsx, [string]$mapPath) {
    $rows = Get-AltRows $xlsx
    $lookup = Get-Lookup $rows
    $map = @(Import-Csv -LiteralPath $mapPath -Encoding UTF8)
    if ($map.Count -eq 0) { throw 'The mapping file is empty.' }
    $rep = $false
    if ($map[0].PSObject.Properties['Replace'] -and ([string]$map[0].Replace).ToUpperInvariant() -eq 'TRUE') { $rep = $true }
    $res = Invoke-Apply $pptx $lookup $map (Get-OutPath $pptx) $rep
    Write-Host $res.Summary
    foreach ($k in $res.KeptList) { Write-Host ('Kept its existing alt text: ' + $k) }
    foreach ($i in $res.Issues) { Write-Host $i }
    if ($res.NotSure.Count) { Write-Host ('Marked not sure, nothing written: ' + ($res.NotSure -join ' ; ')) }
    if ($res.Unmatched.Count) { Write-Host ('Still without a match (write these by hand): ' + ($res.Unmatched -join ' ; ')) }
    Write-Host ''
    Write-Host ('Saved: ' + $res.OutPath)
}

# =====================================================================
# The review page is served from this window so its button can write the deck
# =====================================================================
function Start-ReviewServer([string]$html, [string]$token, [string]$pptx, $lookup) {
    $listener = $null; $port = 0
    foreach ($try in 1..40) {
        $port = Get-Random -Minimum 49200 -Maximum 49900
        $l = New-Object System.Net.HttpListener
        $l.Prefixes.Add("http://localhost:$port/")
        try { $l.Start(); $listener = $l; break } catch { $l.Close() }
    }
    if (-not $listener) { throw 'Could not start the local page server.' }
    $url = "http://localhost:$port/"
    $outPath = Get-OutPath $pptx
    if ($env:ALTMATCH_PORTFILE) { Set-Content -LiteralPath $env:ALTMATCH_PORTFILE -Value "$port|$token" -Encoding ASCII }
    else { Start-Process $url }
    Write-Host ''
    Write-Host ('The review page is open at ' + $url)
    Write-Host 'Leave this window open while you use the page. Close it when you are finished.'
    $quit = $false
    while ($listener.IsListening -and -not $quit) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request; $res = $ctx.Response
        $status = 200; $ctype = 'application/json; charset=utf-8'; $body = ''
        try {
            $path = $req.Url.AbsolutePath
            if ($req.HttpMethod -eq 'GET' -and $path -eq '/') { $ctype = 'text/html; charset=utf-8'; $body = $html }
            elseif ($req.HttpMethod -eq 'POST') {
                if ($req.Headers['X-Token'] -ne $token) { $status = 403; $body = '{"error":"Not allowed"}' }
                else {
                    $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)
                    $raw = $sr.ReadToEnd(); $sr.Close()
                    if ($path -eq '/apply') {
                        $data = $raw | ConvertFrom-Json
                        $result = Invoke-Apply $pptx $lookup @($data.mapping) $outPath ([bool]$data.replace)
                        Write-Host ('Applied: ' + $result.Summary)
                        $body = $result | ConvertTo-Json -Depth 4 -Compress
                    }
                    elseif ($path -eq '/open') {
                        if (Test-Path -LiteralPath $outPath) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $outPath + '"') }
                        $body = '{"ok":true}'
                    }
                    elseif ($path -eq '/quit') { $body = '{"ok":true}'; $quit = $true }
                    else { $status = 404; $body = '{"error":"Not found"}' }
                }
            }
            else { $status = 404; $body = '{"error":"Not found"}' }
        } catch {
            $status = 500
            $body = (@{ error = $_.Exception.Message } | ConvertTo-Json -Compress)
        }
        $bytes = [Text.Encoding]::UTF8.GetBytes($body)
        $res.StatusCode = $status; $res.ContentType = $ctype; $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes, 0, $bytes.Length)
        $res.OutputStream.Close()
    }
    $listener.Stop()
    Write-Host 'Page closed.'
}
# =====================================================================
# main
# =====================================================================
try {
    $mode = $env:ALTMATCH_MODE
    if (-not $mode) {
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $ans = [System.Windows.Forms.MessageBox]::Show($owner, "Yes = Step 1: make the review page (pick the matches)`r`nNo = Step 2: apply the saved matches to the deck`r`nCancel = close", 'Alt text matcher', 'YesNoCancel', 'Question')
        if ($ans -eq 'Yes') { $mode = '1' } elseif ($ans -eq 'No') { $mode = '2' } else { Write-Host 'Closed.'; $mode = '' }
    }
    if ($mode -eq '1' -or $mode -eq '2') {
        $pptx = $env:ALTMATCH_PPTX; if (-not $pptx) { $pptx = Pick-File 'Select the PowerPoint file' 'PowerPoint presentations (*.pptx)|*.pptx' }
        if (-not $pptx) { throw 'No PowerPoint file selected.' }
        $xlsx = $env:ALTMATCH_XLSX; if (-not $xlsx) { $xlsx = Pick-File 'Select the alt text Excel file' 'Excel files (*.xlsx)|*.xlsx' }
        if (-not $xlsx) { throw 'No Excel file selected.' }
        if ($mode -eq '1') {
            $chapter = $env:ALTMATCH_CHAPTER
            if (-not $chapter) {
                $guess = [regex]::Match([IO.Path]::GetFileNameWithoutExtension($pptx), '(\d{1,2})(?=\D*$)').Value
                $chapter = [Microsoft.VisualBasic.Interaction]::InputBox('Which chapter number is this deck? (as written in the Excel Chapter Number column, for example 05)', 'Chapter', $guess)
            }
            if (-not $chapter) { throw 'No chapter number given.' }
            Make-ReviewPage $pptx $xlsx $chapter
        } else {
            $mp = $env:ALTMATCH_MAP; if (-not $mp) { $mp = Pick-File 'Select the saved mapping file (AltText_Mapping ... .csv)' 'Mapping files (*.csv)|*.csv' }
            if (-not $mp) { throw 'No mapping file selected.' }
            Apply-AltText $pptx $xlsx $mp
        }
    }
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
}
