# SLID Macro - numbers repeated slide titles (1 of 4), renames shapes, fixes language tags.
# Saves a remediated copy next to the original. (Self-updates from GitHub via SLID_Update.ps1.)
# SLID MACRO SCRIPT
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
    $m = New-Object Xml.XmlNamespaceManager($xml.NameTable)
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
    foreach ($sp in $slideParts) {
        $m = $nss[$sp]
        $tree = $docs[$sp].SelectSingleNode('//p:cSld/p:spTree', $m)
        $items = New-Object System.Collections.Generic.List[object]
        $order = 0
        foreach ($c in $tree.ChildNodes) {
            if ($c.NodeType -ne 'Element') { continue }
            $order++
            $kind = ''; $nv = $null; $isCredit = $false
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
            $w = 0.0
            if ($off) {
                $y = [double]$off.GetAttribute('y'); $x = [double]$off.GetAttribute('x')
                $ext = $off.ParentNode.SelectSingleNode('a:ext', $m)
                if ($ext) { $w = [double]$ext.GetAttribute('cx') }
            }
            else {
                $lo = Get-LayoutOff $sp $c
                if ($lo) { $y = $lo[0]; $x = $lo[1]; $w = $lo[2] } else { $y = 1e12; $x = 0 }
            }
            $items.Add([pscustomobject]@{ Kind = $kind; Nv = $nv; Node = $c; RK = $rk; Y = [Math]::Round($y / 12700); X = $x; CX = (($x + $w / 2) / 12700); Order = $order })
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
            $renamed++
        }
        $same = $true
        for ($i = 0; $i -lt $sorted.Count; $i++) { if ($sorted[$i].Order -ne $items[$i].Order) { $same = $false; break } }
        if (-not $same) {
            $anchor = $tree.SelectSingleNode('p:extLst', $m)
            foreach ($it in $sorted) {
                $null = $tree.RemoveChild($it.Node)
                if ($anchor) { $null = $tree.InsertBefore($it.Node, $anchor) } else { $null = $tree.AppendChild($it.Node) }
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
            $cm = New-Object Xml.XmlNamespaceManager($core.NameTable)
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
}

if ($env:FIXDECK_FILE) {
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
