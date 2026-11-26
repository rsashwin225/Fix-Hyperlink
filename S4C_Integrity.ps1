# S4C INTEGRITY SCRIPT - compares the text of a file before and after a tool ran and writes an Excel report next to the fixed file.
# S4C-FILE-VERSION: 1.0
# Nothing in the original or the fixed file is changed here: both are only read.
# Every paragraph is compared exactly (spaces and punctuation included). A change a tool is allowed to make (for example the slide counter,
# a typed list label that became a real list, hyperlink text) is marked "Expected". Anything else is marked FLAG for a person to check.
# This script only defines functions; S4C_RunTool.ps1 loads it and calls Invoke-IntegrityRun after a tool has finished.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:IG_NsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$script:IG_NsP = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$script:IG_NsW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$script:IG_NsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function IG-ReadEntry($zip, [string]$name) {
    $e = $zip.GetEntry($name); if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}
function IG-Xml([string]$t) { $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.XmlResolver = $null; $d.LoadXml($t); return $d }
function IG-Target([string]$base, [string]$target) {
    if ($target.StartsWith('/')) { return $target.TrimStart('/') }
    $dir = ''; $i = $base.LastIndexOf('/'); if ($i -ge 0) { $dir = $base.Substring(0, $i) }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($s in (($dir + '/' + $target).Split('/'))) {
        if ($s -eq '' -or $s -eq '.') { continue }
        if ($s -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count - 1) } } else { $parts.Add($s) }
    }
    return ($parts -join '/')
}
function IG-Rels($zip, [string]$part) {
    $map = @{}
    $dir = ''; $i = $part.LastIndexOf('/'); if ($i -ge 0) { $dir = $part.Substring(0, $i + 1) }
    $t = IG-ReadEntry $zip ($dir + '_rels/' + $part.Substring($i + 1) + '.rels'); if (-not $t) { return $map }
    foreach ($m in [regex]::Matches($t, '<Relationship\s[^>]*?>')) {
        $id = [regex]::Match($m.Value, '\bId="([^"]+)"').Groups[1].Value
        $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
        $ty = [regex]::Match($m.Value, 'Type="([^"]+)"').Groups[1].Value
        $map[$id] = @{ Target = (IG-Target $part $tg); Type = $ty }
    }
    return $map
}

# ---------------------------------------------------------------- reading the paragraphs
function IG-PptParas($d) {
    $list = New-Object System.Collections.Generic.List[object]; $k = 0
    foreach ($p in $d.GetElementsByTagName('p', $script:IG_NsA)) {
        $k++; $sb = New-Object Text.StringBuilder; $link = $false
        foreach ($c in $p.ChildNodes) {
            if ($c.NodeType -ne 'Element') { continue }
            $ln = $c.LocalName
            if ($ln -eq 'r' -or $ln -eq 'fld') {
                foreach ($t in $c.ChildNodes) {
                    if ($t.NodeType -ne 'Element') { continue }
                    if ($t.LocalName -eq 't') { [void]$sb.Append($t.InnerText) }
                    elseif ($t.LocalName -eq 'rPr') { foreach ($q in $t.ChildNodes) { if ($q.NodeType -eq 'Element' -and $q.LocalName -eq 'hlinkClick') { $link = $true } } }
                }
            } elseif ($ln -eq 'br') { [void]$sb.Append("`n") }
        }
        $txt = $sb.ToString()
        if ($txt.Trim().Length -gt 0) { $list.Add([pscustomobject]@{ Idx = $k; Text = $txt; Link = $link }) }
    }
    return ,$list
}
function IG-ReadPptx([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $pres = IG-ReadEntry $zip 'ppt/presentation.xml'
        $rels = IG-Rels $zip 'ppt/presentation.xml'
        $out = New-Object System.Collections.Generic.List[object]; $n = 0
        foreach ($m in [regex]::Matches($pres, '<p:sldId\s[^>]*?>')) {
            $n++
            $sid = [regex]::Match($m.Value, '\bid="(\d+)"').Groups[1].Value
            $rid = [regex]::Match($m.Value, 'r:id="([^"]+)"').Groups[1].Value
            if (-not $rels.ContainsKey($rid)) { continue }
            $part = $rels[$rid].Target
            $t = IG-ReadEntry $zip $part; if (-not $t) { continue }
            $paras = IG-PptParas (IG-Xml $t)
            $out.Add([pscustomobject]@{ Key = ('s' + $sid); Label = ('Slide ' + $n); Paras = $paras })
            $sr = IG-Rels $zip $part
            foreach ($r in $sr.Values) {
                if ($r.Type -match '/notesSlide$') {
                    $nt = IG-ReadEntry $zip $r.Target
                    if ($nt) { $out.Add([pscustomobject]@{ Key = ('n' + $sid); Label = ('Slide ' + $n + ' notes'); Paras = (IG-PptParas (IG-Xml $nt)) }) }
                }
            }
        }
        return ,$out
    } finally { $zip.Dispose() }
}
function IG-ReadDocx([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $out = New-Object System.Collections.Generic.List[object]
        $names = @($zip.Entries | ForEach-Object { $_.FullName } | Where-Object { $_ -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$' } | Sort-Object)
        foreach ($nm in $names) {
            $t = IG-ReadEntry $zip $nm; if (-not $t) { continue }
            $d = IG-Xml $t
            $nsm = New-Object Xml.XmlNamespaceManager($d.NameTable); $nsm.AddNamespace('w', $script:IG_NsW)
            $list = New-Object System.Collections.Generic.List[object]; $k = 0
            foreach ($p in $d.GetElementsByTagName('p', $script:IG_NsW)) {
                $k++; $sb = New-Object Text.StringBuilder
                foreach ($x in $p.SelectNodes('.//w:t | .//w:tab | .//w:br', $nsm)) {
                    if ($x.LocalName -eq 't') { [void]$sb.Append($x.InnerText) } elseif ($x.LocalName -eq 'tab') { [void]$sb.Append("`t") } else { [void]$sb.Append("`n") }
                }
                $txt = $sb.ToString()
                if ($txt.Trim().Length -gt 0) { $list.Add([pscustomobject]@{ Idx = $k; Text = $txt; Link = [bool]$p.SelectSingleNode('.//w:hyperlink | .//w:fldSimple', $nsm) }) }
            }
            $label = switch -regex ($nm) { 'document' { 'Paragraph' } 'header' { 'Header' } 'footer' { 'Footer' } 'footnotes' { 'Footnote' } default { 'Endnote' } }
            $out.Add([pscustomobject]@{ Key = $nm; Label = $label; Paras = $list })
        }
        return ,$out
    } finally { $zip.Dispose() }
}

# ---------------------------------------------------------------- comparing
function IG-Sim([string]$a, [string]$b) {
    $max = [Math]::Max($a.Length, $b.Length); if ($max -eq 0) { return 1.0 }
    $m = [Math]::Min($a.Length, $b.Length); $p = 0
    while ($p -lt $m -and $a[$p] -ceq $b[$p]) { $p++ }
    $s = 0
    while ($s -lt ($m - $p) -and $a[$a.Length - 1 - $s] -ceq $b[$b.Length - 1 - $s]) { $s++ }
    return ($p + $s) / $max
}
function IG-Segments([string]$a, [string]$b) {
    $m = [Math]::Min($a.Length, $b.Length); $p = 0
    while ($p -lt $m -and $a[$p] -ceq $b[$p]) { $p++ }
    $s = 0
    while ($s -lt ($m - $p) -and $a[$a.Length - 1 - $s] -ceq $b[$b.Length - 1 - $s]) { $s++ }
    return @($a.Substring($p, $a.Length - $p - $s), $b.Substring($p, $b.Length - $p - $s), $p)
}
function IG-Strip([string]$s) { return ($s -replace '\s', '') }
function IG-Show([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s.Replace("`t", '<tab>').Replace("`n", '<line break>')
    $s = [regex]::Replace($s, '^( +)', { param($m) '<' + $m.Groups[1].Value.Length + ' space(s)>' })
    $s = [regex]::Replace($s, '( +)$', { param($m) '<' + $m.Groups[1].Value.Length + ' space(s)>' })
    if ($s.Length -eq 0) { return '(nothing)' }
    return $s
}
$script:IG_LabelRx = '^\s*(?:\(?[IVXLCDMivxlcdm]{1,6}[\.\)]|\(?[A-Za-z][\.\)]|\(?\d{1,3}[\.\)]|[\u2022\u00B7\u25CF\u25AA\u2013\u2014\*\-]|[\uF0B7\uF0A7\uF0D8\uF0FC])\s*$'

$script:IG_LabelStart = '^\s*(?:\(?[IVXLCDMivxlcdm]{1,6}[\.\)]|\(?[A-Za-z][\.\)]|\(?\d{1,3}[\.\)]|[\u2022\u00B7\u25CF\u25AA\u2013\u2014\*\-]|[\uF0B7\uF0A7\uF0D8\uF0FC])\s*'
$script:IG_LabelAny = '(?:^|\s)(?:\(?[IVXLCDMivxlcdm]{1,6}|\(?[A-Za-z]|\(?\d{1,3})[\.\)](?=\s|$)'
function IG-NoLabels([string]$s) { return ([regex]::Replace($s, $script:IG_LabelAny, ' ')) }

# one row = one changed paragraph. Result is FLAG, Expected or Info.
function IG-Row([string]$result, [string]$type, [string]$loc, [string]$seg, [string]$before, [string]$after, [string]$why) {
    return [pscustomobject]@{ Result = $result; Type = $type; Loc = $loc; Seg = $seg; Before = $before; After = $after; Why = $why }
}
function IG-Vis([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s.Replace("`t", '[tab]').Replace("`n", '[line break]')
    $s = [regex]::Replace($s, '^ +', { param($m) '[space]' * $m.Value.Length })
    $s = [regex]::Replace($s, ' +$', { param($m) '[space]' * $m.Value.Length })
    return $s
}
function IG-Describe([string]$old, [string]$new, [string]$del, [string]$ins, [int]$pos) {
    $where = 'in the middle'
    if ($pos -eq 0) { $where = 'at the start' } elseif (($pos + $del.Length) -ge $old.Length) { $where = 'at the end' }
    if ($del.Length -eq 0 -and $ins.Length -gt 0) { return ('Added "' + (IG-Vis $ins) + '" ' + $where + ' of the text') }
    if ($ins.Length -eq 0 -and $del.Length -gt 0) { return ('Removed "' + (IG-Vis $del) + '" ' + $where + ' of the text') }
    # a replacement: show the whole word(s) that changed so it reads naturally
    $ps = $pos; $er = $pos + $del.Length; $ea = $pos + $ins.Length
    while ($ps -gt 0 -and $old[$ps - 1] -notmatch '\s') { $ps-- }
    while ($er -lt $old.Length -and $old[$er] -notmatch '\s' -and $ea -lt $new.Length -and $new[$ea] -notmatch '\s') { $er++; $ea++ }
    return ('Changed "' + (IG-Vis $old.Substring($ps, $er - $ps)) + '" to "' + (IG-Vis $new.Substring($ps, $ea - $ps)) + '"')
}
function IG-Classify([string]$tool, $r, $a, [string]$loc) {
    $segs = IG-Segments $r.Text $a.Text
    $del = [string]$segs[0]; $ins = [string]$segs[1]; $pos = [int]$segs[2]
    $seg = IG-Describe $r.Text $a.Text $del $ins $pos
    if ($tool -eq 'SLID Macro') {
        $rx = '\s?\(\d{1,2} of \d{1,2}\)$'
        if (([regex]::Replace($r.Text, $rx, '')) -ceq ([regex]::Replace($a.Text, $rx, ''))) { return (IG-Row 'Expected' 'Slide counter' $loc $seg $r.Text $a.Text 'SLID adds or updates the (n of m) counter on repeated titles.') }
    }
    $noLab = [regex]::Replace($r.Text, $script:IG_LabelStart, '')
    if ($tool -eq 'Word Fixer' -and $noLab.Length -lt $r.Text.Length -and ($noLab -ceq $a.Text -or $noLab -ceq $a.Text.TrimStart())) {
        $seg = 'Removed the typed label "' + (IG-Show ($r.Text.Substring(0, $r.Text.Length - $noLab.Length))) + '" from the start'
        return (IG-Row 'Expected' 'Typed list label' $loc $seg $r.Text $a.Text 'A typed list label (I. / A. / 1. / bullet) was removed because the paragraph is now a real list.')
    }
    $kind = 'Text'
    $both = ($del + $ins)
    if ($both.Length -gt 0 -and $both.Trim().Length -eq 0) { $kind = 'Spacing' }
    elseif ($both.Length -gt 0 -and (($both -replace '[\s\p{P}\p{S}]', '').Length -eq 0)) { $kind = 'Punctuation' }
    if ($kind -eq 'Spacing' -or $kind -eq 'Punctuation') {
        return (IG-Row 'FLAG' $kind $loc $seg $r.Text $a.Text ('Only ' + $kind.ToLower() + ' changed. The tools must not add or remove spaces or punctuation (. , - ` etc.). Please check.'))
    }
    if ($tool -eq 'Link Fixer' -and ($r.Link -or $a.Link)) {
        return (IG-Row 'Expected' 'Hyperlink text' $loc $seg $r.Text $a.Text 'The paragraph contains a hyperlink; Hyperlink Tool may rewrite hyperlink text.')
    }
    return (IG-Row 'FLAG' 'Text changed' $loc $seg $r.Text $a.Text 'The words of this paragraph changed and no rule of this tool allows that. Please check.')
}
function IG-ComparePart([string]$tool, [string]$label, $origList, $fixList, $rows, $stats) {
    $om = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]'
    foreach ($o in $origList) { if (-not $om.ContainsKey($o.Text)) { $om[$o.Text] = New-Object 'System.Collections.Generic.List[object]' }; $om[$o.Text].Add($o) }
    $added = New-Object System.Collections.Generic.List[object]
    foreach ($f in $fixList) {
        if ($om.ContainsKey($f.Text) -and $om[$f.Text].Count -gt 0) { $om[$f.Text].RemoveAt(0); $stats.Same++ } else { $added.Add($f) }
    }
    $rem = New-Object System.Collections.Generic.List[object]
    foreach ($kv in $om.GetEnumerator()) { foreach ($o in $kv.Value) { $rem.Add($o) } }
    $remA = @($rem | Sort-Object Idx); $addA = @($added | Sort-Object Idx)
    if ($remA.Count -eq 0 -and $addA.Count -eq 0) { return }
    $usedR = New-Object 'bool[]' $remA.Count; $usedA = New-Object 'bool[]' $addA.Count
    function Loc([string]$lab, $idx) { if ($lab -match '^Slide') { return $lab } else { return ($lab + ' ' + $idx) } }
    # Word Tool may split a paragraph in two or join two; the words stay the same
    if ($tool -eq 'Word Fixer') {
        for ($i = 0; $i -lt $remA.Count; $i++) {
            if ($usedR[$i]) { continue }
            $rs = IG-Strip (IG-NoLabels $remA[$i].Text)
            for ($j = 0; $j -lt $addA.Count; $j++) {
                if ($usedA[$j]) { continue }
                $acc = ''; $k = $j; $ids = @()
                while ($k -lt $addA.Count -and $k -lt ($j + 3) -and -not $usedA[$k]) {
                    $acc += (IG-Strip (IG-NoLabels $addA[$k].Text)); $ids += $k
                    if ($acc -ceq $rs -and $ids.Count -ge 2) {
                        $usedR[$i] = $true; foreach ($q in $ids) { $usedA[$q] = $true }
                        $rows.Add((IG-Row 'Expected' 'Paragraph split' (Loc $label $remA[$i].Idx) ('1 paragraph became ' + $ids.Count) $remA[$i].Text (($ids | ForEach-Object { $addA[$_].Text }) -join ' || ') 'Word Tool split a paragraph that had run two list items together. The words are the same.'))
                        break
                    }
                    if (-not $rs.StartsWith($acc)) { break }
                    $k++
                }
                if ($usedR[$i]) { break }
            }
        }
        for ($j = 0; $j -lt $addA.Count; $j++) {
            if ($usedA[$j]) { continue }
            $as = IG-Strip (IG-NoLabels $addA[$j].Text)
            for ($i = 0; $i -lt $remA.Count; $i++) {
                if ($usedR[$i]) { continue }
                $acc = ''; $k = $i; $ids = @()
                while ($k -lt $remA.Count -and $k -lt ($i + 3) -and -not $usedR[$k]) {
                    $acc += (IG-Strip (IG-NoLabels $remA[$k].Text)); $ids += $k
                    if ($acc -ceq $as -and $ids.Count -ge 2) {
                        $usedA[$j] = $true; foreach ($q in $ids) { $usedR[$q] = $true }
                        $rows.Add((IG-Row 'Expected' 'Paragraphs joined' (Loc $label $remA[$i].Idx) ($ids.Count + ' paragraphs became 1') (($ids | ForEach-Object { $remA[$_].Text }) -join ' || ') $addA[$j].Text 'Word Tool joined paragraphs that belong together. The words are the same.'))
                        break
                    }
                    if (-not $as.StartsWith($acc)) { break }
                    $k++
                }
                if ($usedA[$j]) { break }
            }
        }
    }
    # pair what is left by similarity
    for ($i = 0; $i -lt $remA.Count; $i++) {
        if ($usedR[$i]) { continue }
        $best = -1; $bs = 0.0
        for ($j = 0; $j -lt $addA.Count; $j++) {
            if ($usedA[$j]) { continue }
            if ($addA.Count -gt 80 -and [Math]::Abs($addA[$j].Idx - $remA[$i].Idx) -gt 80) { continue }
            $s = IG-Sim $remA[$i].Text $addA[$j].Text
            if ($s -lt 0.9 -and ($addA[$j].Text.StartsWith($remA[$i].Text) -or $remA[$i].Text.StartsWith($addA[$j].Text) -or $addA[$j].Text.EndsWith($remA[$i].Text) -or $remA[$i].Text.EndsWith($addA[$j].Text))) { $s = 0.9 }
            if ($s -gt $bs) { $bs = $s; $best = $j }
        }
        if ($best -ge 0 -and $bs -ge 0.4) {
            $usedR[$i] = $true; $usedA[$best] = $true
            $rows.Add((IG-Classify $tool $remA[$i] $addA[$best] (Loc $label $remA[$i].Idx)))
        }
    }
    # what is left on the same slide / part: if as many were removed as added, show them side by side as a replacement
    $lr = @(); for ($i = 0; $i -lt $remA.Count; $i++) { if (-not $usedR[$i]) { $lr += $i } }
    $la = @(); for ($j = 0; $j -lt $addA.Count; $j++) { if (-not $usedA[$j]) { $la += $j } }
    if ($lr.Count -gt 0 -and $lr.Count -eq $la.Count -and $lr.Count -le 6) {
        for ($q = 0; $q -lt $lr.Count; $q++) {
            $usedR[$lr[$q]] = $true; $usedA[$la[$q]] = $true
            $rows.Add((IG-Row 'FLAG' 'Text changed' (Loc $label $remA[$lr[$q]].Idx) ('Replaced the whole text: "' + (IG-Show $remA[$lr[$q]].Text) + '" became "' + (IG-Show $addA[$la[$q]].Text) + '"') $remA[$lr[$q]].Text $addA[$la[$q]].Text 'The words changed and no rule of this tool allows that. Please check.'))
        }
    }
    for ($i = 0; $i -lt $remA.Count; $i++) {
        if ($usedR[$i]) { continue }
        $rows.Add((IG-Row 'FLAG' 'Text removed' (Loc $label $remA[$i].Idx) ('Removed a whole paragraph: "' + (IG-Show $remA[$i].Text) + '"') $remA[$i].Text '' 'This text is in the original but not in the fixed file. Please check.'))
    }
    for ($j = 0; $j -lt $addA.Count; $j++) {
        if ($usedA[$j]) { continue }
        $why = 'This text is in the fixed file but not in the original. Please check.'
        if ($tool -eq 'IMBK Macro') { $why = 'Text was added. IMBK Tool should only move figures and resize boxes, not add text. Please check.' }
        $rows.Add((IG-Row 'FLAG' 'Text added' (Loc $label $addA[$j].Idx) ('Added a whole paragraph: "' + (IG-Show $addA[$j].Text) + '"') '' $addA[$j].Text $why))
    }
}

# ---------------------------------------------------------------- the Excel file
function IG-Esc([string]$s) {
    if ($null -eq $s) { return '' }
    if ($s.Length -gt 3000) { $s = $s.Substring(0, 3000) + ' ...' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', '')
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}
function IG-ColName([int]$i) { $s = ''; $i++; while ($i -gt 0) { $r = ($i - 1) % 26; $s = [string][char](65 + $r) + $s; $i = [int][Math]::Floor(($i - 1) / 26) }; return $s }
# cells: a string, or @{ v = text; s = style } (styles: 1 header, 2 red, 3 green, 4 amber, 5 bold)
function IG-SheetXml($rows, $widths) {
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>')
    for ($i = 0; $i -lt $widths.Count; $i++) { [void]$sb.Append('<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + $widths[$i] + '" customWidth="1"/>') }
    [void]$sb.Append('</cols><sheetData>')
    $r = 0
    foreach ($row in $rows) {
        $r++; [void]$sb.Append('<row r="' + $r + '">')
        for ($c = 0; $c -lt $row.Count; $c++) {
            $cell = $row[$c]; $v = ''; $st = 0
            if ($cell -is [hashtable]) { $v = [string]$cell.v; $st = [int]$cell.s } else { $v = [string]$cell }
            if ($r -eq 1 -and $st -eq 0) { $st = 1 }
            [void]$sb.Append('<c r="' + (IG-ColName $c) + $r + '" t="inlineStr" s="' + $st + '"><is><t xml:space="preserve">' + (IG-Esc $v) + '</t></is></c>')
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData></worksheet>')
    return $sb.ToString()
}
function IG-WriteXlsx([string]$path, $sheets) {
    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF3B3B3B"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFAD4D4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDF0E0"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF1CC"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs></styleSheet>'
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = New-Object IO.FileStream($path, [IO.FileMode]::Create)
    $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $enc = New-Object Text.UTF8Encoding($false)
        $add = { param($n, $t) $e = $zip.CreateEntry($n); $w = New-Object IO.StreamWriter($e.Open(), $enc); $w.Write($t); $w.Close() }
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        $wbs = ''; $wbr = ''
        for ($i = 0; $i -lt $sheets.Count; $i++) {
            $n = $i + 1
            $ct += '<Override PartName="/xl/worksheets/sheet' + $n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            $wbs += '<sheet name="' + (IG-Esc $sheets[$i].Name) + '" sheetId="' + $n + '" r:id="rId' + $n + '"/>'
            $wbr += '<Relationship Id="rId' + $n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $n + '.xml"/>'
        }
        $ct += '</Types>'
        & $add '[Content_Types].xml' $ct
        & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        & $add 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbs + '</sheets></workbook>')
        $sn = $sheets.Count + 1
        & $add 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbr + '<Relationship Id="rId' + $sn + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        & $add 'xl/styles.xml' $styles
        for ($i = 0; $i -lt $sheets.Count; $i++) { & $add ('xl/worksheets/sheet' + ($i + 1) + '.xml') (IG-SheetXml $sheets[$i].Rows $sheets[$i].Widths) }
    } finally { $zip.Dispose(); $fs.Dispose() }
}

# ---------------------------------------------------------------- one before / after pair
$script:IG_Rules = @{
    'Word Fixer'  = 'Allowed: a typed list label (I. / A. / 1. / bullet) removed because the paragraph became a real list; a paragraph split or joined with the same words. Everything else is flagged.'
    'SLID Macro'  = 'Allowed: the (n of m) counter added to repeated slide titles. Slides may be duplicated. Shape names and reading order are not text. Everything else is flagged.'
    'IMBK Macro'  = 'Allowed: figures moved and boxes resized (no text change). Any added, removed or changed text is flagged.'
    'Alt Text Tool' = 'Allowed: alt text on pictures (it is not slide text). Any slide text change is flagged.'
    'Link Fixer'  = 'Allowed: hyperlink text rewritten in paragraphs that contain a hyperlink. Spacing or punctuation-only changes are always flagged.'
    'Issue Fixer' = 'Allowed: language tags, reading order, shape names, colours and title placeholders. No text may change; any text change is flagged.'
}
function Test-Integrity([string]$tool, [string]$orig, [string]$fixed, [string]$reportPath) {
    $ext = [IO.Path]::GetExtension($fixed).ToLowerInvariant()
    if ($ext -eq '.pptx') { $origSet = IG-ReadPptx $orig; $fixSet = IG-ReadPptx $fixed } elseif ($ext -eq '.docx') { $origSet = IG-ReadDocx $orig; $fixSet = IG-ReadDocx $fixed } else { return $null }
    $rows = New-Object System.Collections.Generic.List[object]
    $stats = [pscustomobject]@{ Same = 0; OrigParas = 0; FixParas = 0; NewSlides = 0 }
    $fixMap = @{}; foreach ($fx in $fixSet) { $fixMap[$fx.Key] = $fx; $stats.FixParas += $fx.Paras.Count }
    $origMap = @{}; foreach ($og in $origSet) { $origMap[$og.Key] = $og; $stats.OrigParas += $og.Paras.Count }
    foreach ($og in $origSet) {
        if (-not $fixMap.ContainsKey($og.Key)) {
            if ($og.Paras.Count -gt 0) { $rows.Add((IG-Row 'FLAG' 'Part removed' $og.Label ($og.Label + ' is not in the fixed file') ((@($og.Paras | ForEach-Object { $_.Text }) -join ' || ')) '' 'This slide or part is in the original but not in the fixed file.')) }
            continue
        }
        IG-ComparePart $tool $fixMap[$og.Key].Label $og.Paras $fixMap[$og.Key].Paras $rows $stats
    }
    foreach ($fx in $fixSet) {
        if (-not $origMap.ContainsKey($fx.Key)) {
            if ($fx.Paras.Count -gt 0) { $stats.NewSlides++; $rows.Add((IG-Row 'Info' 'New slide or part' $fx.Label ('new: ' + $fx.Paras.Count + ' paragraph(s)') '' ((@($fx.Paras | ForEach-Object { $_.Text }) -join ' || ')) 'This slide or part is not in the original (for example a duplicated slide). Its text was not compared.')) }
        }
    }
    $flag = @($rows | Where-Object { $_.Result -eq 'FLAG' }).Count
    $exp = @($rows | Where-Object { $_.Result -eq 'Expected' }).Count
    $info = @($rows | Where-Object { $_.Result -eq 'Info' }).Count
    $verdict = if ($flag -eq 0) { 'PASS - all text is preserved (' + $exp + ' change(s) allowed by the tool rules)' } else { 'REVIEW - ' + $flag + ' change(s) need a check (see the Changes sheet, red rows)' }
    $ci = [Globalization.CultureInfo]::InvariantCulture
    $sum = New-Object System.Collections.Generic.List[object]
    $sum.Add(@('Integrity report', ''))
    $sum.Add(@(@{ v = 'Result'; s = 5 }, @{ v = $verdict; s = $(if ($flag -eq 0) { 3 } else { 2 }) }))
    $sum.Add(@('Tool', $tool))
    $sum.Add(@('Original file', $orig))
    $sum.Add(@('Fixed file', $fixed))
    $sum.Add(@('Checked on', (Get-Date).ToString('yyyy-MM-dd HH:mm', $ci)))
    $sum.Add(@('Paragraphs in the original', [string]$stats.OrigParas))
    $sum.Add(@('Paragraphs in the fixed file', [string]$stats.FixParas))
    $sum.Add(@('Paragraphs with identical text', [string]$stats.Same))
    $sum.Add(@('Changes flagged for a person to check', [string]$flag))
    $sum.Add(@('Changes allowed by the tool rules', [string]$exp))
    $sum.Add(@('New slides / parts (not compared)', [string]$info))
    $byType = @{}; foreach ($rw in $rows) { $k2 = $rw.Result + ' - ' + $rw.Type; if ($byType.ContainsKey($k2)) { $byType[$k2]++ } else { $byType[$k2] = 1 } }
    foreach ($k2 in ($byType.Keys | Sort-Object)) { $sum.Add(@(('   ' + $k2), [string]$byType[$k2])) }
    $sum.Add(@('Rules for this tool', $(if ($script:IG_Rules.ContainsKey($tool)) { $script:IG_Rules[$tool] } else { '' })))
    $sum.Add(@('How it works', 'Every paragraph of the original and the fixed file is compared exactly, including spaces and punctuation. Paragraphs are matched by their text, so moving a shape or changing the reading order is not a text change. Nothing in either file was changed by this check.'))
    $order = @{ 'FLAG' = 0; 'Expected' = 1; 'Info' = 2 }
    $sorted = @($rows | Sort-Object @{ Expression = { $order[$_.Result] } }, @{ Expression = { $mm = [regex]::Match($_.Loc, '\d+'); if ($mm.Success) { [int]$mm.Value } else { 0 } } })
    $ch = New-Object System.Collections.Generic.List[object]
    $ch.Add(@('#', 'Where', 'Result', 'Type', 'What changed', 'Before (original)', 'After (fixed)', 'Why'))
    $n = 0
    foreach ($r in $sorted) {
        $n++; $st = if ($r.Result -eq 'FLAG') { 2 } elseif ($r.Result -eq 'Expected') { 3 } else { 4 }
        $ch.Add(@(@{ v = [string]$n; s = $st }, @{ v = $r.Loc; s = $st }, @{ v = $r.Result; s = $st }, @{ v = $r.Type; s = $st }, @{ v = $r.Seg; s = $st }, @{ v = $r.Before; s = $st }, @{ v = $r.After; s = $st }, @{ v = $r.Why; s = $st }))
    }
    if ($n -eq 0) { $ch.Add(@('', '', 'None', '', 'No text differences were found.', '', '', '')) }
    $sheets = @(
        @{ Name = 'Summary'; Rows = $sum; Widths = @(38, 110) },
        @{ Name = 'Changes'; Rows = $ch; Widths = @(5, 16, 10, 18, 40, 55, 55, 55) }
    )
    IG-WriteXlsx $reportPath $sheets
    return [pscustomobject]@{ Flagged = $flag; Expected = $exp; Info = $info; Verdict = $verdict; Report = $reportPath }
}

# ---------------------------------------------------------------- after a tool has run: check every file it saved
function Invoke-IntegrityRun([string]$toolFile, [datetime]$since) {
    $names = @{ 'Word_Fixer.ps1' = 'Word Fixer'; 'SLID_Macro.ps1' = 'SLID Macro'; 'IMBK_Macro.ps1' = 'IMBK Macro'; 'AltText_Tool.ps1' = 'Alt Text Tool'; 'Link_Fixer.ps1' = 'Link Fixer'; 'Issue_Fixer.ps1' = 'Issue Fixer' }
    if (-not $names.ContainsKey($toolFile)) { return }
    $tool = $names[$toolFile]
    $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }
    $log = Join-Path (Join-Path (Join-Path $base 'S4C_Tools') 'WorkLog') 'worklog.tsv'
    if (-not (Test-Path -LiteralPath $log)) { return }
    $ci = [Globalization.CultureInfo]::InvariantCulture
    $from = $since.AddMinutes(-1)
    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($l in [IO.File]::ReadAllLines($log, [Text.Encoding]::UTF8)) {
        $c = $l.Split("`t"); if ($c.Count -lt 6 -or $c[3] -ne $tool) { continue }
        $t = [datetime]::MinValue
        if (-not [datetime]::TryParseExact(($c[0] + ' ' + $c[1]), 'yyyy-MM-dd HH:mm', $ci, [Globalization.DateTimeStyles]::None, [ref]$t)) { continue }
        if ($t -lt $from) { continue }
        $pairs.Add(@($c[4], $c[5]))
    }
    $done = @{}
    foreach ($pr in $pairs) {
        $in = $pr[0]; $out = $pr[1]
        if ($done.ContainsKey($out)) { continue }; $done[$out] = $true
        try {
            $orig = $in
            if ($in -eq $out) {
                # IMBK works on the file itself and keeps a backup first
                $dir = Split-Path $out; $bn = [IO.Path]::GetFileNameWithoutExtension($out); $ex = [IO.Path]::GetExtension($out)
                $bk = @(Get-ChildItem -LiteralPath $dir -Filter ($bn + '_BEFORE_IMBKfix*' + $ex) -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
                if ($bk.Count -eq 0) { Write-Host ('Integrity check skipped for ' + (Split-Path $out -Leaf) + ': no backup of the original was found.') -ForegroundColor Yellow; continue }
                $orig = $bk[0].FullName
            }
            if (-not (Test-Path -LiteralPath $orig) -or -not (Test-Path -LiteralPath $out)) { continue }
            $rp = Join-Path (Split-Path $out) ([IO.Path]::GetFileNameWithoutExtension($out) + ' Integrity Report.xlsx')
            Write-Host ('Checking the text of ' + (Split-Path $out -Leaf) + ' against the original...') -ForegroundColor Cyan
            $r = Test-Integrity $tool $orig $out $rp
            if ($null -eq $r) { continue }
            if ($r.Flagged -eq 0) { Write-Host ('Integrity: ' + $r.Verdict) -ForegroundColor Green } else { Write-Host ('Integrity: ' + $r.Verdict) -ForegroundColor Yellow }
            Write-Host ('Integrity flagged: ' + $r.Flagged)
            Write-Host ('Integrity report: ' + $r.Report)
        } catch { Write-Host ('Integrity check could not finish for ' + (Split-Path $out -Leaf) + ': ' + $_.Exception.Message) -ForegroundColor Yellow }
    }
}
