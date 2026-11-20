# WORD FIXER SCRIPT
# Word Fixer (S4Carlisle) - fixes the structure of a Word (.docx) file for accessibility, no Word needed.
#
#   1. Title lines at the top of the chapter (chapter number / chapter title) get the Title style.
#   2. Custom heading styles (a style with an outline level, e.g. "H1") become the real Word headings
#      (Heading 1, Heading 2 ...), keeping the same look.
#   3. Lists that start at list level 2 or deeper are moved up so they start at level 1, and every
#      level below moves up with them (the numbers and bullets you see stay the same).
#   4. Numbers typed by hand in front of list items ("I. Introduction") become real list numbering.
#   5. List items that had their numbering switched off by accident, or that sit in a separate copy
#      of the same list, are put back into the main list so the numbering runs on correctly.
#   6. The file's Title (File > Info) is set from the title lines.
#
# The original file is never changed: the result is saved as "<name> Fixed.docx" with a report next to it.
$WF_VERSION = '1.3'   # shown at the start of the window; the updater compares this number with GitHub

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
$NS_DC = 'http://purl.org/dc/elements/1.1/'

# ----------------------------------------------------------------- helpers
function New-XmlDoc([string]$text) { $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.LoadXml($text); return $d }
function New-Ns([Xml.XmlDocument]$doc) {
    $m = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    [void]$m.AddNamespace('w', $NS_W); [void]$m.AddNamespace('cp', $NS_CP); [void]$m.AddNamespace('dc', $NS_DC); [void]$m.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    return ,$m
}
function Read-Entry($zip, [string]$name) {
    $e = $zip.GetEntry($name); if ($null -eq $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}
function Write-Entry($zip, [string]$name, [string]$text) {
    $e = $zip.GetEntry($name); if ($e) { $e.Delete() }
    $ne = $zip.CreateEntry($name)
    $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
    try { $sw.Write($text) } finally { $sw.Dispose() }
}
function SSN($node, [string]$xp, $ns) { if ($null -eq $node) { return $null }; return $node.SelectSingleNode($xp, $ns) }
function WAttr($el, [string]$name) { if ($null -eq $el) { return $null }; $a = $el.GetAttributeNode($name, $NS_W); if ($a) { return $a.Value }; return $null }
function Set-WAttr($el, [string]$name, [string]$val) { [void]$el.SetAttribute($name, $NS_W, $val) }
function New-W($doc, [string]$name) { return $doc.CreateElement('w', $name, $NS_W) }
function Get-ParaText($p, $nsm) { return (($p.SelectNodes('.//w:t', $nsm) | ForEach-Object { $_.InnerText }) -join '') }

# ----------------------------------------------------------------- styles
function Get-StyleInfo($stylesDoc, $nsm) {
    $info = @{}
    foreach ($s in $stylesDoc.SelectNodes('/w:styles/w:style[@w:type="paragraph"]', $nsm)) {
        $id = WAttr $s 'styleId'
        $nm = WAttr (SSN $s 'w:name' $nsm) 'val'
        $info[$id] = @{
            node = $s; name = [string]$nm; basedOn = (WAttr (SSN $s 'w:basedOn' $nsm) 'val')
            numPr = $s.SelectSingleNode('w:pPr/w:numPr', $nsm); outline = (WAttr (SSN $s 'w:pPr/w:outlineLvl' $nsm) 'val')
        }
    }
    return $info
}
function Get-StyleChainValue($info, [string]$sid, [string]$what) {
    $seen = @{}
    while ($sid -and $info.ContainsKey($sid) -and -not $seen.ContainsKey($sid)) {
        $seen[$sid] = $true
        $v = $info[$sid][$what]
        if ($null -ne $v -and $v -ne '') { return $v }
        $sid = $info[$sid].basedOn
    }
    return $null
}
function Get-DefaultParaStyle($stylesDoc, $nsm) {
    $d = $stylesDoc.SelectSingleNode('/w:styles/w:style[@w:type="paragraph" and @w:default="1"]', $nsm)
    if ($d) { return (WAttr $d 'styleId') }
    return 'Normal'
}
function Test-BuiltinHeading([string]$name) { return ($name -match '^(?i)heading [1-9]$') }

# Returns the styleId of the built-in style with this name (e.g. "heading 1", "Title"), creating it if needed.
# When the built-in style is not used anywhere in the document yet, it takes over the look of $lookFrom
# (the custom style it replaces), so the page looks the same as before.
function Use-BuiltinStyle($stylesDoc, $nsm, $info, [string]$builtinName, [string]$builtinId, [string]$lookFrom, $usedIds, [int]$outline = -1) {
    $existing = $null
    foreach ($k in $info.Keys) { if ($info[$k].name -ieq $builtinName) { $existing = $k; break } }
    $target = $existing
    $takeLook = $false
    if (-not $existing) {
        $target = $builtinId
        $s = New-W $stylesDoc 'style'; Set-WAttr $s 'type' 'paragraph'; Set-WAttr $s 'styleId' $target
        $n = New-W $stylesDoc 'name'; Set-WAttr $n 'val' $builtinName; [void]$s.AppendChild($n)
        $q = New-W $stylesDoc 'qFormat'; [void]$s.AppendChild($q)
        [void]$stylesDoc.DocumentElement.AppendChild($s)
        $info[$target] = @{ node = $s; name = $builtinName; basedOn = $null; numPr = $null; outline = $null }
        $takeLook = $true
    } elseif (-not $usedIds.ContainsKey($existing)) { $takeLook = $true }
    if ($takeLook -and $lookFrom -and $info.ContainsKey($lookFrom)) {
        $s = $info[$target].node; $src = $info[$lookFrom].node
        foreach ($tag in @('basedOn', 'pPr', 'rPr')) {
            $old = $s.SelectSingleNode('w:' + $tag, $nsm); if ($old) { [void]$s.RemoveChild($old) }
            $new = $src.SelectSingleNode('w:' + $tag, $nsm)
            if ($new) { [void]$s.AppendChild($new.CloneNode($true)) }
        }
        $pPr = $s.SelectSingleNode('w:pPr', $nsm)
        if ($pPr) { foreach ($x in @($pPr.SelectNodes('w:numPr|w:outlineLvl', $nsm))) { [void]$pPr.RemoveChild($x) } }
        $info[$target].basedOn = WAttr (SSN $s 'w:basedOn' $nsm) 'val'
        $info[$target].numPr = $null
        $script:LookOf[$target] = $lookFrom
    }
    if ($outline -ge 0) {
        $s = $info[$target].node
        $pPr = $s.SelectSingleNode('w:pPr', $nsm); if (-not $pPr) { $pPr = New-W $stylesDoc 'pPr'; [void]$s.AppendChild($pPr) }
        $ol = $pPr.SelectSingleNode('w:outlineLvl', $nsm); if (-not $ol) { $ol = New-W $stylesDoc 'outlineLvl'; [void]$pPr.AppendChild($ol) }
        Set-WAttr $ol 'val' ([string]$outline)
        $info[$target].outline = [string]$outline
    }
    # keep the w:style children in the order Word expects
    $s = $info[$target].node
    $order = @('name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr')
    $kids = @($s.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
    foreach ($k in $kids) { [void]$s.RemoveChild($k) }
    foreach ($k in ($kids | Sort-Object { $i = [array]::IndexOf($order, $_.LocalName); if ($i -lt 0) { 99 } else { $i } })) { [void]$s.AppendChild($k) }
    return $target
}
function Set-ParaStyle($p, $nsm, [string]$sid) {
    $doc = $p.OwnerDocument
    $pPr = $p.SelectSingleNode('w:pPr', $nsm)
    if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.PrependChild($pPr) }
    $ps = $pPr.SelectSingleNode('w:pStyle', $nsm)
    if (-not $ps) { $ps = New-W $doc 'pStyle'; [void]$pPr.PrependChild($ps) }
    Set-WAttr $ps 'val' $sid
}

$PPR_ORDER = @('pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange')
$RPR_ORDER = @('rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange')
# puts the children of a w:pPr / w:rPr back in the order Word requires (unknown ones keep their place at the end)
function Sort-Kids($el, $order) {
    if ($null -eq $el) { return }
    $kids = @($el.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
    if ($kids.Count -lt 2) { return }
    $sorted = @($kids | Sort-Object { $i = [array]::IndexOf($order, $_.LocalName); if ($i -lt 0) { 999 } else { $i } })
    foreach ($k in $kids) { [void]$el.RemoveChild($k) }
    foreach ($k in $sorted) { [void]$el.AppendChild($k) }
}
# formatting a style really gives (its own, plus what it inherits): localName -> node, the nearest style wins
function Get-ChainProps($info, $snsm, [string]$sid, [string]$tag) {
    $out = @{}; $seen = @{}
    while ($sid -and $info.ContainsKey($sid) -and -not $seen.ContainsKey($sid)) {
        $seen[$sid] = $true
        $pr = $info[$sid].node.SelectSingleNode('w:' + $tag, $snsm)
        if ($pr) { foreach ($c in $pr.ChildNodes) { if ($c.NodeType -eq 'Element' -and -not $out.ContainsKey($c.LocalName)) { $out[$c.LocalName] = $c } } }
        $sid = $info[$sid].basedOn
    }
    return $out
}
$script:TypedRemoved = @{}  # paragraph -> typed number text that was replaced by real numbering
$script:Adopted = @{}  # parent list level -> the sublist level whose look it took
$script:LookOf = @{}   # built-in style -> the custom style whose look it took over
# Gives a paragraph a new style but keeps exactly how it looked: what the old style gave it is written onto the
# paragraph, and what the new style would add (bold, numbering ...) is switched off.
function Set-StyleKeepLook($p, $nsm, $info, $snsm, $sty, [string]$srcSid, [string]$tgtSid, [string]$defSid) {
    $doc = $p.OwnerDocument
    if ($script:LookOf.ContainsKey($tgtSid) -and $script:LookOf[$tgtSid] -eq $srcSid) { Set-ParaStyle $p $nsm $tgtSid; return }
    $srcP = Get-ChainProps $info $snsm $srcSid 'pPr'; $tgtP = Get-ChainProps $info $snsm $tgtSid 'pPr'; $defP = Get-ChainProps $info $snsm $defSid 'pPr'
    $srcR = Get-ChainProps $info $snsm $srcSid 'rPr'; $tgtR = Get-ChainProps $info $snsm $tgtSid 'rPr'; $defR = Get-ChainProps $info $snsm $defSid 'rPr'
    $ddR = $sty.SelectSingleNode('/w:styles/w:docDefaults/w:rPrDefault/w:rPr', $snsm)
    if ($ddR) { foreach ($c in $ddR.ChildNodes) { if ($c.NodeType -eq 'Element' -and -not $defR.ContainsKey($c.LocalName)) { $defR[$c.LocalName] = $c } } }
    $ddP = $sty.SelectSingleNode('/w:styles/w:docDefaults/w:pPrDefault/w:pPr', $snsm)
    if ($ddP) { foreach ($c in $ddP.ChildNodes) { if ($c.NodeType -eq 'Element' -and -not $defP.ContainsKey($c.LocalName)) { $defP[$c.LocalName] = $c } } }
    Set-ParaStyle $p $nsm $tgtSid
    $pPr = $p.SelectSingleNode('w:pPr', $nsm)
    foreach ($k in $srcP.Keys) {
        if ($k -in @('numPr', 'outlineLvl', 'pStyle', 'rPr')) { continue }
        if (-not $pPr.SelectSingleNode('w:' + $k, $nsm)) { [void]$pPr.AppendChild($doc.ImportNode($srcP[$k], $true)) }
    }
    foreach ($k in $tgtP.Keys) {
        if ($k -in @('numPr', 'outlineLvl', 'pStyle', 'rPr') -or $srcP.ContainsKey($k)) { continue }
        if (-not $pPr.SelectSingleNode('w:' + $k, $nsm) -and $defP.ContainsKey($k)) { [void]$pPr.AppendChild($doc.ImportNode($defP[$k], $true)) }
    }
    if ($tgtP.ContainsKey('numPr') -and -not $srcP.ContainsKey('numPr') -and -not $pPr.SelectSingleNode('w:numPr', $nsm)) {
        $np = New-W $doc 'numPr'; $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' '0'; $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' '0'
        [void]$np.AppendChild($il); [void]$np.AppendChild($ni); [void]$pPr.AppendChild($np)
    }
    Sort-Kids $pPr $PPR_ORDER
    foreach ($r in $p.SelectNodes('.//w:r', $nsm)) {
        $rPr = $r.SelectSingleNode('w:rPr', $nsm); if (-not $rPr) { $rPr = New-W $doc 'rPr'; [void]$r.PrependChild($rPr) }
        $csNames = @{}
        $rs = WAttr (SSN $rPr 'w:rStyle' $nsm) 'val'
        if ($rs) { $cs = $sty.SelectSingleNode('/w:styles/w:style[@w:styleId="' + $rs + '"]/w:rPr', $snsm); if ($cs) { foreach ($c in $cs.ChildNodes) { if ($c.NodeType -eq 'Element') { $csNames[$c.LocalName] = $true } } } }
        foreach ($k in $srcR.Keys) {
            if ($csNames.ContainsKey($k) -or $rPr.SelectSingleNode('w:' + $k, $nsm)) { continue }
            [void]$rPr.AppendChild($doc.ImportNode($srcR[$k], $true))
        }
        foreach ($k in $tgtR.Keys) {
            if ($srcR.ContainsKey($k) -or $csNames.ContainsKey($k) -or $rPr.SelectSingleNode('w:' + $k, $nsm)) { continue }
            if ($k -in @('b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'vanish', 'u')) {
                $off = New-W $doc $k; Set-WAttr $off 'val' $(if ($k -eq 'u') { 'none' } else { '0' }); [void]$rPr.AppendChild($off)
            } elseif ($defR.ContainsKey($k)) { [void]$rPr.AppendChild($doc.ImportNode($defR[$k], $true)) }
            elseif ($k -eq 'color') { $cl = New-W $doc 'color'; Set-WAttr $cl 'val' 'auto'; [void]$rPr.AppendChild($cl) }
        }
        Sort-Kids $rPr $RPR_ORDER
    }
}

# ----------------------------------------------------------------- numbering
$ROMAN = @{ 'I' = 1; 'V' = 5; 'X' = 10; 'L' = 50; 'C' = 100 }
function ConvertFrom-Roman([string]$r) {
    $t = 0; $prev = 0
    $chars = $r.ToUpperInvariant().ToCharArray()
    for ($i = $chars.Count - 1; $i -ge 0; $i--) { $v = $ROMAN[[string]$chars[$i]]; if ($v -lt $prev) { $t -= $v } else { $t += $v; $prev = $v } }
    return $t
}
# typed number at the start of a list item: returns @{ kind; value; length } or $null
function Get-TypedNumber([string]$text) {
    $m = [regex]::Match($text, '^\s*([IVXLC]+)[\.\)]\s+')
    if ($m.Success) { return @{ kind = 'upperRoman'; value = (ConvertFrom-Roman $m.Groups[1].Value); length = $m.Length } }
    $m = [regex]::Match($text, '^\s*([ivxlc]+)[\.\)]\s+')
    if ($m.Success) { return @{ kind = 'lowerRoman'; value = (ConvertFrom-Roman $m.Groups[1].Value); length = $m.Length } }
    $m = [regex]::Match($text, '^\s*(\d{1,3})[\.\)]\s+')
    if ($m.Success) { return @{ kind = 'decimal'; value = [int]$m.Groups[1].Value; length = $m.Length } }
    $m = [regex]::Match($text, '^\s*([A-Z])[\.\)]\s+')
    if ($m.Success) { return @{ kind = 'upperLetter'; value = ([int][char]$m.Groups[1].Value - 64); length = $m.Length } }
    $m = [regex]::Match($text, '^\s*([a-z])[\.\)]\s+')
    if ($m.Success) { return @{ kind = 'lowerLetter'; value = ([int][char]$m.Groups[1].Value - 96); length = $m.Length } }
    return $null
}
function Remove-LeadingChars($p, $nsm, [int]$count) {
    foreach ($t in $p.SelectNodes('.//w:t', $nsm)) {
        if ($count -le 0) { break }
        $s = $t.InnerText
        if ($s.Length -le $count) { $count -= $s.Length; $t.InnerText = '' }
        else { $t.InnerText = $s.Substring($count); $count = 0 }
    }
    # strip any tab left right after the typed number
    $first = $p.SelectSingleNode('.//w:r[w:t[string-length(.)>0] or w:tab]', $nsm)
    if ($first -and $first.SelectSingleNode('w:tab', $nsm) -and -not $first.SelectSingleNode('w:t[string-length(.)>0]', $nsm)) { [void]$first.ParentNode.RemoveChild($first) }
}
function Format-Number([string]$kind, [int]$n) {
    switch ($kind) {
        'upperRoman' { return (ConvertTo-Roman $n) }
        'lowerRoman' { return (ConvertTo-Roman $n).ToLowerInvariant() }
        'upperLetter' { return [string][char](64 + $n) }
        'lowerLetter' { return [string][char](96 + $n) }
        default { return [string]$n }
    }
}
function ConvertTo-Roman([int]$n) {
    $vals = @(100, 90, 50, 40, 10, 9, 5, 4, 1); $syms = @('C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I'); $s = ''
    for ($i = 0; $i -lt $vals.Count; $i++) { while ($n -ge $vals[$i]) { $s += $syms[$i]; $n -= $vals[$i] } }
    return $s
}

# ----------------------------------------------------------------- safety check: what the reader sees
# Works out, for every paragraph, the list number / bullet Word shows in front of it (the same way for the file
# before and after the fix), so the two can be compared.
function Get-ListLabels($paras, $nsm, $sty, $snsm, $num, $nnsm, [string]$defStyle) {
    $labels = @{}
    if (-not $num) { foreach ($p in $paras) { $labels[$p] = '' }; return $labels }
    $sInfo = Get-StyleInfo $sty $snsm
    $numToAbs = @{}; $starts = @{}
    foreach ($n in $num.SelectNodes('/w:numbering/w:num', $nnsm)) {
        $id = WAttr $n 'numId'; $numToAbs[$id] = WAttr (SSN $n 'w:abstractNumId' $nnsm) 'val'
        $starts[$id] = @{}
        foreach ($ov in $n.SelectNodes('w:lvlOverride', $nnsm)) { $so = SSN $ov 'w:startOverride' $nnsm; if ($so) { $starts[$id][[int](WAttr $ov 'ilvl')] = [int](WAttr $so 'val') } }
    }
    $absNode = @{}; foreach ($an in $num.SelectNodes('/w:numbering/w:abstractNum', $nnsm)) { $absNode[(WAttr $an 'abstractNumId')] = $an }
    $counters = @{}; $seenNum = @{}
    foreach ($p in $paras) {
        $labels[$p] = ''
        $sid = WAttr (SSN $p 'w:pPr/w:pStyle' $nsm) 'val'; if (-not $sid) { $sid = $defStyle }
        $pn = SSN $p 'w:pPr/w:numPr' $nsm
        $sn = Get-StyleChainValue $sInfo $sid 'numPr'
        $nid = WAttr (SSN $pn 'w:numId' $nsm) 'val'; $il = WAttr (SSN $pn 'w:ilvl' $nsm) 'val'
        if (-not $nid -and $sn) { $nid = WAttr (SSN $sn 'w:numId' $snsm) 'val' }
        if (-not $il -and $sn -and -not (WAttr (SSN $pn 'w:numId' $nsm) 'val')) { $il = WAttr (SSN $sn 'w:ilvl' $snsm) 'val' }
        if (-not $il) { $il = '0' }
        if (-not $nid -or $nid -eq '0' -or -not $numToAbs.ContainsKey($nid)) { continue }
        $a = $numToAbs[$nid]; if (-not $absNode.ContainsKey($a)) { continue }
        $an = $absNode[$a]; $L = [int]$il
        if (-not $counters.ContainsKey($a)) { $counters[$a] = @{} }
        $ctr = $counters[$a]
        if (-not $seenNum.ContainsKey($nid)) { $seenNum[$nid] = $true; foreach ($k in $starts[$nid].Keys) { $ctr[$k] = $starts[$nid][$k] - 1 } }
        $lv = SSN $an ('w:lvl[@w:ilvl="' + $L + '"]') $nnsm
        if (-not $lv) { continue }
        $st = WAttr (SSN $lv 'w:start' $nnsm) 'val'; if (-not $st) { $st = '0' }
        if (-not $ctr.ContainsKey($L)) { $ctr[$L] = [int]$st - 1 }
        $ctr[$L] = $ctr[$L] + 1
        foreach ($k in @($ctr.Keys)) { if ($k -gt $L) { [void]$ctr.Remove($k) } }
        $txt = [string](WAttr (SSN $lv 'w:lvlText' $nnsm) 'val')
        $fmt = WAttr (SSN $lv 'w:numFmt' $nnsm) 'val'
        if ($fmt -eq 'bullet') { $labels[$p] = $txt; continue }
        $labels[$p] = [regex]::Replace($txt, '%(\d)', {
            param($mm)
            $k = [int]$mm.Groups[1].Value - 1
            $l2 = SSN $an ('w:lvl[@w:ilvl="' + $k + '"]') $nnsm
            $f2 = WAttr (SSN $l2 'w:numFmt' $nnsm) 'val'
            $v = if ($ctr.ContainsKey($k)) { $ctr[$k] } else { $s2 = WAttr (SSN $l2 'w:start' $nnsm) 'val'; if ($s2) { [int]$s2 } else { 1 } }
            switch ($f2) { 'none' { '' } 'upperRoman' { ConvertTo-Roman $v } 'lowerRoman' { (ConvertTo-Roman $v).ToLowerInvariant() } 'upperLetter' { [string][char](64 + (($v - 1) % 26) + 1) } 'lowerLetter' { [string][char](96 + (($v - 1) % 26) + 1) } default { [string]$v } }
        })
    }
    return $labels
}

# ----------------------------------------------------------------- the fix
function Repair-Docx([string]$inPath, [string]$outPath) {
    Copy-Item -LiteralPath $inPath -Destination $outPath -Force
    $report = New-Object System.Collections.Generic.List[string]
    $zip = [IO.Compression.ZipFile]::Open($outPath, 'Update')
    try {
        $doc = New-XmlDoc (Read-Entry $zip 'word/document.xml'); $nsm = New-Ns $doc
        $sty = New-XmlDoc (Read-Entry $zip 'word/styles.xml'); $snsm = New-Ns $sty
        $numText = Read-Entry $zip 'word/numbering.xml'
        $num = $null; $nnsm = $null
        if ($numText) { $num = New-XmlDoc $numText; $nnsm = New-Ns $num }
        $info = Get-StyleInfo $sty $snsm
        $defStyle = Get-DefaultParaStyle $sty $snsm
        $body = $doc.SelectSingleNode('/w:document/w:body', $nsm)
        $paras = @($body.SelectNodes('.//w:p', $nsm))
        $script:TypedRemoved = @{}; $script:Adopted = @{}; $script:LookOf = @{}; $script:Absorbed = @{}; $script:NewNumbering = $false; $script:MarkerAdded = @{}
        # what the reader sees before the fix: text and list number of every paragraph
        $textBefore = @{}; foreach ($p in $paras) { $textBefore[$p] = Get-ParaText $p $nsm }
        $labelsBefore = Get-ListLabels $paras $nsm $sty $snsm $num $nnsm $defStyle
        $pStyleOf = { param($p) $v = WAttr (SSN $p 'w:pPr/w:pStyle' $nsm) 'val'; if ($v) { $v } else { $defStyle } }
        $usedIds = @{}; foreach ($p in $paras) { $usedIds[(& $pStyleOf $p)] = $true }

        # ---- 1. custom heading styles -> real headings
        $headingMap = @{}
        foreach ($sid in @($usedIds.Keys)) {
            if (-not $info.ContainsKey($sid)) { continue }
            if (Test-BuiltinHeading $info[$sid].name) { continue }
            if ($info[$sid].name -ieq 'Title') { continue }
            if ($info[$sid].name -match '(?i)^(toc|index|caption|header|footer|footnote|endnote)') { continue }
            if (Get-StyleChainValue $info $sid 'numPr') { continue }   # numbered outline styles are lists, not headings
            $ol = Get-StyleChainValue $info $sid 'outline'
            if ($null -ne $ol -and [int]$ol -le 8) { $lvl = [int]$ol + 1 }
            else {
                # a heading style that was never marked as one: its name says so (H1, OBJ_H2, Head2, Heading A ...)
                $mh = [regex]::Match($info[$sid].name, '(?i)(?:^|[_\s\-])(?:h|head|heading)\s?([1-6])(?=$|[_\s\-])')
                if (-not $mh.Success) { continue }
                # only when its paragraphs really look like headings: short, outside tables, not list items
                $ps = @($paras | Where-Object { (& $pStyleOf $_) -eq $sid -and (Get-ParaText $_ $nsm).Trim() })
                if ($ps.Count -eq 0) { continue }
                $long = @($ps | Where-Object { (Get-ParaText $_ $nsm).Trim().Length -gt 150 -or $_.SelectSingleNode('ancestor::w:tbl', $nsm) -or ((WAttr (SSN $_ 'w:pPr/w:numPr/w:numId' $nsm) 'val') -notin @($null, '0')) })
                if ($long.Count -gt 0) { continue }
                $lvl = [int]$mh.Groups[1].Value
                $ol = [string]($lvl - 1)
            }
            $target = Use-BuiltinStyle $sty $snsm $info ('heading ' + $lvl) ('Heading' + $lvl) $sid $usedIds $ol
            $headingMap[$sid] = $target
        }
        # lines in a heading style that are really a typed list ("1. D & I", "2. H" ...) stay out of the headings
        $typedRun = @{}; $runL = New-Object System.Collections.Generic.List[object]; $prevV = 0
        foreach ($p in $paras) {
            $t = (Get-ParaText $p $nsm).Trim(); if (-not $t) { continue }
            $m = [regex]::Match($t, '^(\d{1,3})[\.\)](?:\s+|(?=[^\d\s]))(?=\S)')
            if ($m.Success -and ($runL.Count -eq 0 -or [int]$m.Groups[1].Value -eq $prevV + 1)) { $runL.Add($p); $prevV = [int]$m.Groups[1].Value; continue }
            if ($runL.Count -ge 2) { foreach ($x in $runL) { $typedRun[$x] = $true } }
            $runL = New-Object System.Collections.Generic.List[object]; $prevV = 0
            if ($m.Success) { $runL.Add($p); $prevV = [int]$m.Groups[1].Value }
        }
        if ($runL.Count -ge 2) { foreach ($x in $runL) { $typedRun[$x] = $true } }
        $nHead = 0
        $refText = @{}   # old style name -> text of its first paragraph (for STYLEREF fields in headers/footers)
        foreach ($p in $paras) { $sid = & $pStyleOf $p; if ($headingMap.ContainsKey($sid) -and -not $typedRun.ContainsKey($p)) { if (-not $refText.ContainsKey($info[$sid].name)) { $refText[$info[$sid].name] = (Get-ParaText $p $nsm).Trim() }; Set-StyleKeepLook $p $nsm $info $snsm $sty $sid $headingMap[$sid] $defStyle; $nHead++ } }
        foreach ($k in $headingMap.Keys) { $report.Add("Headings: style '" + $info[$k].name + "' changed to '" + $info[$headingMap[$k]].name + "'") }
        # empty paragraphs in a heading style are read out as empty headings: make them plain paragraphs (same look)
        $emptyH = 0
        foreach ($p in $paras) {
            if ((Get-ParaText $p $nsm).Trim() -or $p.SelectSingleNode('.//w:drawing|.//w:pict|.//w:object|.//m:oMath', $nsm)) { continue }
            $sid = & $pStyleOf $p
            if (-not $info.ContainsKey($sid) -or -not (Test-BuiltinHeading $info[$sid].name)) { continue }
            Set-StyleKeepLook $p $nsm $info $snsm $sty $sid $defStyle $defStyle
            $emptyH++
        }
        if ($emptyH) { $report.Add("Headings: $emptyH empty paragraph(s) in a heading style made plain paragraphs (empty headings confuse screen readers)") }
        $usedIds = @{}; foreach ($p in $paras) { $usedIds[(& $pStyleOf $p)] = $true }

        # ---- 1b. lines that only LOOK like headings (bold / italic / bigger text, but no heading style)
        $isListPara = { param($p) $n0 = WAttr (SSN $p 'w:pPr/w:numPr/w:numId' $nsm) 'val'; if ($n0) { return ($n0 -ne '0') }; return [bool](Get-StyleChainValue $info (& $pStyleOf $p) 'numPr') }
        $isHeadPara = { param($p) $sid = & $pStyleOf $p; $ol = Get-StyleChainValue $info $sid 'outline'; $pol = WAttr (SSN $p 'w:pPr/w:outlineLvl' $nsm) 'val'; (($null -ne $ol -and [int]$ol -le 8) -or ($null -ne $pol -and [int]$pol -le 8)) }
        $inBox = { param($p) [bool]$p.SelectSingleNode('ancestor::w:tbl|ancestor::w:txbxContent', $nsm) }
        $docDefSz = 20; $dd = $sty.SelectSingleNode('/w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:sz', $snsm); if ($dd) { $docDefSz = [int](WAttr $dd 'val') }
        $onOff = { param($node) if ($null -eq $node) { return $null }; $v = WAttr $node 'val'; return -not ($v -in @('0', 'false', 'off')) }
        # the look of a paragraph's text: @{ b; i; caps; sz } when every text run agrees, else $null
        $paraLook = {
            param($p)
            $sid = & $pStyleOf $p
            $pr = Get-ChainProps $info $snsm $sid 'rPr'
            $res = $null
            foreach ($r in $p.SelectNodes('.//w:r[w:t]', $nsm)) {
                $t = (($r.SelectNodes('w:t', $nsm) | ForEach-Object { $_.InnerText }) -join ''); if (-not $t.Trim()) { continue }
                $rPr = SSN $r 'w:rPr' $nsm
                $cs = @{}; $rs = WAttr (SSN $rPr 'w:rStyle' $nsm) 'val'
                if ($rs) { $cn = $sty.SelectSingleNode('/w:styles/w:style[@w:styleId="' + $rs + '"]/w:rPr', $snsm); if ($cn) { foreach ($c in $cn.ChildNodes) { if ($c.NodeType -eq 'Element') { $cs[$c.LocalName] = $c } } } }
                $get = { param($k) $d = SSN $rPr ('w:' + $k) $nsm; if ($d) { return $d }; if ($cs.ContainsKey($k)) { return $cs[$k] }; if ($pr.ContainsKey($k)) { return $pr[$k] }; return $null }
                $b = & $onOff (& $get 'b'); $i = & $onOff (& $get 'i'); $c = & $onOff (& $get 'caps')
                $szn = & $get 'sz'; $sz = if ($szn) { [int](WAttr $szn 'val') } else { $docDefSz }
                $one = @{ b = [bool]$b; i = [bool]$i; caps = [bool]$c; sz = $sz }
                if ($null -eq $res) { $res = $one }
                elseif ($res.b -ne $one.b -or $res.i -ne $one.i -or $res.caps -ne $one.caps -or $res.sz -ne $one.sz) { return $null }
            }
            return $res
        }
        $szCount = @{}
        foreach ($p in $paras) { $t = (Get-ParaText $p $nsm).Trim(); if ($t.Length -ge 80 -and -not (& $inBox $p)) { $lk = & $paraLook $p; if ($lk) { $szCount[$lk.sz] = 1 + [int]$szCount[$lk.sz] } } }
        $bodySz = if ($szCount.Count) { ($szCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key } else { $docDefSz }
        $cands = New-Object System.Collections.Generic.List[object]
        $nonEmpty = @($paras | Where-Object { (Get-ParaText $_ $nsm).Trim() -or $_.SelectSingleNode('.//w:drawing|.//w:pict', $nsm) })
        for ($ix = 0; $ix -lt $nonEmpty.Count; $ix++) {
            $p = $nonEmpty[$ix]
            $t = (Get-ParaText $p $nsm).Trim()
            if ($t.Length -lt 2 -or $t.Length -gt 100 -or $t -match '[\.,;]$' -or $t -match '^\s*(\d{1,3}|[A-Za-z])[\.\)]') { continue }
            if (& $inBox $p) { continue }; if (& $isListPara $p) { continue }; if (& $isHeadPara $p) { continue }
            $sid = & $pStyleOf $p; $sn = [string]$info[$sid].name
            if ($sn -match '(?i)(title|ttl|chapter|chap|_num|ch_?no|\bnum\b|caption|toc|header|footer|table|figure|note|quote|list)') { continue }
            if (Get-StyleChainValue $info $sid 'numPr') { continue }   # a list style whose number is switched off
            if ($ix -eq $nonEmpty.Count - 1) { continue }
            $lk = & $paraLook $p; if (-not $lk) { continue }
            if (-not ($lk.b -or $lk.i -or $lk.sz -ge $bodySz + 4)) { continue }
            $cands.Add(@{ p = $p; ix = $ix; sig = ('{0:D3}|{1}|{2}|{3}' -f $lk.sz, [int]$lk.b, [int]$lk.i, [int]$lk.caps); look = $lk })
        }
        # three or more formatted lines in a row are a block of text, not headings
        $drop = @{}
        for ($k = 0; $k -lt $cands.Count; $k++) {
            $run = 1; while (($k + $run) -lt $cands.Count -and $cands[$k + $run].ix -eq $cands[$k].ix + $run -and $cands[$k + $run].sig -eq $cands[$k].sig) { $run++ }
            if ($run -ge 3) { for ($m = 0; $m -lt $run; $m++) { $drop[$k + $m] = $true } }
        }
        $cands = @(for ($k = 0; $k -lt $cands.Count; $k++) { if (-not $drop.ContainsKey($k)) { $cands[$k] } })
        # the first line of the document, bigger than any other such line, is the title
        $hasTitle = @($paras | Where-Object { $info.ContainsKey((& $pStyleOf $_)) -and $info[(& $pStyleOf $_)].name -ieq 'Title' }).Count -gt 0
        if (-not $hasTitle -and $cands.Count -gt 1 -and $cands[0].ix -eq 0) {
            $first = $cands[0]; $others = @($cands | Select-Object -Skip 1)
            if (@($others | Where-Object { $_.sig -eq $first.sig }).Count -eq 0 -and $first.look.sz -ge (($others | ForEach-Object { $_.look.sz } | Measure-Object -Maximum).Maximum)) {
                $tId = Use-BuiltinStyle $sty $snsm $info 'Title' 'Title' $null $usedIds
                Set-StyleKeepLook $first.p $nsm $info $snsm $sty (& $pStyleOf $first.p) $tId $defStyle
                $report.Add("Title: '" + (Get-ParaText $first.p $nsm).Trim() + "' set to the Title style")
                $cands = $others; $hasTitle = $true
            }
        }
        if ($cands.Count -gt 0) {
            $sigs = @($cands | ForEach-Object { $_.sig } | Sort-Object -Unique | Sort-Object @{ Expression = { [int]$_.Split('|')[0] }; Descending = $true }, @{ Expression = { [int]$_.Split('|')[1] }; Descending = $true }, @{ Expression = { [int]$_.Split('|')[3] }; Descending = $true })
            $rankOf = @{}; for ($k = 0; $k -lt $sigs.Count; $k++) { $rankOf[$sigs[$k]] = $k + 1 }
            $candSet = @{}; foreach ($c in $cands) { $candSet[$c.p] = $c }
            $base = 0; $made = 0
            foreach ($p in $paras) {
                if ($candSet.ContainsKey($p)) {
                    $lvl = [math]::Min(6, $base + $rankOf[$candSet[$p].sig])
                    $tId = Use-BuiltinStyle $sty $snsm $info ('heading ' + $lvl) ('Heading' + $lvl) $null $usedIds ($lvl - 1)
                    Set-StyleKeepLook $p $nsm $info $snsm $sty (& $pStyleOf $p) $tId $defStyle
                    $made++
                    continue
                }
                $sid = & $pStyleOf $p
                if ($info.ContainsKey($sid) -and (Test-BuiltinHeading $info[$sid].name) -and (Get-ParaText $p $nsm).Trim()) { $base = [int]($info[$sid].name -replace '\D', '') }
            }
            if ($made) { $report.Add("Headings: $made line(s) that were only bold / italic / bigger text made real headings (keeping their look)") }
        }
        # titles of boxes and features (styles like "BoxTitle", "Feature Head"): a heading one level below the heading before them
        $boxMade = 0; $base = 0
        foreach ($p in $paras) {
            $sid = & $pStyleOf $p; if (-not $info.ContainsKey($sid)) { continue }
            $nm = [string]$info[$sid].name; $t = (Get-ParaText $p $nsm).Trim()
            if ((Test-BuiltinHeading $nm) -and $t) { $base = [int]($nm -replace '\D', ''); continue }
            if (-not $t -or $t.Length -gt 150 -or $nm -notmatch '(?i)(box|sidebar|feature|case|callout)\s*_?(title|head|heading|ttl)') { continue }
            if ((& $inBox $p) -or (& $isListPara $p) -or (& $isHeadPara $p)) { continue }
            $lvl = [math]::Min(6, $base + 1)
            $tId = Use-BuiltinStyle $sty $snsm $info ('heading ' + $lvl) ('Heading' + $lvl) $null $usedIds ($lvl - 1)
            Set-StyleKeepLook $p $nsm $info $snsm $sty $sid $tId $defStyle
            $boxMade++
        }
        if ($boxMade) { $report.Add("Headings: $boxMade box / feature title(s) made real headings (keeping their look)") }
        # headings that repeat the same opening words with numbers (e.g. "NAACLS Entry Level Competencies: 9.0 ...")
        # are labels, not headings
        $hp = @($paras | Where-Object { $s0 = & $pStyleOf $_; $info.ContainsKey($s0) -and (Test-BuiltinHeading $info[$s0].name) -and (Get-ParaText $_ $nsm).Trim() })
        $lead = @{}; foreach ($p in $hp) { $w = ((Get-ParaText $p $nsm).Trim() -split '\s+' | Select-Object -First 3) -join ' '; $lead[$w] = 1 + [int]$lead[$w] }
        $demoted = 0
        foreach ($p in $hp) {
            $t = (Get-ParaText $p $nsm).Trim(); $w = ($t -split '\s+' | Select-Object -First 3) -join ' '
            if ($lead[$w] -lt 3 -or ([regex]::Matches($t, '\d+(\.\d+)?')).Count -lt 2) { continue }
            $next = $null; $seen = $false
            foreach ($q in $paras) { if ($seen -and (Get-ParaText $q $nsm).Trim() -and -not (& $isHeadPara $q)) { $next = $q; break }; if ([object]::ReferenceEquals($q, $p)) { $seen = $true } }
            $bodySid = if ($next) { & $pStyleOf $next } else { $defStyle }
            if ($info.ContainsKey($bodySid) -and (Test-BuiltinHeading $info[$bodySid].name)) { $bodySid = $defStyle }
            Set-StyleKeepLook $p $nsm $info $snsm $sty (& $pStyleOf $p) $bodySid $defStyle
            $demoted++
        }
        if ($demoted) { $report.Add("Headings: $demoted repeated label line(s) (same words and numbers each time) were wrongly marked as headings - made normal text") }
        # the first line is a heading 1 and more heading 1s follow: it is the document's title (e.g. the book title)
        if (-not $hasTitle) {
            $firstP = @($paras | Where-Object { (Get-ParaText $_ $nsm).Trim() })[0]
            if ($firstP) {
                $s1 = & $pStyleOf $firstP
                $h1s = @($paras | Where-Object { $s0 = & $pStyleOf $_; $info.ContainsKey($s0) -and $info[$s0].name -ieq 'heading 1' -and (Get-ParaText $_ $nsm).Trim() })
                # ... and the very next line with text is another heading 1 (no text in between)
                $nextT = $null; $seen1 = $false
                foreach ($q in $paras) { if ($seen1 -and (Get-ParaText $q $nsm).Trim()) { $nextT = $q; break }; if ([object]::ReferenceEquals($q, $firstP)) { $seen1 = $true } }
                $nextIsH1 = $nextT -and $info.ContainsKey((& $pStyleOf $nextT)) -and $info[(& $pStyleOf $nextT)].name -ieq 'heading 1'
                if ($info.ContainsKey($s1) -and $info[$s1].name -ieq 'heading 1' -and $nextIsH1 -and (Get-ParaText $firstP $nsm).Trim().Length -le 150) {
                    $tId = Use-BuiltinStyle $sty $snsm $info 'Title' 'Title' $null $usedIds
                    Set-StyleKeepLook $firstP $nsm $info $snsm $sty $s1 $tId $defStyle
                    $report.Add("Title: '" + (Get-ParaText $firstP $nsm).Trim() + "' (the first line, a heading 1) set to the Title style")
                }
            }
        }
        $usedIds = @{}; foreach ($p in $paras) { $usedIds[(& $pStyleOf $p)] = $true }

        # ---- 2. title lines at the top
        $isHeading = { param($p) $sid = & $pStyleOf $p; $ol = Get-StyleChainValue $info $sid 'outline'; $pol = WAttr (SSN $p 'w:pPr/w:outlineLvl' $nsm) 'val'; (($null -ne $ol -and [int]$ol -le 8) -or ($null -ne $pol -and [int]$pol -le 8)) }
        $titleParas = @()
        foreach ($p in $paras) {
            $txt = (Get-ParaText $p $nsm).Trim()
            if (-not $txt) { continue }
            if (& $isHeading $p) { break }
            $sid = & $pStyleOf $p
            if ($info.ContainsKey($sid) -and $info[$sid].name -ieq 'Title') { $titleParas += $p; continue }
            if ($txt.Length -le 150 -and $titleParas.Count -lt 3 -and $info.ContainsKey($sid) -and $info[$sid].name -match '(?i)(title|ttl|chapter|chap|_num|ch_?no|\bnum\b)') { $titleParas += $p; continue }
            break
        }
        $titleText = ''
        if ($titleParas.Count -gt 0) {
            $lookFrom = & $pStyleOf $titleParas[-1]
            $titleId = $null
            foreach ($p in $titleParas) {
                $sid = & $pStyleOf $p
                if ($info.ContainsKey($sid) -and $info[$sid].name -ieq 'Title') { continue }
                if (-not $titleId) { $titleId = Use-BuiltinStyle $sty $snsm $info 'Title' 'Title' $lookFrom $usedIds }
                if (-not $refText.ContainsKey($info[$sid].name)) { $refText[$info[$sid].name] = (Get-ParaText $p $nsm).Trim() }
                Set-StyleKeepLook $p $nsm $info $snsm $sty $sid $titleId $defStyle
                $report.Add("Title: '" + (Get-ParaText $p $nsm).Trim() + "' set to the Title style")
            }
            $parts = @()
            foreach ($p in $titleParas) {
                $t = (Get-ParaText $p $nsm).Trim()
                $caps = $p.SelectSingleNode('.//w:rPr/w:caps|.//w:rPr/w:smallCaps', $nsm)
                $sid0 = & $pStyleOf $p
                if (-not $caps -and $info.ContainsKey($sid0)) { $caps = $info[$sid0].node.SelectSingleNode('w:rPr/w:caps|w:rPr/w:smallCaps', $snsm) }
                if ($caps -and (WAttr $caps 'val') -notin @('0', 'false')) { $t = $t.ToUpperInvariant() }
                if ($t -match '^(?i)chapter\s+\d+$') { $t = $t.ToUpperInvariant() }
                $parts += $t
                # stop once the title itself is in (a line that is only "Chapter 6" needs the next line too)
                if ($t -notmatch '^(?i)(chapter|chap\.?|lesson|unit|part|module|section)?\s*[\dIVXLC]+[\.:]?$') { break }
            }
            $titleText = ($parts -join ' ').Trim()
        }

        # ---- 3. lists
        if ($num) {
            $numToAbs = @{}
            foreach ($n in $num.SelectNodes('/w:numbering/w:num', $nnsm)) { $numToAbs[(WAttr $n 'numId')] = (WAttr (SSN $n 'w:abstractNumId' $nnsm) 'val') }
            $absNode = @{}
            foreach ($a in $num.SelectNodes('/w:numbering/w:abstractNum', $nnsm)) { $absNode[(WAttr $a 'abstractNumId')] = $a }
            # effective numbering of every paragraph
            $eff = @{}
            foreach ($p in $paras) {
                $sid = & $pStyleOf $p
                $pn = $p.SelectSingleNode('w:pPr/w:numPr', $nsm)
                $sn = Get-StyleChainValue $info $sid 'numPr'
                $nid = WAttr (SSN $pn 'w:numId' $nsm) 'val'
                $il = WAttr (SSN $pn 'w:ilvl' $nsm) 'val'
                if (-not $nid -and $sn) { $nid = WAttr (SSN $sn 'w:numId' $snsm) 'val' }
                if (-not $il -and $sn -and -not (WAttr (SSN $pn 'w:numId' $nsm) 'val')) { $il = WAttr (SSN $sn 'w:ilvl' $snsm) 'val' }
                if (-not $il) { $il = '0' }
                $styleNid = $null; $styleIl = $null
                if ($sn) { $styleNid = WAttr (SSN $sn 'w:numId' $snsm) 'val'; $styleIl = WAttr (SSN $sn 'w:ilvl' $snsm) 'val'; if (-not $styleIl) { $styleIl = '0' } }
                $eff[$p] = @{ numId = $nid; ilvl = [int]$il; style = $sid; styleNumId = $styleNid; styleIlvl = $styleIl; heading = (& $isHeading $p) }
            }
            # 3a. items switched off by accident (numId 0) or sitting in a copy of the same list - only when they sit
            #     right inside their own list (the list item before them belongs to the main list), so separate lists
            #     that were deliberately restarted are left alone.
            $fixedOff = 0; $merged = 0
            $prevItemOf = @{}; $nextItemOf = @{}; $lastItem = $null
            foreach ($p in $paras) {
                $e = $eff[$p]
                $txt = (Get-ParaText $p $nsm).Trim()
                $listish = ($e.numId -and $numToAbs.ContainsKey($e.numId)) -or ($e.numId -eq '0' -and $e.styleNumId)
                if ($listish -and -not $e.heading) {
                    if ($lastItem) { $prevItemOf[$p] = $lastItem; $nextItemOf[$lastItem] = $p }
                    $lastItem = $p
                } elseif ($e.heading -or $txt) { $lastItem = $null }
            }
            foreach ($p in $paras) {
                $e = $eff[$p]
                if (-not $e.styleNumId -or $e.styleNumId -eq '0' -or $e.heading) { continue }
                if (-not $numToAbs.ContainsKey($e.styleNumId)) { continue }
                $styleAbs = $numToAbs[$e.styleNumId]
                $pn = $p.SelectSingleNode('w:pPr/w:numPr', $nsm)
                if (-not $pn) { continue }
                $prev = $prevItemOf[$p]; $next = $nextItemOf[$p]
                $prevMain = $prev -and $eff[$prev].numId -eq $e.styleNumId
                if (-not $prevMain) { continue }
                if ($e.numId -eq '0') {
                    if (-not (Get-ParaText $p $nsm).Trim()) { continue }
                    $nextSame = $next -and $eff[$next].numId -and $numToAbs.ContainsKey($eff[$next].numId) -and $numToAbs[$eff[$next].numId] -eq $styleAbs
                    if (-not $nextSame) { continue }
                    [void]$pn.ParentNode.RemoveChild($pn); $e.numId = $e.styleNumId; $e.ilvl = [int]$e.styleIlvl; $fixedOff++
                } elseif ($e.numId -ne $e.styleNumId -and $numToAbs.ContainsKey($e.numId) -and $numToAbs[$e.numId] -eq $styleAbs) {
                    $nextMain = $next -and $eff[$next].numId -eq $e.styleNumId
                    if (-not $nextMain) { continue }
                    [void]$pn.ParentNode.RemoveChild($pn); $e.numId = $e.styleNumId; $e.ilvl = [int]$e.styleIlvl; $merged++
                }
            }
            if ($fixedOff) { $report.Add("Lists: $fixedOff list item(s) had their numbering switched off - put back into their list") }
            if ($merged) { $report.Add("Lists: $merged list item(s) were in a separate copy of their list - joined to the main list") }

            $containerOf = {
                param($p)
                $n = $p.ParentNode
                while ($n -and $n.LocalName -ne 'tc' -and $n.LocalName -ne 'txbxContent' -and $n.LocalName -ne 'body') { $n = $n.ParentNode }
                return $n
            }
            # 3a2. a separate list that is only indented further (e.g. "List Bullet 3" under "List Bullet 2") is
            #      really a sublist: it joins the list above it one level deeper. Only done when that deeper level
            #      is free, so it can take over the sublist's own bullet / number look.
            $indentOf = {
                param($p)
                $e = $eff[$p]
                $d = SSN $p 'w:pPr/w:ind' $nsm
                if ($d) { $v = WAttr $d 'left'; if (-not $v) { $v = WAttr $d 'start' }; if ($v) { return [int]$v } }
                if ($e.numId -and $numToAbs.ContainsKey($e.numId) -and $absNode.ContainsKey($numToAbs[$e.numId])) {
                    $lv = $absNode[$numToAbs[$e.numId]].SelectSingleNode('w:lvl[@w:ilvl="' + $e.ilvl + '"]/w:pPr/w:ind', $nnsm)
                    if ($lv) { $v = WAttr $lv 'left'; if (-not $v) { $v = WAttr $lv 'start' }; if ($v) { return [int]$v } }
                }
                $sp = Get-ChainProps $info $snsm $e.style 'pPr'
                if ($sp.ContainsKey('ind')) { $v = WAttr $sp['ind'] 'left'; if (-not $v) { $v = WAttr $sp['ind'] 'start' }; if ($v) { return [int]$v } }
                return 0
            }
            $levelUse = @{}
            foreach ($p in $paras) { $e = $eff[$p]; if ($e.numId -and $e.numId -ne '0' -and $numToAbs.ContainsKey($e.numId)) { $k0 = $numToAbs[$e.numId] + '|' + $e.ilvl; $levelUse[$k0] = 1 + [int]$levelUse[$k0] } }
            # runs of list items in reading order (text or a heading ends a run, empty paragraphs do not)
            $seqs = New-Object System.Collections.Generic.List[object]; $cur = $null
            foreach ($p in $paras) {
                $e = $eff[$p]
                $isItem = $e.numId -and $e.numId -ne '0' -and -not $e.heading -and $numToAbs.ContainsKey($e.numId)
                if ($isItem) {
                    if (-not $cur -or -not [object]::ReferenceEquals((& $containerOf $cur[-1]), (& $containerOf $p))) { $cur = New-Object System.Collections.Generic.List[object]; $seqs.Add($cur) }
                    $cur.Add($p)
                } elseif ($e.heading -or (Get-ParaText $p $nsm).Trim()) { $cur = $null }
            }
            $styleStep = {
                param($sidA, $sidB)
                $na = [string]$info[$sidA].name; $nb = [string]$info[$sidB].name
                $ma = [regex]::Match($na, '^(.*?)\s*(\d*)$'); $mb = [regex]::Match($nb, '^(.*?)\s*(\d*)$')
                if ($ma.Groups[1].Value -ne $mb.Groups[1].Value -or -not $ma.Groups[1].Value) { return $false }
                $da = if ($ma.Groups[2].Value) { [int]$ma.Groups[2].Value } else { 1 }
                $db = if ($mb.Groups[2].Value) { [int]$mb.Groups[2].Value } else { 1 }
                return ($db -gt $da)
            }
            $joined = 0
            foreach ($sq in $seqs) {
                $i = 1
                while ($i -lt $sq.Count) {
                    $par = $sq[$i - 1]; $pe = $eff[$par]
                    $x0 = $sq[$i]; $xe = $eff[$x0]
                    $pa = $numToAbs[$pe.numId]; $xa = $numToAbs[$xe.numId]
                    if ($pa -eq $xa -or $xe.ilvl -gt $pe.ilvl -or $pe.ilvl -ge 8 -or -not $absNode.ContainsKey($pa) -or -not $absNode.ContainsKey($xa)) { $i++; continue }
                    # the run of the other list, all at one level
                    $j = $i; while ($j -lt $sq.Count -and $eff[$sq[$j]].numId -eq $xe.numId -and $eff[$sq[$j]].ilvl -eq $xe.ilvl) { $j++ }
                    $sandwich = ($j -lt $sq.Count) -and ($numToAbs[$eff[$sq[$j]].numId] -eq $pa) -and ($eff[$sq[$j]].ilvl -le $pe.ilvl)
                    $deeper = (& $indentOf $x0) -gt (& $indentOf $par) + 100
                    $named = & $styleStep $pe.style $xe.style
                    if (-not (($sandwich -and $deeper) -or $named)) { $i = $j; continue }
                    $target = $pe.ilvl + 1
                    $parentLvl = $absNode[$pa].SelectSingleNode('w:lvl[@w:ilvl="' + $target + '"]', $nnsm)
                    $childLvl = $absNode[$xa].SelectSingleNode('w:lvl[@w:ilvl="' + $xe.ilvl + '"]', $nnsm)
                    $free = -not $levelUse.ContainsKey($pa + '|' + $target)
                    $ok = $childLvl -and ($free -or $script:Adopted[$pa + '|' + $target] -eq ($xa + '|' + $xe.ilvl))
                    if (-not $ok) { $i = $j; continue }
                    if (-not $parentLvl) {
                        # a one-level list: give it the extra level
                        $parentLvl = New-W $num 'lvl'; Set-WAttr $parentLvl 'ilvl' ([string]$target)
                        $before = @($absNode[$pa].SelectNodes('w:lvl', $nnsm) | Where-Object { [int](WAttr $_ 'ilvl') -lt $target })
                        if ($before.Count) { [void]$absNode[$pa].InsertAfter($parentLvl, $before[-1]) } else { [void]$absNode[$pa].AppendChild($parentLvl) }
                        $mlt = $absNode[$pa].SelectSingleNode('w:multiLevelType', $nnsm)
                        if ($mlt) { Set-WAttr $mlt 'val' 'hybridMultilevel' }
                    }
                    if ($free) {
                        $nl = $childLvl.CloneNode($true)
                        Set-WAttr $nl 'ilvl' ([string]$target)
                        foreach ($x in @($nl.SelectNodes('w:pStyle', $nnsm))) { [void]$nl.RemoveChild($x) }
                        $lt = $nl.SelectSingleNode('w:lvlText', $nnsm)
                        if ($lt) { $v = WAttr $lt 'val'; $own = $xe.ilvl; $v = [regex]::Replace($v, '%(\d)', { param($mm) if ([int]$mm.Groups[1].Value -eq $own + 1) { '%' + ($target + 1) } else { '' } }); Set-WAttr $lt 'val' $v }
                        [void]$absNode[$pa].ReplaceChild($nl, $parentLvl)
                        $levelUse[$pa + '|' + $target] = 1
                        $script:Adopted[$pa + '|' + $target] = $xa + '|' + $xe.ilvl
                    }
                    for ($k = $i; $k -lt $j; $k++) {
                        $p = $sq[$k]; $e = $eff[$p]
                        $pPr = $p.SelectSingleNode('w:pPr', $nsm); if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.PrependChild($pPr) }
                        $pn = $pPr.SelectSingleNode('w:numPr', $nsm)
                        if (-not $pn) { $pn = New-W $doc 'numPr'; [void]$pPr.AppendChild($pn); Sort-Kids $pPr $PPR_ORDER }
                        foreach ($x in @($pn.ChildNodes)) { [void]$pn.RemoveChild($x) }
                        $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' ([string]$target); [void]$pn.AppendChild($il)
                        $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' $pe.numId; [void]$pn.AppendChild($ni)
                        $e.numId = $pe.numId; $e.ilvl = $target
                        $joined++
                    }
                    $i = $j
                }
            }
            if ($joined) { $report.Add("Lists: $joined item(s) of a separate, further-indented list made a real sublist of the list above them") }

            # 3b. list levels: inside every heading section, the outermost list level becomes level 1,
            #     its sublist level 2, and so on, with no gaps. The look of every item stays the same:
            #     the level definitions move together with the items.
            $absUsedByHeading = @{}
            foreach ($p in $paras) { $e = $eff[$p]; if ($e.heading -and $e.numId -and $numToAbs.ContainsKey($e.numId)) { $absUsedByHeading[$numToAbs[$e.numId]] = $true } }
            foreach ($sid in $info.Keys) {
                $ol = Get-StyleChainValue $info $sid 'outline'; $sn = $info[$sid].numPr
                if ($null -ne $ol -and $sn) { $nid = WAttr (SSN $sn 'w:numId' $snsm) 'val'; if ($nid -and $numToAbs.ContainsKey($nid)) { $absUsedByHeading[$numToAbs[$nid]] = $true } }
            }
            # runs of list items (empty paragraphs do not break a run; text, headings and table cells do)
            $groups = New-Object System.Collections.Generic.List[object]
            $cur = $null; $section = 0
            foreach ($p in $paras) {
                $e = $eff[$p]
                $isItem = $e.numId -and $e.numId -ne '0' -and -not $e.heading -and $numToAbs.ContainsKey($e.numId)
                $cont = & $containerOf $p
                if ($e.heading) { $section++; $cur = $null; continue }
                if ($isItem) {
                    if (-not $cur -or -not [object]::ReferenceEquals($cur.cont, $cont)) {
                        $cur = @{ cont = $cont; section = $section; items = (New-Object System.Collections.Generic.List[object]); abs = @{} }
                        $groups.Add($cur)
                    }
                    $cur.items.Add($p); $cur.abs[$numToAbs[$e.numId]] = $true
                } elseif ((Get-ParaText $p $nsm).Trim()) { $cur = $null }
            }
            # runs in the same section that use the same list belong together (text between items of one list)
            $merged = New-Object System.Collections.Generic.List[object]
            foreach ($g in $groups) {
                $homeGrp = $null
                foreach ($m in $merged) {
                    if ($m.section -ne $g.section -or -not [object]::ReferenceEquals($m.cont, $g.cont)) { continue }
                    foreach ($a0 in $g.abs.Keys) { if ($m.abs.ContainsKey($a0)) { $homeGrp = $m; break } }
                    if ($homeGrp) { break }
                }
                if ($homeGrp) { foreach ($it in $g.items) { $homeGrp.items.Add($it) }; foreach ($a0 in $g.abs.Keys) { $homeGrp.abs[$a0] = $true } }
                else {
                    $li = New-Object System.Collections.Generic.List[object]; foreach ($it in $g.items) { $li.Add($it) }
                    $merged.Add(@{ cont = $g.cont; section = $g.section; items = $li; abs = $g.abs.Clone() })
                }
            }
            # wanted level for every item, and the level map each list needs
            $mapsOfAbs = @{}    # abstract -> @{ key -> map }
            $itemMapKey = @{}   # paragraph -> map key for its abstract
            foreach ($m in $merged) {
                # Each list gets its own levels 1, 2, 3 ... A list that starts right under an item of another list,
                # one level deeper than that item, is a sublist of it and continues below that item's level.
                $order = New-Object System.Collections.Generic.List[string]; $parentOf = @{}
                for ($i = 0; $i -lt $m.items.Count; $i++) {
                    $ab = $numToAbs[$eff[$m.items[$i]].numId]
                    if ($order.Contains($ab)) { continue }
                    $order.Add($ab)
                    if ($i -gt 0) {
                        $pp = $m.items[$i - 1]; $pab = $numToAbs[$eff[$pp].numId]
                        if ($pab -ne $ab -and ($eff[$m.items[$i]].ilvl -gt $eff[$pp].ilvl -or (& $indentOf $m.items[$i]) -gt (& $indentOf $pp) + 100)) { $parentOf[$ab] = @($pab, $eff[$pp].ilvl) }
                    }
                }
                $mapOf = @{}
                foreach ($ab in $order) {
                    $lv = @($m.items | Where-Object { $numToAbs[$eff[$_].numId] -eq $ab } | ForEach-Object { $eff[$_].ilvl } | Sort-Object -Unique)
                    $base = 0
                    if ($parentOf.ContainsKey($ab) -and $mapOf.ContainsKey($parentOf[$ab][0])) { $pm = $mapOf[$parentOf[$ab][0]]; if ($pm.ContainsKey($parentOf[$ab][1])) { $base = $pm[$parentOf[$ab][1]] + 1 } }
                    $mm2 = @{}; for ($i = 0; $i -lt $lv.Count; $i++) { $mm2[[int]$lv[$i]] = [math]::Min(8, $base + $i) }
                    $mapOf[$ab] = $mm2
                }
                foreach ($a0 in $m.abs.Keys) {
                    $mine = @{}
                    foreach ($p in $m.items) { if ($numToAbs[$eff[$p].numId] -eq $a0) { $mine[$eff[$p].ilvl] = $mapOf[$a0][$eff[$p].ilvl] } }
                    $key = (($mine.Keys | Sort-Object | ForEach-Object { "$_>$($mine[$_])" }) -join ',')
                    if (-not $mapsOfAbs.ContainsKey($a0)) { $mapsOfAbs[$a0] = @{} }
                    $mapsOfAbs[$a0][$key] = $mine
                    foreach ($p in $m.items) { if ($numToAbs[$eff[$p].numId] -eq $a0) { $itemMapKey[$p] = $key } }
                }
            }
            $isIdentity = { param($mp) foreach ($k in $mp.Keys) { if ($k -ne $mp[$k]) { return $false } }; return $true }
            $mapLevel = {
                param($mp, [int]$old)
                if ($mp.ContainsKey($old)) { return [int]$mp[$old] }
                $maxOld = ($mp.Keys | Measure-Object -Maximum).Maximum; $maxNew = ($mp.Values | Measure-Object -Maximum).Maximum
                if ($old -gt $maxOld) { return [int][math]::Min(8, $old - $maxOld + $maxNew) }
                $minOld = ($mp.Keys | Measure-Object -Minimum).Minimum
                if ($old -lt $minOld) { return -1 }
                # a level between two used levels: take the next used one below it
                $below = @($mp.Keys | Where-Object { $_ -lt $old } | Sort-Object -Descending)[0]
                return [int]$mp[$below]
            }
            $remapAbs = {
                param($an, $mp, [bool]$stripLinks)
                $lvls = @($an.SelectNodes('w:lvl', $nnsm))
                $byL = @{}; foreach ($l in $lvls) { $byL[[int](WAttr $l 'ilvl')] = $l }
                $inv = @{}; foreach ($k in $mp.Keys) { $inv[[int]$mp[$k]] = [int]$k }
                $maxOld = ($mp.Keys | Measure-Object -Maximum).Maximum; $maxNew = ($mp.Values | Measure-Object -Maximum).Maximum
                foreach ($l in $lvls) { [void]$an.RemoveChild($l) }
                for ($k = 0; $k -le 8; $k++) {
                    $old = if ($inv.ContainsKey($k)) { $inv[$k] } else { [math]::Min(8, $maxOld + $k - $maxNew) }
                    $src = $byL[[int]$old]; if (-not $src) { $src = $lvls[-1] }
                    $nl = $src.CloneNode($true)
                    Set-WAttr $nl 'ilvl' ([string]$k)
                    if ($stripLinks) { foreach ($x in @($nl.SelectNodes('w:pStyle', $nnsm))) { [void]$nl.RemoveChild($x) } }
                    $lt = $nl.SelectSingleNode('w:lvlText', $nnsm)
                    if ($lt) {
                        $v = WAttr $lt 'val'
                        $v = [regex]::Replace($v, '%(\d)', { param($mm) $n2 = & $mapLevel $mp ([int]$mm.Groups[1].Value - 1); if ($n2 -ge 0) { '%' + ($n2 + 1) } else { '' } })
                        Set-WAttr $lt 'val' $v
                    }
                    [void]$an.AppendChild($nl)
                }
            }
            $remapOverrides = {
                param($n, $mp)
                foreach ($ov in @($n.SelectNodes('w:lvlOverride', $nnsm))) {
                    $nl = & $mapLevel $mp ([int](WAttr $ov 'ilvl'))
                    if ($nl -lt 0) { [void]$n.RemoveChild($ov) } else { Set-WAttr $ov 'ilvl' ([string]$nl); $il2 = $ov.SelectSingleNode('w:lvl', $nnsm); if ($il2) { Set-WAttr $il2 'ilvl' ([string]$nl) } }
                }
            }
            $nextAbs = 1 + (@($absNode.Keys | ForEach-Object { [int]$_ }) + 0 | Measure-Object -Maximum).Maximum
            $nextNum = 1 + (@($numToAbs.Keys | ForEach-Object { [int]$_ }) + 0 | Measure-Object -Maximum).Maximum
            $variant = @{}   # "numId|key" -> new numId
            $levelFixes = 0
            $origIlvl = @{}; foreach ($p in $paras) { $origIlvl[$p] = $eff[$p].ilvl }
            $countOfKey = @{}; foreach ($p in $itemMapKey.Keys) { $k0 = $numToAbs[$eff[$p].numId] + '#' + $itemMapKey[$p]; $countOfKey[$k0] = 1 + [int]$countOfKey[$k0] }
            # two level maps can be used together when they agree wherever they overlap and keep the order
            $mergeMaps = {
                param($m1, $m2)
                $u = @{}; foreach ($k in $m1.Keys) { $u[$k] = $m1[$k] }
                foreach ($k in $m2.Keys) { if ($u.ContainsKey($k) -and $u[$k] -ne $m2[$k]) { return $null }; $u[$k] = $m2[$k] }
                $ks = @($u.Keys | Sort-Object); for ($i = 1; $i -lt $ks.Count; $i++) { if ($u[$ks[$i]] -le $u[$ks[$i - 1]]) { return $null } }
                return $u
            }
            foreach ($a in @($mapsOfAbs.Keys)) {
                if (-not $absNode.ContainsKey($a)) { continue }
                $an = $absNode[$a]
                if ($an.SelectNodes('w:lvl', $nnsm).Count -eq 0) { continue }   # list defined by a list style elsewhere
                $keys = @($mapsOfAbs[$a].Keys | Sort-Object { - [int]$countOfKey[$a + '#' + $_] })
                $clusters = New-Object System.Collections.Generic.List[object]
                foreach ($k in $keys) {
                    $done = $false
                    foreach ($c in $clusters) { $u = & $mergeMaps $c.map $mapsOfAbs[$a][$k]; if ($u) { $c.map = $u; $c.keys += $k; $done = $true; break } }
                    if (-not $done) { $clusters.Add(@{ map = $mapsOfAbs[$a][$k].Clone(); keys = @($k) }) }
                }
                if (@($clusters | Where-Object { -not (& $isIdentity $_.map) }).Count -eq 0) { continue }
                $orig = $an.CloneNode($true)
                $inPlace = $false
                $main = $clusters[0]
                if (-not (& $isIdentity $main.map) -and -not $absUsedByHeading.ContainsKey($a)) {
                    # the way most of this list is used: change the list itself
                    $inPlace = $true
                    $mp = $main.map
                    & $remapAbs $an $mp $false
                    foreach ($n in $num.SelectNodes('/w:numbering/w:num', $nnsm)) { if ((WAttr (SSN $n 'w:abstractNumId' $nnsm) 'val') -eq $a) { & $remapOverrides $n $mp } }
                    foreach ($sid in $info.Keys) {
                        $sn = $info[$sid].numPr; if (-not $sn) { continue }
                        $nid = WAttr (SSN $sn 'w:numId' $snsm) 'val'
                        if (-not $nid -or -not $numToAbs.ContainsKey($nid) -or $numToAbs[$nid] -ne $a) { continue }
                        $il = $sn.SelectSingleNode('w:ilvl', $snsm)
                        $old = if ($il) { [int](WAttr $il 'val') } else { 0 }
                        $nw = & $mapLevel $mp $old; if ($nw -lt 0) { $nw = 0 }
                        if (-not $il) { $il = New-W $sty 'ilvl'; [void]$sn.PrependChild($il) }
                        Set-WAttr $il 'val' ([string]$nw)
                    }
                    foreach ($p in $paras) {
                        $e = $eff[$p]
                        if (-not $e.numId -or -not $numToAbs.ContainsKey($e.numId) -or $numToAbs[$e.numId] -ne $a) { continue }
                        $nw = & $mapLevel $mp $e.ilvl; if ($nw -lt 0) { $nw = 0 }
                        if ($nw -ne $e.ilvl -and $itemMapKey.ContainsKey($p) -and ($main.keys -contains $itemMapKey[$p])) { $levelFixes++ }
                        $e.ilvl = $nw
                        $il = $p.SelectSingleNode('w:pPr/w:numPr/w:ilvl', $nsm)
                        if ($il) { Set-WAttr $il 'val' ([string]$nw) }
                    }
                }
                # every other way this list is used gets its own copy of the list, with its own levels
                for ($ci = 0; $ci -lt $clusters.Count; $ci++) {
                    $c = $clusters[$ci]
                    if ($inPlace -and $ci -eq 0) { continue }
                    if ((& $isIdentity $c.map) -and -not $inPlace) { continue }
                    foreach ($p in $paras) {
                        if (-not $itemMapKey.ContainsKey($p) -or ($c.keys -notcontains $itemMapKey[$p])) { continue }
                        $e = $eff[$p]
                        if ($numToAbs[$e.numId] -ne $a) { continue }
                        $vk = $e.numId + '|' + $a + '|' + $ci
                        if (-not $variant.ContainsKey($vk)) {
                            $na = $orig.CloneNode($true)
                            Set-WAttr $na 'abstractNumId' ([string]$nextAbs)
                            foreach ($x in @($na.SelectNodes('w:styleLink|w:numStyleLink', $nnsm))) { [void]$na.RemoveChild($x) }
                            $nsid = $na.SelectSingleNode('w:nsid', $nnsm); if ($nsid) { Set-WAttr $nsid 'val' ('{0:X8}' -f (Get-Random -Minimum 268435456 -Maximum 2147483647)) }
                            if (-not (& $isIdentity $c.map)) { & $remapAbs $na $c.map $true } else { foreach ($l in $na.SelectNodes('w:lvl', $nnsm)) { foreach ($x in @($l.SelectNodes('w:pStyle', $nnsm))) { [void]$l.RemoveChild($x) } } }
                            $firstNum = $num.SelectSingleNode('/w:numbering/w:num', $nnsm)
                            if ($firstNum) { [void]$num.DocumentElement.InsertBefore($na, $firstNum) } else { [void]$num.DocumentElement.AppendChild($na) }
                            $origNum = $num.SelectSingleNode('/w:numbering/w:num[@w:numId="' + $e.numId + '"]', $nnsm)
                            $nn = $origNum.CloneNode($true)
                            Set-WAttr $nn 'numId' ([string]$nextNum)
                            Set-WAttr $nn.SelectSingleNode('w:abstractNumId', $nnsm) 'val' ([string]$nextAbs)
                            foreach ($x in @($nn.Attributes | Where-Object { $_.LocalName -eq 'durableId' })) { [void]$nn.Attributes.Remove($x) }
                            if ($inPlace) {
                                # the overrides were already moved with the main list: move them back first
                                $inv = @{}; foreach ($k in $main.map.Keys) { $inv[[int]$main.map[$k]] = [int]$k }
                                foreach ($ov in @($nn.SelectNodes('w:lvlOverride', $nnsm))) { $o = [int](WAttr $ov 'ilvl'); if ($inv.ContainsKey($o)) { Set-WAttr $ov 'ilvl' ([string]$inv[$o]); $l2 = $ov.SelectSingleNode('w:lvl', $nnsm); if ($l2) { Set-WAttr $l2 'ilvl' ([string]$inv[$o]) } } }
                            }
                            if (-not (& $isIdentity $c.map)) { & $remapOverrides $nn $c.map }
                            $cleanup = $num.SelectSingleNode('/w:numbering/w:numIdMacAtCleanup', $nnsm)
                            if ($cleanup) { [void]$num.DocumentElement.InsertBefore($nn, $cleanup) } else { [void]$num.DocumentElement.AppendChild($nn) }
                            $absNode[[string]$nextAbs] = $na; $numToAbs[[string]$nextNum] = [string]$nextAbs
                            $variant[$vk] = [string]$nextNum
                            $nextAbs++; $nextNum++
                        }
                        $o = $origIlvl[$p]
                        $nw = if ($c.map.ContainsKey($o)) { [int]$c.map[$o] } else { $o }
                        $pPr = $p.SelectSingleNode('w:pPr', $nsm)
                        if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.PrependChild($pPr) }
                        $pn = $pPr.SelectSingleNode('w:numPr', $nsm)
                        if (-not $pn) { $pn = New-W $doc 'numPr'; [void]$pPr.AppendChild($pn); Sort-Kids $pPr $PPR_ORDER }
                        foreach ($x in @($pn.ChildNodes)) { [void]$pn.RemoveChild($x) }
                        $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' ([string]$nw); [void]$pn.AppendChild($il)
                        $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' $variant[$vk]; [void]$pn.AppendChild($ni)
                        if ($nw -ne $o) { $levelFixes++ }
                        $e.numId = $variant[$vk]; $e.ilvl = $nw
                    }
                }
            }
            if ($levelFixes) { $report.Add("Lists: $levelFixes list item(s) moved to the right list level (first level after each heading = level 1, each sublist one level deeper, no gaps)") }

            # 3c. numbers typed by hand -> real numbering
            $usedLv = @{}
            foreach ($p in $paras) { $e = $eff[$p]; if ($e.numId -and $e.numId -ne '0' -and -not $e.heading -and $numToAbs.ContainsKey($e.numId)) { $usedLv[$numToAbs[$e.numId]] = $true } }
            foreach ($a in @($usedLv.Keys)) {
                if (-not $absNode.ContainsKey($a)) { continue }
                $an = $absNode[$a]
                for ($k = 0; $k -le 8; $k++) {
                    $lv = $an.SelectSingleNode('w:lvl[@w:ilvl="' + $k + '"]', $nnsm)
                    if (-not $lv) { continue }
                    $fmt = WAttr (SSN $lv 'w:numFmt' $nnsm) 'val'
                    $ltxt = WAttr (SSN $lv 'w:lvlText' $nnsm) 'val'
                    if ($fmt -ne 'none' -and $ltxt -match '%') { continue }
                    $items = @($paras | Where-Object { $e = $eff[$_]; $e.numId -and $numToAbs.ContainsKey($e.numId) -and $numToAbs[$e.numId] -eq $a -and $e.ilvl -eq $k })
                    if ($items.Count -lt 2) { continue }
                    $typed = @($items | ForEach-Object { Get-TypedNumber (Get-ParaText $_ $nsm) })
                    $kinds = @($typed | Where-Object { $_ } | ForEach-Object { $_.kind } | Group-Object | Sort-Object Count -Descending)
                    if ($kinds.Count -eq 0) { continue }
                    $kind = $kinds[0].Name
                    $hits = $kinds[0].Count
                    if ($hits -lt [math]::Ceiling($items.Count * 0.8)) { continue }
                    $nf = $lv.SelectSingleNode('w:numFmt', $nnsm); if (-not $nf) { $nf = New-W $num 'numFmt'; [void]$lv.AppendChild($nf) }
                    Set-WAttr $nf 'val' $kind
                    $lt = $lv.SelectSingleNode('w:lvlText', $nnsm); if (-not $lt) { $lt = New-W $num 'lvlText'; [void]$lv.AppendChild($lt) }
                    Set-WAttr $lt 'val' ('%' + ($k + 1) + '.')
                    $st = $lv.SelectSingleNode('w:start', $nnsm); if (-not $st) { $st = New-W $num 'start'; [void]$lv.PrependChild($st) }
                    Set-WAttr $st 'val' '1'
                    $sfx = $lv.SelectSingleNode('w:suff', $nnsm); if (-not $sfx) { $sfx = New-W $num 'suff'; [void]$lv.InsertBefore($sfx, $lv.SelectSingleNode('w:lvlText', $nnsm)) }
                    Set-WAttr $sfx 'val' 'space'
                    # remove the typed numbers; where the typed numbers start again from 1 (a new section),
                    # the real numbering starts again from 1 there too
                    $expect = 0; $odd = @(); $restartAt = @{}
                    for ($i = 0; $i -lt $items.Count; $i++) {
                        $expect++
                        $tn = $typed[$i]
                        if ($tn -and $tn.kind -eq $kind) {
                            if ($tn.value -eq 1 -and $expect -ne 1) { $restartAt[$items[$i]] = $true; $expect = 1 }
                            if ($tn.value -ne $expect) { $odd += ("'" + (Get-ParaText $items[$i] $nsm).Trim() + "' was typed as " + (Format-Number $kind $tn.value) + ", Word now numbers it " + (Format-Number $kind $expect)) }
                            $script:TypedRemoved[$items[$i]] = (Get-ParaText $items[$i] $nsm).Substring(0, $tn.length)
                            Remove-LeadingChars $items[$i] $nsm $tn.length
                        } else { $odd += ("'" + (Get-ParaText $items[$i] $nsm).Trim() + "' had no typed number, Word now numbers it " + (Format-Number $kind $expect)) }
                    }
                    if ($restartAt.Count -gt 0) {
                        $nextNum2 = 1 + (@($num.SelectNodes('/w:numbering/w:num', $nnsm) | ForEach-Object { [int](WAttr $_ 'numId') }) + 0 | Measure-Object -Maximum).Maximum
                        $segNum = $null
                        foreach ($p in $paras) {
                            $e = $eff[$p]
                            if (-not $e.numId -or -not $numToAbs.ContainsKey($e.numId) -or $numToAbs[$e.numId] -ne $a) { continue }
                            if ($restartAt.ContainsKey($p)) {
                                $src = $num.SelectSingleNode('/w:numbering/w:num[@w:numId="' + $e.numId + '"]', $nnsm)
                                $nn = $src.CloneNode($true)
                                Set-WAttr $nn 'numId' ([string]$nextNum2)
                                foreach ($x in @($nn.Attributes | Where-Object { $_.LocalName -eq 'durableId' })) { [void]$nn.Attributes.Remove($x) }
                                foreach ($ov in @($nn.SelectNodes('w:lvlOverride[@w:ilvl="' + $k + '"]', $nnsm))) { [void]$nn.RemoveChild($ov) }
                                $ov = New-W $num 'lvlOverride'; Set-WAttr $ov 'ilvl' ([string]$k)
                                $so = New-W $num 'startOverride'; Set-WAttr $so 'val' '1'; [void]$ov.AppendChild($so); [void]$nn.AppendChild($ov)
                                $cleanup = $num.SelectSingleNode('/w:numbering/w:numIdMacAtCleanup', $nnsm)
                                if ($cleanup) { [void]$num.DocumentElement.InsertBefore($nn, $cleanup) } else { [void]$num.DocumentElement.AppendChild($nn) }
                                $numToAbs[[string]$nextNum2] = $a
                                $segNum = [string]$nextNum2; $nextNum2++
                            }
                            if ($segNum) {
                                $pPr = $p.SelectSingleNode('w:pPr', $nsm); if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.PrependChild($pPr) }
                                $pn = $pPr.SelectSingleNode('w:numPr', $nsm)
                                if (-not $pn) { $pn = New-W $doc 'numPr'; [void]$pPr.AppendChild($pn); Sort-Kids $pPr $PPR_ORDER }
                                foreach ($x in @($pn.ChildNodes)) { [void]$pn.RemoveChild($x) }
                                $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' ([string]$e.ilvl); [void]$pn.AppendChild($il)
                                $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' $segNum; [void]$pn.AppendChild($ni)
                                $e.numId = $segNum
                            }
                        }
                        $report.Add("Lists: the typed numbers started again from " + (Format-Number $kind 1) + " in " + $restartAt.Count + " place(s) - the real numbering starts again there too")
                    }
                    $report.Add("Lists: $hits typed numbers (" + (Format-Number $kind 1) + ", " + (Format-Number $kind 2) + ", ...) at list level " + ($k + 1) + " replaced by real list numbering")
                    foreach ($o in $odd) { $report.Add('CHECK: ' + $o) }
                }
            }
        }

        # ---- 3d. lists typed by hand ("1. U", "2. C" ...): they become real numbered lists. The number shown stays
        #      exactly the same (also where a list carries on after a heading, e.g. 27, 28 ...).
        $script:Absorbed = @{}
        $typedSegs = New-Object System.Collections.Generic.List[object]
        $isRealItem = { param($p) $n0 = WAttr (SSN $p 'w:pPr/w:numPr/w:numId' $nsm) 'val'; if ($n0) { return ($n0 -ne '0') }; return [bool](Get-StyleChainValue $info (& $pStyleOf $p) 'numPr') }
        $typedOf = {
            param($p)
            if ($p.SelectSingleNode('ancestor::w:tbl|ancestor::w:txbxContent', $nsm)) { return $null }
            if (& $isRealItem $p) { return $null }
            $s0 = & $pStyleOf $p; if ($info.ContainsKey($s0) -and ((Test-BuiltinHeading $info[$s0].name) -or $info[$s0].name -ieq 'Title')) { return $null }
            $t = Get-ParaText $p $nsm
            $m = [regex]::Match($t, '^(\s*(\d{1,3})[\.\)](?:\s+|(?=[^\d\s])))(?=\S)')
            if ($m.Success) { return @{ kind = 'decimal'; value = [int]$m.Groups[2].Value; len = $m.Groups[1].Length } }
            $m = [regex]::Match($t, '^(\s*([A-Z])[\.\)]\s+)(?=\S)')
            if ($m.Success) { return @{ kind = 'upperLetter'; value = ([int][char]$m.Groups[2].Value - 64); len = $m.Groups[1].Length } }
            $m = [regex]::Match($t, '^(\s*([a-z])[\.\)]\s+)(?=\S)')
            if ($m.Success) { return @{ kind = 'lowerLetter'; value = ([int][char]$m.Groups[2].Value - 96); len = $m.Groups[1].Length } }
            return $null
        }
        $cur = $null
        foreach ($p in $paras) {
            $t = (Get-ParaText $p $nsm).Trim()
            if (-not $t) { continue }   # empty lines do not break a typed list
            $tn = & $typedOf $p
            if ($tn -and $cur -and $cur.kind -eq $tn.kind -and $tn.value -eq $cur.last + 1 -and (& $pStyleOf $p) -eq $cur.style -and $cur.open) {
                $cur.items.Add(@{ p = $p; tn = $tn }); $cur.last = $tn.value; continue
            }
            if ($tn) {
                $cur = @{ kind = $tn.kind; first = $tn.value; last = $tn.value; style = (& $pStyleOf $p); items = (New-Object System.Collections.Generic.List[object]); open = $true }
                $cur.items.Add(@{ p = $p; tn = $tn }); $typedSegs.Add($cur); continue
            }
            if ($cur) { $cur.open = $false; $cur = $null }
        }
        $toConvert = @($typedSegs | ForEach-Object { @{ seg = $_; join = $null } })
        $joinLast = @{}   # style -> @{ numId; ilvl; value } of the last typed item made a real list item
        if ($toConvert.Count -gt 0) {
            if (-not $num) {
                if (@($toConvert | Where-Object { $_.seg.items.Count -ge 2 }).Count -eq 0) { $toConvert = @() }
                $num = New-XmlDoc ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="' + $NS_W + '"/>'); $nnsm = New-Ns $num
                $script:NewNumbering = $true
            }
            $nextA = 1 + (@($num.SelectNodes('/w:numbering/w:abstractNum', $nnsm) | ForEach-Object { [int](WAttr $_ 'abstractNumId') }) + 0 | Measure-Object -Maximum).Maximum
            $nextN = 1 + (@($num.SelectNodes('/w:numbering/w:num', $nnsm) | ForEach-Object { [int](WAttr $_ 'numId') }) + 0 | Measure-Object -Maximum).Maximum
            $absFor = @{}
            $nItems = 0
            foreach ($c in $toConvert) {
                $sg = $c.seg
                if ($sg.items.Count -lt 2) {
                    # a single typed item counts only when it carries on a real list just before it (e.g. questions 1, 2
                    # are a real list and "3." was typed)
                    if ($sg.kind -ne 'decimal' -or $sg.first -lt 2) { continue }
                    $p0 = $sg.items[0].p; $lastReal = $null
                    foreach ($q in $paras) { if ([object]::ReferenceEquals($q, $p0)) { break }; if ($q.ParentNode -and (& $isRealItem $q) -and ((& $pStyleOf $q) -eq $sg.style)) { $lastReal = $q } }
                    if (-not $lastReal) { continue }
                    $curLabels = Get-ListLabels @($paras | Where-Object { $_.ParentNode }) $nsm $sty $snsm $num $nnsm $defStyle
                    $lv = [string]$curLabels[$lastReal]
                    if ($lv -notmatch '^(\d+)\.$' -or [int]$Matches[1] -ne $sg.first - 1) { continue }
                    $pnR = SSN $lastReal 'w:pPr/w:numPr' $nsm; $snR = Get-StyleChainValue $info (& $pStyleOf $lastReal) 'numPr'
                    $jn = WAttr (SSN $pnR 'w:numId' $nsm) 'val'; if (-not $jn -and $snR) { $jn = WAttr (SSN $snR 'w:numId' $snsm) 'val' }
                    $ji = WAttr (SSN $pnR 'w:ilvl' $nsm) 'val'; if (-not $ji -and $snR) { $ji = WAttr (SSN $snR 'w:ilvl' $snsm) 'val' }; if (-not $ji) { $ji = '0' }
                    $c.join = @{ numId = $jn; ilvl = $ji }
                }
                if ($c.join) { $numId = $c.join.numId; $ilvl = $c.join.ilvl }
                else {
                    # the number sits where the typed number was: same indent, followed by a space
                    $p0 = $sg.items[0].p
                    $indL = 0; $d = SSN $p0 'w:pPr/w:ind' $nsm
                    if (-not $d) { $sp = Get-ChainProps $info $snsm (& $pStyleOf $p0) 'pPr'; if ($sp.ContainsKey('ind')) { $d = $sp['ind'] } }
                    if ($d) { $v = WAttr $d 'left'; if (-not $v) { $v = WAttr $d 'start' }; if ($v) { $indL = [int]$v }; $fl = WAttr $d 'firstLine'; if ($fl) { $indL += [int]$fl }; $hg = WAttr $d 'hanging'; if ($hg) { $indL -= [int]$hg } }
                    $ak = $sg.kind + '|' + $indL
                    if (-not $absFor.ContainsKey($ak)) {
                        $fmtTxt = '%1.'
                        $an = $num.CreateElement('w', 'abstractNum', $NS_W); Set-WAttr $an 'abstractNumId' ([string]$nextA)
                        $mt = New-W $num 'multiLevelType'; Set-WAttr $mt 'val' 'singleLevel'; [void]$an.AppendChild($mt)
                        $lv = New-W $num 'lvl'; Set-WAttr $lv 'ilvl' '0'
                        foreach ($pair in @(@('start', '1'), @('numFmt', $sg.kind), @('suff', 'space'), @('lvlText', $fmtTxt), @('lvlJc', 'left'))) { $x = New-W $num $pair[0]; Set-WAttr $x 'val' $pair[1]; [void]$lv.AppendChild($x) }
                        $lp = New-W $num 'pPr'; $ind = New-W $num 'ind'; Set-WAttr $ind 'left' ([string][math]::Max(0, $indL)); Set-WAttr $ind 'firstLine' '0'; [void]$lp.AppendChild($ind); [void]$lv.AppendChild($lp)
                        [void]$an.AppendChild($lv)
                        $firstNum = $num.SelectSingleNode('/w:numbering/w:num', $nnsm)
                        if ($firstNum) { [void]$num.DocumentElement.InsertBefore($an, $firstNum) } else { [void]$num.DocumentElement.AppendChild($an) }
                        $absFor[$ak] = [string]$nextA; $nextA++
                    }
                    $nn = $num.CreateElement('w', 'num', $NS_W); Set-WAttr $nn 'numId' ([string]$nextN)
                    $ai = New-W $num 'abstractNumId'; Set-WAttr $ai 'val' $absFor[$ak]; [void]$nn.AppendChild($ai)
                    $ov = New-W $num 'lvlOverride'; Set-WAttr $ov 'ilvl' '0'; $so = New-W $num 'startOverride'; Set-WAttr $so 'val' ([string]$sg.first); [void]$ov.AppendChild($so); [void]$nn.AppendChild($ov)
                    $cleanup = $num.SelectSingleNode('/w:numbering/w:numIdMacAtCleanup', $nnsm)
                    if ($cleanup) { [void]$num.DocumentElement.InsertBefore($nn, $cleanup) } else { [void]$num.DocumentElement.AppendChild($nn) }
                    $numId = [string]$nextN; $ilvl = '0'; $nextN++
                }
                foreach ($it in $sg.items) {
                    $p = $it.p
                    $script:TypedRemoved[$p] = (Get-ParaText $p $nsm).Substring(0, $it.tn.len)
                    Remove-LeadingChars $p $nsm $it.tn.len
                    $pPr = $p.SelectSingleNode('w:pPr', $nsm); if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.PrependChild($pPr) }
                    $pn = $pPr.SelectSingleNode('w:numPr', $nsm); if (-not $pn) { $pn = New-W $doc 'numPr'; [void]$pPr.AppendChild($pn); Sort-Kids $pPr $PPR_ORDER }
                    foreach ($x in @($pn.ChildNodes)) { [void]$pn.RemoveChild($x) }
                    $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' $ilvl; [void]$pn.AppendChild($il)
                    $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' $numId; [void]$pn.AppendChild($ni)
                    $nItems++
                    # a sentence that was broken over two paragraphs ("...After carefully" / "checking the patient's ID ...")
                    # or the rest of the question on the next line: joined into this list item
                    $nx = $p.NextSibling
                    while ($nx -and $nx.LocalName -eq 'p' -and (Get-ParaText $nx $nsm).Trim() -and -not (& $typedOf $nx) -and -not (& $isRealItem $nx) -and (& $pStyleOf $nx) -eq (& $pStyleOf $p) -and (Get-ParaText $nx $nsm).Trim() -notmatch '^(?i)answer\b' -and -not $nx.SelectSingleNode('.//w:drawing|.//w:pict|.//w:object', $nsm)) {
                        $after2 = $nx.NextSibling
                        $blockEnds = (-not $after2) -or ($after2.LocalName -ne 'p') -or -not (Get-ParaText $after2 $nsm).Trim()
                        if (-not $blockEnds) { break }
                        $sep = New-W $doc 'r'; $st = New-W $doc 't'; [void]$st.SetAttribute('xml:space', 'preserve'); $st.InnerText = ' '; [void]$sep.AppendChild($st)
                        [void]$p.AppendChild($sep)
                        foreach ($ch in @($nx.ChildNodes)) { if ($ch.LocalName -ne 'pPr') { [void]$p.AppendChild($ch) } }
                        if (-not $script:Absorbed.ContainsKey($p)) { $script:Absorbed[$p] = New-Object System.Collections.Generic.List[object] }
                        $script:Absorbed[$p].Add($nx)
                        $nx2 = $nx.NextSibling; [void]$nx.ParentNode.RemoveChild($nx); $nx = $nx2
                    }
                }
            }
            if ($nItems) { $report.Add("Lists: $nItems paragraph(s) with numbers or letters typed by hand made a real list (same numbers shown)") }
            if ($script:Absorbed.Count) { $report.Add("Lists: " + $script:Absorbed.Count + " list item(s) that ran on into the next paragraph joined into one item (no text changed)") }
        }

        # ---- 3e. lists with no bullets or numbers at all, right under a heading such as "Key Terms" or "Learning
        #      Objectives" (one term / objective per line): they become a real list. Terms get bullets, objectives
        #      get numbers - this adds a marker in front of each line, the text stays the same.
        $unmarked = 0; $headRx = '(?i)(key\s*terms|terms|vocabulary|glossary|abbreviations|objectives|outcomes|goals|competencies|materials|resources|equipment|supplies)'
        $live2 = @($paras | Where-Object { $_.ParentNode })
        for ($ix = 0; $ix -lt $live2.Count; $ix++) {
            $h = $live2[$ix]
            $hs = & $pStyleOf $h
            if (-not $info.ContainsKey($hs) -or -not (Test-BuiltinHeading $info[$hs].name)) { continue }
            $ht = (Get-ParaText $h $nsm).Trim(); if ($ht -notmatch $headRx) { continue }
            $isObj = $ht -match '(?i)(objectives|outcomes|goals|competencies)'
            $run = New-Object System.Collections.Generic.List[object]; $j = $ix + 1; $rs = $null
            while ($j -lt $live2.Count) {
                $q = $live2[$j]; $t = (Get-ParaText $q $nsm).Trim()
                if (-not $t) { if ($run.Count) { break } else { $j++; continue } }
                if ((& $isHeadPara $q) -or (& $isListPara $q) -or (& $inBox $q)) { break }
                $qs = & $pStyleOf $q
                if ($rs -and $qs -ne $rs) { break }
                if (-not $isObj -and ($t.Length -gt 80 -or $t -match '[\.;]$')) { break }
                if ($t.Length -gt 400) { break }
                $rs = $qs; $run.Add($q); $j++
            }
            if ($run.Count -lt 3) { continue }
            if (-not $num) { $num = New-XmlDoc ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="' + $NS_W + '"/>'); $nnsm = New-Ns $num; $script:NewNumbering = $true }
            $uAbsId = 1 + (@($num.SelectNodes('/w:numbering/w:abstractNum', $nnsm) | ForEach-Object { [int](WAttr $_ 'abstractNumId') }) + 0 | Measure-Object -Maximum).Maximum
            $uNumId = 1 + (@($num.SelectNodes('/w:numbering/w:num', $nnsm) | ForEach-Object { [int](WAttr $_ 'numId') }) + 0 | Measure-Object -Maximum).Maximum
            $an = $num.CreateElement('w', 'abstractNum', $NS_W); Set-WAttr $an 'abstractNumId' ([string]$uAbsId)
            $mt = New-W $num 'multiLevelType'; Set-WAttr $mt 'val' 'singleLevel'; [void]$an.AppendChild($mt)
            $lv = New-W $num 'lvl'; Set-WAttr $lv 'ilvl' '0'
            $fmt = if ($isObj) { @('decimal', '%1.') } else { @('bullet', [string][char]0x2022) }
            foreach ($pair in @(@('start', '1'), @('numFmt', $fmt[0]), @('lvlText', $fmt[1]), @('lvlJc', 'left'))) { $x = New-W $num $pair[0]; Set-WAttr $x 'val' $pair[1]; [void]$lv.AppendChild($x) }
            $lp = New-W $num 'pPr'; $ind = New-W $num 'ind'; Set-WAttr $ind 'left' '720'; Set-WAttr $ind 'hanging' '360'; [void]$lp.AppendChild($ind); [void]$lv.AppendChild($lp)
            [void]$an.AppendChild($lv)
            $firstNum = $num.SelectSingleNode('/w:numbering/w:num', $nnsm)
            if ($firstNum) { [void]$num.DocumentElement.InsertBefore($an, $firstNum) } else { [void]$num.DocumentElement.AppendChild($an) }
            $nn = $num.CreateElement('w', 'num', $NS_W); Set-WAttr $nn 'numId' ([string]$uNumId)
            $ai = New-W $num 'abstractNumId'; Set-WAttr $ai 'val' ([string]$uAbsId); [void]$nn.AppendChild($ai)
            $cleanup = $num.SelectSingleNode('/w:numbering/w:numIdMacAtCleanup', $nnsm)
            if ($cleanup) { [void]$num.DocumentElement.InsertBefore($nn, $cleanup) } else { [void]$num.DocumentElement.AppendChild($nn) }
            foreach ($q in $run) {
                $pPr = $q.SelectSingleNode('w:pPr', $nsm); if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$q.PrependChild($pPr) }
                $pn = New-W $doc 'numPr'; $il = New-W $doc 'ilvl'; Set-WAttr $il 'val' '0'; [void]$pn.AppendChild($il)
                $ni = New-W $doc 'numId'; Set-WAttr $ni 'val' ([string]$uNumId); [void]$pn.AppendChild($ni)
                foreach ($x in @($pPr.SelectNodes('w:numPr|w:ind', $nsm))) { [void]$pPr.RemoveChild($x) }
                [void]$pPr.AppendChild($pn); Sort-Kids $pPr $PPR_ORDER
                $script:MarkerAdded[$q] = $true; $unmarked++
            }
            $report.Add("Lists: " + $run.Count + " line(s) under '" + $ht + "' made a " + $(if ($isObj) { 'numbered' } else { 'bulleted' }) + " list (a marker is added in front of each line)")
        }

        # ---- checks the tool reports but does not change
        $prevH = 0
        foreach ($p in $paras) {
            $sid = & $pStyleOf $p
            if (-not $info.ContainsKey($sid) -or -not (Test-BuiltinHeading $info[$sid].name)) { continue }
            $txt = (Get-ParaText $p $nsm).Trim(); if (-not $txt) { continue }
            $h = [int]($info[$sid].name -replace '\D', '')
            if ($h -gt $prevH + 1) { $report.Add("CHECK: heading '" + $txt + "' is Heading $h but the heading before it is " + $(if ($prevH) { "Heading $prevH" } else { 'none' }) + " (a heading level is skipped)") }
            $prevH = $h
        }
        $fake = New-Object System.Collections.Generic.List[object]; $run = New-Object System.Collections.Generic.List[object]
        foreach ($p in $paras) {
            if (-not $p.ParentNode) { continue }
            $txt = Get-ParaText $p $nsm
            $hasNum = $p.SelectSingleNode('w:pPr/w:numPr/w:numId[@w:val!="0"]', $nsm) -or (Get-StyleChainValue $info (& $pStyleOf $p) 'numPr')
            if (-not $hasNum -and $txt -match '^\s*([\u2022\u25CF\u25AA\u25A0\u2013\u2014\-\*\u00B7]|\d{1,2}[\.\)]|[a-z][\.\)])[\s\t]') { $run.Add($p) }
            elseif ($txt.Trim()) { if ($run.Count -ge 2) { $fake.Add($run.ToArray()) }; $run = New-Object System.Collections.Generic.List[object] }
        }
        if ($run.Count -ge 2) { $fake.Add($run.ToArray()) }
        foreach ($f in $fake) { $report.Add("CHECK: " + $f.Count + " paragraphs starting '" + ((Get-ParaText $f[0] $nsm).Trim() -replace '^(.{50}).+$', '$1...') + "' look like a list typed by hand (bullets or numbers typed as text, not a real Word list)") }

        # ---- 4. document title
        if ($titleText) {
            $coreText = Read-Entry $zip 'docProps/core.xml'
            if ($coreText) {
                $core = New-XmlDoc $coreText; $cnsm = New-Ns $core
                $t = $core.SelectSingleNode('//dc:title', $cnsm)
                if (-not $t) { $t = $core.CreateElement('dc', 'title', $NS_DC); [void]$core.DocumentElement.PrependChild($t) }
                if ($t.InnerText -ne $titleText) {
                    $t.InnerText = $titleText
                    Write-Entry $zip 'docProps/core.xml' $core.OuterXml
                    $report.Add("Document title set to '" + $titleText + "'")
                }
            }
        }

        # ---- 5. running heads: STYLEREF fields that point to a style that is no longer used become plain text
        if ($refText.Count -gt 0) {
            foreach ($ent in @($zip.Entries | Where-Object { $_.FullName -match '^word/(header|footer)\d*\.xml$' } | ForEach-Object { $_.FullName })) {
                $hd = New-XmlDoc (Read-Entry $zip $ent); $hn = New-Ns $hd; $changed = $false
                foreach ($hp in @($hd.SelectNodes('//w:p', $hn))) {
                    $runs = @($hp.SelectNodes('.//w:r', $hn))
                    $i = 0
                    while ($i -lt $runs.Count) {
                        $fc = $runs[$i].SelectSingleNode('w:fldChar', $hn)
                        if (-not $fc -or (WAttr $fc 'fldCharType') -ne 'begin') { $i++; continue }
                        $instr = ''; $j = $i + 1; $sep = -1; $end = -1
                        while ($j -lt $runs.Count) {
                            $f2 = $runs[$j].SelectSingleNode('w:fldChar', $hn)
                            if ($f2 -and (WAttr $f2 'fldCharType') -eq 'separate') { $sep = $j }
                            elseif ($f2 -and (WAttr $f2 'fldCharType') -eq 'end') { $end = $j; break }
                            elseif ($sep -lt 0) { $it = $runs[$j].SelectSingleNode('w:instrText', $hn); if ($it) { $instr += $it.InnerText } }
                            $j++
                        }
                        if ($end -lt 0) { break }
                        $mm = [regex]::Match($instr, '^\s*STYLEREF\s+"?([^"\\]+?)"?\s*(\\|$)')
                        if ($mm.Success -and $refText.ContainsKey($mm.Groups[1].Value.Trim())) {
                            $txt = $refText[$mm.Groups[1].Value.Trim()]
                            $result = @(); if ($sep -ge 0) { $result = @($runs[($sep + 1)..($end - 1)] | Where-Object { $_ -and $_.SelectSingleNode('w:t', $hn) }) }
                            if ($result.Count -eq 0) {
                                $nr = $runs[$i].CloneNode($true); foreach ($x in @($nr.SelectNodes('w:fldChar', $hn))) { [void]$nr.RemoveChild($x) }
                                $t = New-W $hd 't'; [void]$t.SetAttribute('xml:space', 'preserve'); $t.InnerText = $txt; [void]$nr.AppendChild($t)
                                [void]$runs[$i].ParentNode.InsertBefore($nr, $runs[$i]); $result = @($nr)
                            } else {
                                $first = $true
                                foreach ($rr in $result) { $t = $rr.SelectSingleNode('w:t', $hn); if ($first) { $t.InnerText = $txt; [void]$t.SetAttribute('xml:space', 'preserve'); $first = $false } else { $t.InnerText = '' } }
                            }
                            for ($k = $i; $k -le $end; $k++) { if ($result -notcontains $runs[$k] -and $runs[$k].ParentNode) { [void]$runs[$k].ParentNode.RemoveChild($runs[$k]) } }
                            $changed = $true
                            $report.Add("Running head: the reference to style '" + $mm.Groups[1].Value.Trim() + "' replaced by its text '" + $txt + "'")
                        }
                        $i = $end + 1
                    }
                }
                if ($changed) { Write-Entry $zip $ent $hd.OuterXml }
            }
        }

        # ---- safety check: the text must be exactly the same (only typed list numbers may have become real
        #      numbering); any list number / bullet the reader sees that changed is listed to check
        $absorbedSet = @{}; foreach ($h in $script:Absorbed.Keys) { foreach ($x in $script:Absorbed[$h]) { $absorbedSet[$x] = $true } }
        foreach ($p in $paras) {
            if ($absorbedSet.ContainsKey($p)) { continue }
            $before = $textBefore[$p]; $after = Get-ParaText $p $nsm
            $expected = $before
            if ($script:TypedRemoved.ContainsKey($p)) { $expected = $before.Substring($script:TypedRemoved[$p].Length) }
            if ($script:Absorbed.ContainsKey($p)) { foreach ($x in $script:Absorbed[$p]) { $expected += ' ' + $textBefore[$x] } }
            if ($after -ne $expected) { throw ("Safety check failed - the text of a paragraph would change ('" + $before.Trim() + "'). Nothing was saved.") }
        }
        $live = @($paras | Where-Object { -not $absorbedSet.ContainsKey($_) })
        $labelsAfter = Get-ListLabels $live $nsm $sty $snsm $num $nnsm $defStyle
        $changedLabels = @()
        foreach ($p in $live) {
            if ($script:MarkerAdded.ContainsKey($p)) { continue }
            $bt = $textBefore[$p]; if ($script:Absorbed.ContainsKey($p)) { foreach ($x in $script:Absorbed[$p]) { $bt += $textBefore[$x] } }
            $vb = (([string]$labelsBefore[$p]) + $bt) -replace '\s', ''
            $va = (([string]$labelsAfter[$p]) + (Get-ParaText $p $nsm)) -replace '\s', ''
            if ($vb -ne $va) { $changedLabels += ("'" + ((Get-ParaText $p $nsm).Trim() -replace '^(.{60}).+$', '$1...') + "': shown as '" + ([string]$labelsBefore[$p] + ' ' + $(if ($script:TypedRemoved.ContainsKey($p)) { $script:TypedRemoved[$p].Trim() } else { '' })).Trim() + "' before, '" + ([string]$labelsAfter[$p]).Trim() + "' now") }
        }
        if ($changedLabels.Count -eq 0) { $report.Add('Safety check: all text is unchanged, and every list number and bullet looks exactly as before.') }
        else {
            $report.Add('Safety check: all text is unchanged. ' + $changedLabels.Count + ' list number(s) / bullet(s) look different now:')
            foreach ($c in ($changedLabels | Select-Object -First 40)) { $report.Add('CHECK: ' + $c) }
            if ($changedLabels.Count -gt 40) { $report.Add('CHECK: ... and ' + ($changedLabels.Count - 40) + ' more') }
        }

        Write-Entry $zip 'word/document.xml' $doc.OuterXml
        Write-Entry $zip 'word/styles.xml' $sty.OuterXml
        if ($num) { Write-Entry $zip 'word/numbering.xml' $num.OuterXml }
        if ($script:NewNumbering) {
            # a document that had no lists at all: register the new numbering part
            $ct = New-XmlDoc (Read-Entry $zip '[Content_Types].xml')
            if (-not ($ct.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Override' -and $_.GetAttribute('PartName') -eq '/word/numbering.xml' })) {
                $o = $ct.CreateElement('Override', $ct.DocumentElement.NamespaceURI); [void]$o.SetAttribute('PartName', '/word/numbering.xml'); [void]$o.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'); [void]$ct.DocumentElement.AppendChild($o)
                Write-Entry $zip '[Content_Types].xml' $ct.OuterXml
            }
            $rl = New-XmlDoc (Read-Entry $zip 'word/_rels/document.xml.rels')
            if (-not ($rl.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Relationship' -and $_.GetAttribute('Type') -match '/numbering$' })) {
                $ids = @($rl.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Relationship' } | ForEach-Object { $_.GetAttribute('Id') })
                $k = 1; while ($ids -contains ('rId' + $k)) { $k++ }
                $r0 = $rl.CreateElement('Relationship', $rl.DocumentElement.NamespaceURI); [void]$r0.SetAttribute('Id', 'rId' + $k); [void]$r0.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'); [void]$r0.SetAttribute('Target', 'numbering.xml'); [void]$rl.DocumentElement.AppendChild($r0)
                Write-Entry $zip 'word/_rels/document.xml.rels' $rl.OuterXml
            }
        }
    } finally { $zip.Dispose() }
    $chk = [IO.Compression.ZipFile]::OpenRead($outPath); $chk.Dispose()
    if ($report.Count -eq 0) { $report.Add('Nothing needed fixing in this file.') }
    return ,$report
}

# ----------------------------------------------------------------- main
Write-Host ('Word Fixer - running version ' + $WF_VERSION) -ForegroundColor Cyan
Write-Host ''
try {
    $files = @()
    if ($env:WF_FILES) { $files = @($env:WF_FILES.Split(';') | Where-Object { $_ }) }
    elseif ($args.Count -gt 0) { $files = @($args) }
    else {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $owner = New-Object System.Windows.Forms.Form; $owner.TopMost = $true
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select the Word file(s) to fix'; $dlg.Filter = 'Word documents (*.docx)|*.docx'; $dlg.Multiselect = $true
        # a tiny invisible window in front of everything, so the file window opens on top (not behind other windows)
        $owner.ShowInTaskbar = $false; $owner.StartPosition = 'Manual'; $owner.Location = New-Object Drawing.Point(-2000, -2000); $owner.Size = New-Object Drawing.Size(1, 1)
        $owner.Show(); $owner.Activate()
        if ($dlg.ShowDialog($owner) -eq 'OK') { $files = $dlg.FileNames }
        $owner.Close()
    }
    if ($files.Count -eq 0) { Write-Host 'No file selected. (If no file window appeared, it may have opened behind other windows - run Word Fixer again.)' }
    foreach ($f in $files) {
        $f = [IO.Path]::GetFullPath($f)
        Write-Host ('Fixing: ' + [IO.Path]::GetFileName($f))
        try {
            $out = [IO.Path]::Combine([IO.Path]::GetDirectoryName($f), [IO.Path]::GetFileNameWithoutExtension($f) + ' Fixed.docx')
            $rep = Repair-Docx $f $out
            foreach ($l in $rep) { if ($l.StartsWith('CHECK')) { Write-Host ('  ' + $l) -ForegroundColor Yellow } else { Write-Host ('  ' + $l) } }
            $repPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($f), [IO.Path]::GetFileNameWithoutExtension($f) + ' Fixed report.txt')
            $rep | Set-Content -LiteralPath $repPath -Encoding UTF8
            Write-Host ('  Saved: ' + $out) -ForegroundColor Green
        } catch {
            Write-Host ('  FAILED: ' + $_.Exception.Message) -ForegroundColor Red
            Write-Host ('  (version ' + $WF_VERSION + ', at line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + ($_.InvocationInfo.Line + '').Trim() + ')') -ForegroundColor DarkGray
        }
        Write-Host ''
    }
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
}
