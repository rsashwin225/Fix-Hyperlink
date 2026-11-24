# SLID Macro + Language fix + Title metadata - all in one, no PowerPoint macro needed.
#
# For every .pptx file you choose, this:
#   1. Numbers repeated slide titles: (1 of 4), (2 of 4), (3 of 4), (4 of 4).
#      Titles that appear only once, or that already end in "(n of N)", are left alone.
#   2. Renames shapes on every slide: Title, TextBox 1, TextBox 2 ... and
#      Picture (one picture) or Picture 1, Picture 2 ... On slide 1, a single
#      lone text shape is named Subtitle instead of TextBox.
#   3. Changes any language tag to English (US) (English UK, English India,
#      French, anything else). Text with no language set is left alone.
#   5. Turns the loose caption / credit text boxes under a figure into real
#      layout placeholders (adds two text placeholders to the slide layout;
#      the shapes keep their names, e.g. TextBox 1), so the accessibility checker no longer flags them.
#   4. Sets the file's Title (File > Info > Title) to match slide 1's Title
#      text combined with its Subtitle text, after step 1 has run.
#
# A backup copy is saved first, next to the original (name_BEFORE_SLIDfix.pptx).
$IMBK_VERSION = '4.4'   # shown at the start of the window; the updater compares this number with GitHub
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


$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------- namespaces ----------
$NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
$NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
$NS_DC = 'http://purl.org/dc/elements/1.1/'

function Select-Files {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select the PowerPoint file(s) to fix"
    $dlg.Filter = "PowerPoint files (*.pptx;*.pptm)|*.pptx;*.pptm"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return @() }
    return ,$dlg.FileNames
}

function Backup-Name([string]$path, [string]$suffix) {
    $dir = Split-Path $path; $base = [IO.Path]::GetFileNameWithoutExtension($path); $ext = [IO.Path]::GetExtension($path)
    $n = 0; $cand = Join-Path $dir ($base + $suffix + $ext)
    while (Test-Path $cand) { $n++; $cand = Join-Path $dir ($base + $suffix + $n + $ext) }
    return $cand
}

function New-XmlDoc([string]$text) {
    $d = New-Object Xml.XmlDocument
    $d.PreserveWhitespace = $true
    [void]$d.LoadXml($text)
    return $d
}

function New-Nsmgr([Xml.XmlDocument]$doc) {
    $nsm = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    [void]$nsm.AddNamespace('p', $NS_P); [void]$nsm.AddNamespace('a', $NS_A)
    [void]$nsm.AddNamespace('r', $NS_R); [void]$nsm.AddNamespace('rel', $NS_REL)
    return ,$nsm
}

function Read-Entry($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { return $null }
    $s = $e.Open(); $m = New-Object IO.MemoryStream; [void]$s.CopyTo($m); [void]$s.Dispose()
    return [Text.Encoding]::UTF8.GetString($m.ToArray())
}

function Write-Entry($zip, [string]$name, [string]$text) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { $e = $zip.CreateEntry($name) }
    $s = $e.Open(); [void]$s.SetLength(0)
    $bytes = [Text.Encoding]::UTF8.GetBytes($text)
    [void]$s.Write($bytes, 0, $bytes.Length); [void]$s.Dispose()
}

function Get-Rels($zip, [string]$partName) {
    $dir = Split-Path $partName -Parent
    $base = Split-Path $partName -Leaf
    $relPath = if ($dir) { "$dir/_rels/$base.rels" } else { "_rels/$base.rels" }
    $relPath = $relPath -replace '\\','/'
    $text = Read-Entry $zip $relPath
    $map = @{}
    if ($null -eq $text) { return $map }
    $doc = New-XmlDoc $text
    foreach ($rel in $doc.DocumentElement.ChildNodes) {
        $map[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
    }
    return $map
}

function Resolve-Part([string]$fromPart, [string]$target) {
    $dir = Split-Path $fromPart -Parent
    $combined = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "$dir/$target" }
    $parts = @()
    foreach ($seg in ($combined -replace '\\','/').Split('/')) {
        if ($seg -eq '..') { if ($parts.Count -gt 0) { $parts = $parts[0..($parts.Count-2)] } }
        elseif ($seg -eq '.' -or $seg -eq '') { continue }
        else { $parts += $seg }
    }
    return ($parts -join '/')
}

function Get-SlideOrder($zip) {
    $presText = Read-Entry $zip 'ppt/presentation.xml'
    $doc = New-XmlDoc $presText; $nsm = New-Nsmgr $doc
    $rels = Get-Rels $zip 'ppt/presentation.xml'
    $order = @()
    foreach ($sldId in $doc.SelectNodes('//p:sldIdLst/p:sldId', $nsm)) {
        $rid = $sldId.GetAttribute('id', $NS_R)
        $order += (Resolve-Part 'ppt/presentation.xml' $rels[$rid])
    }
    return ,$order
}

function Get-PhInfo([Xml.XmlElement]$sp, $nsm) {
    $ph = $sp.SelectSingleNode('.//p:nvPr/p:ph', $nsm)
    if ($null -eq $ph) { return $null }
    return @{ type = $ph.GetAttribute('type'); idx = $ph.GetAttribute('idx') }
}

function Get-Xfrm([Xml.XmlElement]$elem, $nsm, [string]$tag) {
    $xfrm = $elem.SelectSingleNode("$tag/a:xfrm", $nsm)
    if ($null -eq $xfrm) { return $null }
    $off = $xfrm.SelectSingleNode('a:off', $nsm)
    if ($null -eq $off) { return $null }
    return @{ x = [long]$off.GetAttribute('x'); y = [long]$off.GetAttribute('y') }
}

function Find-PhShape([Xml.XmlDocument]$doc, $nsm, [string]$type, [string]$idx) {
    foreach ($sp in $doc.SelectNodes('//p:sp', $nsm)) {
        $info = Get-PhInfo $sp $nsm
        if ($null -eq $info) { continue }
        if ($idx -and $info.idx -eq $idx) { return $sp }
        if ((-not $idx) -and $type -and $info.type -eq $type) { return $sp }
    }
    return $null
}

function Get-LayoutMasterNames($zip, [string]$slideName) {
    $rels = Get-Rels $zip $slideName
    $layoutTarget = ($rels.Values | Where-Object { $_ -match 'slideLayout' } | Select-Object -First 1)
    $layoutName = Resolve-Part $slideName $layoutTarget
    $lrels = Get-Rels $zip $layoutName
    $masterTarget = ($lrels.Values | Where-Object { $_ -match 'slideMaster' } | Select-Object -First 1)
    $masterName = Resolve-Part $layoutName $masterTarget
    return @($layoutName, $masterName)
}

function Resolve-Position($zip, [string]$slideName, [Xml.XmlElement]$elem, $nsm, [string]$tag) {
    if ($tag) {
        $pos = Get-Xfrm $elem $nsm $tag
        if ($pos) { return $pos }
    }
    $info = Get-PhInfo $elem $nsm
    if ($null -eq $info) { return $null }
    $names = Get-LayoutMasterNames $zip $slideName
    $layoutText = Read-Entry $zip $names[0]
    if ($null -eq $layoutText) { return $null }
    $ldoc = New-XmlDoc $layoutText
    $lsp = Find-PhShape $ldoc $nsm $info.type $info.idx
    if ($lsp) {
        $pos = Get-Xfrm $lsp $nsm 'p:spPr'
        if ($pos) { return $pos }
        $linfo = Get-PhInfo $lsp $nsm
        $masterText = Read-Entry $zip $names[1]
        if ($masterText) {
            $mdoc = New-XmlDoc $masterText
            $wantType = if ($linfo) { $linfo.type } else { $info.type }
            $msp = Find-PhShape $mdoc $nsm $wantType $null
            if ($msp) { return (Get-Xfrm $msp $nsm 'p:spPr') }
        }
    }
    return $null
}

function Resolve-Extent($zip, [string]$slideName, [Xml.XmlElement]$elem, $nsm, [string]$tag) {
    $ext = $elem.SelectSingleNode("$tag/a:xfrm/a:ext", $nsm)
    if ($ext) { return @{ cx = [long]$ext.GetAttribute('cx'); cy = [long]$ext.GetAttribute('cy') } }
    $info = Get-PhInfo $elem $nsm
    if ($null -eq $info) { return $null }
    $names = Get-LayoutMasterNames $zip $slideName
    $layoutText = Read-Entry $zip $names[0]
    if ($null -eq $layoutText) { return $null }
    $ldoc = New-XmlDoc $layoutText
    $lsp = Find-PhShape $ldoc $nsm $info.type $info.idx
    if ($lsp) {
        $lext = $lsp.SelectSingleNode("p:spPr/a:xfrm/a:ext", $nsm)
        if ($lext) { return @{ cx = [long]$lext.GetAttribute('cx'); cy = [long]$lext.GetAttribute('cy') } }
        $linfo = Get-PhInfo $lsp $nsm
        $masterText = Read-Entry $zip $names[1]
        if ($masterText) {
            $mdoc = New-XmlDoc $masterText
            $wantType = if ($linfo) { $linfo.type } else { $info.type }
            $msp = Find-PhShape $mdoc $nsm $wantType $null
            if ($msp) {
                $mext = $msp.SelectSingleNode("p:spPr/a:xfrm/a:ext", $nsm)
                if ($mext) { return @{ cx = [long]$mext.GetAttribute('cx'); cy = [long]$mext.GetAttribute('cy') } }
            }
        }
    }
    return $null
}

function Ensure-ExplicitXfrm($zip, [string]$slideName, [Xml.XmlElement]$elem, [string]$tag, $nsm) {
    $existing = $elem.SelectSingleNode("$tag/a:xfrm", $nsm)
    if ($existing) { return }
    $pos = Resolve-Position $zip $slideName $elem $nsm $tag
    $ext = Resolve-Extent $zip $slideName $elem $nsm $tag
    if ($null -eq $pos) { $pos = @{ x = 0; y = 0 } }
    if ($null -eq $ext) { $ext = @{ cx = 914400; cy = 914400 } }
    $spPr = $elem.SelectSingleNode($tag, $nsm)
    $doc = $elem.OwnerDocument
    $xfrm = $doc.CreateElement('a', 'xfrm', $NS_A)
    $off = $doc.CreateElement('a', 'off', $NS_A)
    [void]$off.SetAttribute('x', [string]$pos.x); [void]$off.SetAttribute('y', [string]$pos.y)
    $extEl = $doc.CreateElement('a', 'ext', $NS_A)
    [void]$extEl.SetAttribute('cx', [string]$ext.cx); [void]$extEl.SetAttribute('cy', [string]$ext.cy)
    [void]$xfrm.AppendChild($off); [void]$xfrm.AppendChild($extEl)
    [void]$spPr.PrependChild($xfrm)
}

function Test-IsTitlePh($sp, $nsm) {
    $info = Get-PhInfo $sp $nsm
    return ($null -ne $info) -and ($info.type -eq 'title' -or $info.type -eq 'ctrTitle')
}

function Test-IsPicture([Xml.XmlElement]$elem, $nsm) {
    $local = $elem.LocalName
    if ($local -eq 'pic') { return $true }
    if ($local -eq 'grpSp') {
        foreach ($child in $elem.ChildNodes) {
            if ($child.LocalName -in @('sp','pic','grpSp','graphicFrame')) {
                if (Test-IsPicture $child $nsm) { return $true }
            }
        }
        return $false
    }
    if ($local -eq 'sp') {
        $info = Get-PhInfo $elem $nsm
        if ($info -and $info.type -eq 'pic') { return $true }
    }
    return $false
}

function Test-HasText([Xml.XmlElement]$sp, $nsm) {
    foreach ($t in $sp.SelectNodes('.//a:t', $nsm)) {
        if ($t.InnerText -and $t.InnerText.Trim().Length -gt 0) { return $true }
    }
    return $false
}

function Get-TitleText([Xml.XmlDocument]$doc, $nsm) {
    $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
    foreach ($elem in $spTree.ChildNodes) {
        if ($elem.LocalName -eq 'sp' -and (Test-IsTitlePh $elem $nsm)) {
            return @{ shape = $elem; text = (Get-SlideText $elem $nsm) }
        }
    }
    return $null
}

function Get-ShapeTextByName([Xml.XmlDocument]$doc, $nsm, [string]$shapeName) {
    $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
    foreach ($elem in $spTree.ChildNodes) {
        if ($elem.LocalName -ne 'sp') { continue }
        $cNvPr = $elem.SelectSingleNode('.//p:cNvPr', $nsm)
        if ($cNvPr -and $cNvPr.GetAttribute('name') -eq $shapeName) {
            return (Get-SlideText $elem $nsm)
        }
    }
    return $null
}

function Get-SlidKey([string]$txt) {
    $t = $txt -replace "`r"," " -replace [char]0x0B," " -replace [char]0xA0," "
    while ($t -match '  ') { $t = $t -replace '  ',' ' }
    return $t.Trim().ToLowerInvariant()
}

function Test-IsNumbered([string]$txt) {
    return [regex]::IsMatch($txt.Trim(), '\(\d{1,2} of \d{1,2}\)$')
}

# ---------- Step 1: number repeated titles ----------
function Invoke-TitleDedup($zip, [string[]]$order) {
    $docs = @{}; $titleInfo = @{}
    foreach ($name in $order) {
        $text = Read-Entry $zip $name
        $doc = New-XmlDoc $text; $nsm = New-Nsmgr $doc
        $docs[$name] = @{ doc = $doc; nsm = $nsm }
        $ti = Get-TitleText $doc $nsm
        $titleInfo[$name] = $ti
    }
    $counts = @{}
    foreach ($name in $order) {
        $ti = $titleInfo[$name]
        if ($null -eq $ti -or $ti.text.Trim().Length -eq 0) { continue }
        $k = Get-SlidKey $ti.text
        if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 }
    }
    $seen = @{}; $changed = 0
    foreach ($name in $order) {
        $ti = $titleInfo[$name]
        if ($null -eq $ti -or $ti.text.Trim().Length -eq 0) { continue }
        $k = Get-SlidKey $ti.text
        if ($counts[$k] -le 1) { continue }
        if (Test-IsNumbered $ti.text) { continue }
        if ($seen.ContainsKey($k)) { $seen[$k]++ } else { $seen[$k] = 1 }
        $suffix = "(" + $seen[$k] + " of " + $counts[$k] + ")"
        if (-not $ti.text.EndsWith(' ')) { $suffix = ' ' + $suffix }
        $nsm = $docs[$name].nsm
        $lastPara = $ti.shape.SelectNodes('.//a:p', $nsm) | Select-Object -Last 1
        $lastRun = $lastPara.SelectNodes('.//a:r', $nsm) | Select-Object -Last 1
        if ($null -eq $lastRun) { continue }
        $lastT = $lastRun.SelectSingleNode('a:t', $nsm)
        $lastT.InnerText = $lastT.InnerText + $suffix
        $changed++
    }
    foreach ($name in $order) { Write-Entry $zip $name $docs[$name].doc.OuterXml }
    return @{ changed = $changed; docs = $docs }
}

# ---------- Step 2: rename shapes on every slide ----------
function Invoke-RenameShapes($zip, [string[]]$order, $docs, $reorderSet) {
    $firstSlide = $order[0]
    foreach ($name in $order) {
        $doc = $docs[$name].doc; $nsm = $docs[$name].nsm
        $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
        $titles = @(); $pics = @(); $texts = @()
        foreach ($elem in @($spTree.ChildNodes)) {
            $local = $elem.LocalName
            if ($local -notin @('sp','pic','grpSp','graphicFrame','cxnSp')) { continue }
            if ($local -eq 'sp') { $phx = Get-PhInfo $elem $nsm; if ($phx -and $phx.type -eq 'body' -and ($phx.idx -eq '13' -or $phx.idx -eq '14')) { continue } }
            $tag = switch ($local) { 'sp'{'p:spPr'} 'pic'{'p:spPr'} 'grpSp'{'p:grpSpPr'} 'cxnSp'{'p:spPr'} default{$null} }
            $pos = if ($tag) { Resolve-Position $zip $name $elem $nsm $tag } elseif ($local -eq 'graphicFrame') { Get-Xfrm $elem $nsm 'p:xfrm' } else { $null }
            $cNvPr = $elem.SelectSingleNode('.//p:cNvPr', $nsm)
            $entry = @{ elem = $elem; cNvPr = $cNvPr; pos = $pos }
            if ($local -eq 'sp' -and (Test-IsTitlePh $elem $nsm)) { $titles += $entry }
            elseif (Test-IsPicture $elem $nsm) { $pics += $entry }
            elseif ($local -eq 'sp' -and (Test-HasText $elem $nsm)) { $texts += $entry }
        }
        $sortKey = { if ($_.pos) { [double]$_.pos.y * 1000000000 + [double]$_.pos.x } else { [double]::MaxValue } }
        $titles = @($titles | Sort-Object -Property @{Expression=$sortKey})
        $pics   = @($pics   | Sort-Object -Property @{Expression=$sortKey})
        $texts  = @($texts  | Sort-Object -Property @{Expression=$sortKey})
        for ($i=0; $i -lt $titles.Count; $i++) { [void]$titles[$i].cNvPr.SetAttribute('name', $(if ($titles.Count -eq 1) {'Title'} else {'Title ' + ($i+1)})) }
        for ($i=0; $i -lt $texts.Count; $i++)  {
            $label = if ($texts.Count -eq 1) { if ($name -eq $firstSlide) { 'Subtitle' } else { 'TextBox' } } else { 'TextBox ' + ($i+1) }
            [void]$texts[$i].cNvPr.SetAttribute('name', $label)
        }
        for ($i=0; $i -lt $pics.Count; $i++)   { [void]$pics[$i].cNvPr.SetAttribute('name',   $(if ($pics.Count   -eq 1) {'Picture'} else {'Picture ' + ($i+1)})) }

        # Reading / tab order = back to front: title, picture, textbox(es) -
        # only for slides where a figure was actually laid out this run,
        # matching the reference macro's own behaviour. Moving a shape to
        # the end of spTree is exactly what PowerPoint's "bring to front"
        # does, and spTree order is also reading/tab order.
        if ($reorderSet -and $reorderSet.ContainsKey($name)) {
            foreach ($t in $titles) { [void]$spTree.AppendChild($t.elem) }
            foreach ($p in $pics)   { [void]$spTree.AppendChild($p.elem) }
            foreach ($x in $texts)  { [void]$spTree.AppendChild($x.elem) }
        }

        Write-Entry $zip $name $doc.OuterXml
    }
}

# ---------- Step 3: language fix (any language -> English US, except "no language") ----------
function Invoke-LanguageFix($zip) {
    $L = [Text.Encoding]::GetEncoding(28591)
    $q = [char]34
    $rx = New-Object Text.RegularExpressions.Regex('(?<=<[^<>]*?\s(?:lang|altLang)=' + $q + ')[^' + $q + ']+(?=' + $q + ')')
    $eval = { param($m) if ($m.Value -ieq 'en-US' -or $m.Value -ieq 'x-none') { $m.Value } else { 'en-US' } }
    $hit = @{}
    foreach ($e in $zip.Entries) {
        if ($e.FullName -like '*.xml') {
            $s = $e.Open(); $m = New-Object IO.MemoryStream; [void]$s.CopyTo($m); [void]$s.Dispose()
            $text = $L.GetString($m.ToArray())
            $c = @($rx.Matches($text) | Where-Object { $_.Value -ine 'en-US' -and $_.Value -ine 'x-none' }).Count
            if ($c -gt 0) { $hit[$e.FullName] = $c }
        }
    }
    foreach ($name in @($hit.Keys)) {
        $entry = $zip.GetEntry($name)
        $s = $entry.Open(); $m = New-Object IO.MemoryStream; [void]$s.CopyTo($m)
        $bytes = $L.GetBytes($rx.Replace($L.GetString($m.ToArray()), $eval))
        [void]$s.SetLength(0); [void]$s.Write($bytes, 0, $bytes.Length); [void]$s.Dispose()
    }
    $total = 0; foreach ($v in $hit.Values) { $total += $v }
    return $total
}

# ---------- Step 4: set the document Title to match slide 1's title+subtitle ----------
function Set-DocTitle($zip, [string]$titleText) {
    $corePath = 'docProps/core.xml'
    $text = Read-Entry $zip $corePath
    if ($null -eq $text) { return $false }
    $doc = New-XmlDoc $text
    $nsm = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    [void]$nsm.AddNamespace('cp', $NS_CP); [void]$nsm.AddNamespace('dc', $NS_DC)
    $titleNode = $doc.SelectSingleNode('//dc:title', $nsm)
    if ($null -eq $titleNode) {
        $titleNode = $doc.CreateElement('dc', 'title', $NS_DC)
        [void]$doc.DocumentElement.PrependChild($titleNode)
    }
    $titleNode.InnerText = $titleText
    Write-Entry $zip $corePath $doc.OuterXml
    return $true
}

# ============================================================
# IMBK FIGURE LAYOUT (ported from IMBK_Macro_Nov.bas)
#
# Idempotency guard (the bug fix): before splitting a caption, check
# whether another shape on the same slide already shows exactly the
# bare figure number. If so, this slide has already been split in an
# earlier run - skip it, instead of creating another duplicate.
# ============================================================

function Test-CaptionLabel([string]$t) {
    # "Figure 11.UNF02", "Table 11.T01", "Figure 03.PD01_UNF01" - a word,
    # then a chapter.item code (letters/digits/underscore/hyphen).
    if ($t -match '^[A-Za-z]+\s+[A-Za-z0-9_\-]*\d[A-Za-z0-9_\-]*\.[A-Za-z0-9_\-]+$') { return $true }
    # "01.F0106" - bare chapter.item code, no leading word.
    if ($t -match '^\d\d\.[A-Za-z][A-Za-z0-9_\-]*$') { return $true }
    return $false
}

function Get-FigureNumber([string]$txt) {
    $t = $txt -replace [char]0xA0,' ' -replace "`r",' ' -replace [char]0x0B,' '
    $p = $t.IndexOf(':')
    if ($p -lt 0 -or $p -gt 40) { return '' }
    $t = $t.Substring(0, $p).Trim()
    while ($t -match '  ') { $t = $t -replace '  ',' ' }
    if (Test-CaptionLabel $t) { return $t }
    return ''
}

function Get-BareCaptionLabel([string]$txt) {
    # A caption-style label with no colon and no separate description at all
    # (e.g. "Figure 11.UNF02" on its own) - nothing to split, the whole
    # shape just needs to move up and be marked decorative.
    $t = $txt -replace [char]0xA0,' ' -replace "`r",' ' -replace [char]0x0B,' '
    $t = $t.Trim()
    while ($t -match '  ') { $t = $t -replace '  ',' ' }
    if ($t.IndexOf(':') -ge 0) { return '' }
    if ($t.Length -eq 0 -or $t.Length -gt 40) { return '' }
    if (Test-CaptionLabel $t) { return $t }
    return ''
}

function Get-SlideText([Xml.XmlElement]$elem, $nsm) {
    # Each a:p is a separate paragraph (often a soft line-wrap within a
    # caption); joining them with nothing between would run the last word
    # of one into the first word of the next, so join with a space instead.
    $paraTexts = @()
    foreach ($p in $elem.SelectNodes('.//a:p', $nsm)) {
        $sb = New-Object Text.StringBuilder
        foreach ($t in $p.SelectNodes('.//a:t', $nsm)) { [void]$sb.Append($t.InnerText) }
        $paraTexts += $sb.ToString()
    }
    return ($paraTexts -join ' ')
}

function Get-ShapeArea($zip, [string]$slideName, [Xml.XmlElement]$elem, $nsm) {
    $tag = switch ($elem.LocalName) { 'sp'{'p:spPr'} 'pic'{'p:spPr'} 'grpSp'{'p:grpSpPr'} default{$null} }
    if (-not $tag) { return 0 }
    $ext = $elem.SelectSingleNode("$tag/a:xfrm/a:ext", $nsm)
    if ($null -eq $ext) { return 0 }
    return [double]$ext.GetAttribute('cx') * [double]$ext.GetAttribute('cy')
}

function Get-ShapeExt([Xml.XmlElement]$elem, [string]$tag, $nsm) {
    $ext = $elem.SelectSingleNode("$tag/a:xfrm/a:ext", $nsm)
    if ($null -eq $ext) { return $null }
    return @{ cx = [long]$ext.GetAttribute('cx'); cy = [long]$ext.GetAttribute('cy') }
}

function Set-ShapeOff([Xml.XmlElement]$elem, [string]$tag, $nsm, [long]$x, [long]$y) {
    $off = $elem.SelectSingleNode("$tag/a:xfrm/a:off", $nsm)
    $off.SetAttribute('x', [string]$x); $off.SetAttribute('y', [string]$y)
}

function Set-ShapeExt([Xml.XmlElement]$elem, [string]$tag, $nsm, [long]$cx, [long]$cy) {
    $ext = $elem.SelectSingleNode("$tag/a:xfrm/a:ext", $nsm)
    $ext.SetAttribute('cx', [string]$cx); $ext.SetAttribute('cy', [string]$cy)
}

function Get-ShapeOff([Xml.XmlElement]$elem, [string]$tag, $nsm) {
    $off = $elem.SelectSingleNode("$tag/a:xfrm/a:off", $nsm)
    if ($null -eq $off) { return $null }
    return @{ x = [long]$off.GetAttribute('x'); y = [long]$off.GetAttribute('y') }
}

# Random double in [lo, hi)
$script:Rng = New-Object Random
function Get-RandRange([double]$lo, [double]$hi) { return $lo + $script:Rng.NextDouble() * ($hi - $lo) }

function Get-TitleStyleDefaults($zip, [string]$slideName, $nsm) {
    # The original caption is a title placeholder whose run usually carries
    # no explicit size/font at all - it inherits both from the layout's (or
    # master's) title style. A brand-new, non-placeholder textbox has no
    # such inheritance, so to make it look the same, resolve and bake in
    # the same effective size/font explicitly.
    $names = Get-LayoutMasterNames $zip $slideName
    $result = @{ sz = $null; typeface = $null }
    $layoutText = Read-Entry $zip $names[0]
    if ($layoutText) {
        $ldoc = New-XmlDoc $layoutText
        $defRPr = $ldoc.SelectSingleNode('//p:sp[.//p:ph[@type="title" or @type="ctrTitle"]]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr', $nsm)
        if ($defRPr) {
            if ($defRPr.GetAttribute('sz')) { $result.sz = $defRPr.GetAttribute('sz') }
            $latin = $defRPr.SelectSingleNode('a:latin', $nsm)
            if ($latin) { $result.typeface = $latin.GetAttribute('typeface') }
        }
    }
    if ((-not $result.sz) -or (-not $result.typeface)) {
        $masterText = Read-Entry $zip $names[1]
        if ($masterText) {
            $mdoc = New-XmlDoc $masterText
            $defRPr2 = $mdoc.SelectSingleNode('//p:titleStyle/a:lvl1pPr/a:defRPr', $nsm)
            if ($defRPr2) {
                if ((-not $result.sz) -and $defRPr2.GetAttribute('sz')) { $result.sz = $defRPr2.GetAttribute('sz') }
                if (-not $result.typeface) {
                    $latin2 = $defRPr2.SelectSingleNode('a:latin', $nsm)
                    if ($latin2) { $result.typeface = $latin2.GetAttribute('typeface') }
                }
            }
        }
    }
    return $result
}

# The caption text box is made just tall enough for its text, the way PowerPoint itself sizes an
# auto-fitting box once you click into it - so it does not look squeezed when the file is opened.
function Set-CaptionBoxHeight([Xml.XmlElement]$sp, $nsm) {
    $ext = $sp.SelectSingleNode('p:spPr/a:xfrm/a:ext', $nsm)
    if ($null -eq $ext) { return $false }
    $PT = 12700.0
    $run = $sp.SelectSingleNode('.//a:r/a:rPr', $nsm)
    $sz = 18.0
    if ($run -and $run.GetAttribute('sz')) { $sz = [double]$run.GetAttribute('sz') / 100.0 }
    $body = $sp.SelectSingleNode('p:txBody/a:bodyPr', $nsm)
    $insL = 91440; $insR = 91440; $insT = 45720; $insB = 45720
    if ($body) {
        if ($body.GetAttribute('lIns')) { $insL = [double]$body.GetAttribute('lIns') }
        if ($body.GetAttribute('rIns')) { $insR = [double]$body.GetAttribute('rIns') }
        if ($body.GetAttribute('tIns')) { $insT = [double]$body.GetAttribute('tIns') }
        if ($body.GetAttribute('bIns')) { $insB = [double]$body.GetAttribute('bIns') }
    }
    $w = [double]$ext.GetAttribute('cx') - $insL - $insR
    if ($w -le 0) { return $false }
    $lines = 0
    foreach ($para in $sp.SelectNodes('p:txBody/a:p', $nsm)) {
        $txt = (($para.SelectNodes('.//a:t', $nsm) | ForEach-Object { $_.InnerText }) -join '')
        $textW = $txt.Length * $sz * 0.5 * $PT
        $lines += [Math]::Max(1, [Math]::Ceiling($textW / $w))
    }
    if ($lines -lt 1) { $lines = 1 }
    $h = [long]([Math]::Ceiling($lines * $sz * 1.2 * $PT + $insT + $insB))
    if ([Math]::Abs([double]$ext.GetAttribute('cy') - $h) -lt 6350) { return $false }
    [void]$ext.SetAttribute('cy', [string]$h)
    return $true
}

function Invoke-ImbkLayout($zip, [string]$slideName, [Xml.XmlDocument]$doc, $nsm, [double]$slideW, [double]$slideH, [bool]$hideTopTitle) {
    $PT = 12700.0  # EMU per point
    $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
    $shapes = @($spTree.ChildNodes | Where-Object { $_.LocalName -in @('sp','pic','grpSp') })

    # Find the first shape whose text matches a figure caption - either a
    # label with a colon and a description ("Table 11.T01: ...description"),
    # or a bare label on its own with no separate description at all
    # ("Figure 11.UNF02"). Either way a textbox is always created below the
    # picture: with the description when there is one, or with a second
    # copy of the same bare label when there is not - every slide that gets
    # moved up gets a readable caption below the picture as well.
    # A with-colon match (a real description) always wins over a bare
    # match (just a label, no description) if both exist on the slide -
    # scan the whole slide for a with-colon match first, and only look
    # for a bare match if nothing richer was found anywhere.
    $cap = $null; $figNum = ''
    foreach ($elem in $shapes) {
        if ($elem.LocalName -ne 'sp') { continue }
        $elemCNvPr = $elem.SelectSingleNode('.//p:cNvPr', $nsm)
        if ($elemCNvPr -and $elemCNvPr.GetAttribute('hidden') -eq '1') { continue }
        $txt = (Get-SlideText $elem $nsm).Trim()
        if ($txt.Length -lt 4) { continue }
        $fn = Get-FigureNumber $txt
        if ($fn) { $cap = $elem; $figNum = $fn; break }
    }
    if ($null -eq $cap) {
        foreach ($elem in $shapes) {
            if ($elem.LocalName -ne 'sp') { continue }
            $elemCNvPr = $elem.SelectSingleNode('.//p:cNvPr', $nsm)
            if ($elemCNvPr -and $elemCNvPr.GetAttribute('hidden') -eq '1') { continue }
            $txt = (Get-SlideText $elem $nsm).Trim()
            if ($txt.Length -lt 4) { continue }
            $bl = Get-BareCaptionLabel $txt
            if ($bl) { $cap = $elem; $figNum = $bl; break }
        }
    }
    if ($null -eq $cap) { return $null }   # nothing to do on this slide

    # Idempotency guard: has this exact figure number already been split out
    # onto its own shape elsewhere on the slide? Either this slide was already
    # processed in an earlier run, or the book's own original design already
    # authored it this way (both are fine - nothing needs to be re-split or
    # duplicated). The only thing still worth doing in that case is making
    # sure the bare-number shape has the hidden/decorative accessibility
    # marking, since a natively pre-split slide may never have had it set.
    foreach ($elem in $shapes) {
        if ($elem -eq $cap) { continue }
        if ($elem.LocalName -ne 'sp') { continue }
        $t2 = (Get-SlideText $elem $nsm).Trim()
        if ($t2 -eq $figNum) {
            # a caption box made by an earlier run: give it the right height too
            if ($cap.SelectSingleNode('p:txBody/a:bodyPr/a:spAutoFit', $nsm)) { if (Set-CaptionBoxHeight $cap $nsm) { $script:BoxesResized++ } }
            $alreadyMarked = $elem.SelectSingleNode('.//p:cNvPr[@hidden="1"]', $nsm)
            if ($hideTopTitle -and (-not $alreadyMarked)) {
                Set-ShapeDecorativeHidden $elem $nsm
                return 'marked-only'
            }
            return 'already-processed'
        }
    }

    # Largest picture on the slide. The picture itself is never moved or
    # resized - only its current, unchanged position is used to decide
    # where the new caption goes.
    $pic = $null; $bestArea = 0
    foreach ($elem in $shapes) {
        if ($elem -eq $cap) { continue }
        if (-not (Test-IsPicture $elem $nsm)) { continue }
        $a = Get-ShapeArea $zip $slideName $elem $nsm
        if ($a -gt $bestArea) { $bestArea = $a; $pic = $elem }
    }
    if ($null -eq $pic) { return 'no-picture' }
    Ensure-ExplicitXfrm $zip $slideName $pic 'p:spPr' $nsm
    $picOff = Get-ShapeOff $pic 'p:spPr' $nsm
    $picExt = Get-ShapeExt $pic 'p:spPr' $nsm

    # Top of the credit/source line (lowest text shape), if any - the new
    # caption must not be pushed down into it.
    $creditTop = $slideH - (30 * $PT)
    foreach ($elem in $shapes) {
        if ($elem -eq $cap -or $elem -eq $pic) { continue }
        if ($elem.LocalName -ne 'sp') { continue }
        if (-not (Test-HasText $elem $nsm)) { continue }
        $off = Get-ShapeOff $elem 'p:spPr' $nsm
        if ($null -eq $off) { continue }
        if ($off.y -gt ($slideH * 0.5) -and $off.y -lt $creditTop) { $creditTop = $off.y }
    }

    # Look of the original caption: first run's formatting.
    $firstRun = $cap.SelectNodes('.//a:r', $nsm) | Select-Object -First 1
    $rPr = $null
    if ($firstRun) { $rPr = $firstRun.SelectSingleNode('a:rPr', $nsm) }
    $capSizeAttr = if ($rPr) { $rPr.GetAttribute('sz') } else { '' }
    $capSize = if ($capSizeAttr) { [double]$capSizeAttr / 100.0 } else { 14.0 }
    if ($capSize -le 0) { $capSize = 14.0 }

    Ensure-ExplicitXfrm $zip $slideName $cap 'p:spPr' $nsm
    $capOff = Get-ShapeOff $cap 'p:spPr' $nsm
    $capExt = Get-ShapeExt $cap 'p:spPr' $nsm

    # Small natural-looking variation so slides are not identical.
    $TOP_START_EMU = (Get-RandRange 21 27) * $PT
    $GAP_BELOW = (Get-RandRange 6 11) * $PT

    # The new caption is centred on the picture itself - not on the title
    # placeholder's own (often much wider, off-centre) inherited geometry -
    # so a narrow picture gets a caption that visually belongs to it rather
    # than one stretching across unrelated space. But the box still needs
    # to be wide enough to hold the whole caption on one line when the
    # slide has room for it - only wrap if there genuinely isn't room.
    $picCenterX = $picOff.x + ($picExt.cx / 2.0)
    $nbPad = 36.0 * $PT
    $slideMargin = 9.0 * $PT
    $maxAvailableW = $slideW - (2 * $slideMargin)
    $estCaptionLen = (Get-SlideText $cap $nsm).Trim().Length
    $estTextW = $estCaptionLen * $capSize * 0.52 * $PT + $nbPad
    $nbW = [Math]::Max($picExt.cx + $nbPad, $estTextW)
    if ($nbW -gt $maxAvailableW) { $nbW = $maxAvailableW }
    $nbLeft = $picCenterX - ($nbW / 2.0)
    if ($nbLeft -lt $slideMargin) { $nbLeft = $slideMargin }
    if (($nbLeft + $nbW) -gt ($slideW - $slideMargin)) { $nbLeft = $slideW - $slideMargin - $nbW }
    if ($nbLeft -lt $slideMargin) { $nbLeft = $slideMargin }

    $capW = $capExt.cx - (Get-RandRange 0 60) * $PT
    $capLeft = $capOff.x + ($capExt.cx - $capW) * $script:Rng.NextDouble()

    # New text box below the picture (the picture's own, unchanged position)
    # with the full caption, copying the caption's own first run as a
    # template so font/size/color/bold/italic/lang all come across unchanged,
    # with the title's own effective size/font baked in explicitly for
    # whatever the original run did not already set itself.
    $fullText = Get-SlideText $cap $nsm
    $nbId = 900000 + (Get-Random -Minimum 1 -Maximum 89999)
    $templateRunXml = if ($firstRun) { $firstRun.OuterXml } else { '<a:r xmlns:a="' + $NS_A + '"><a:t></a:t></a:r>' }
    $nbXml = @"
<p:sp xmlns:p="$NS_P" xmlns:a="$NS_A">
  <p:nvSpPr>
    <p:cNvPr id="$nbId" name="TextBox $nbId"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="$([long]$nbLeft)" y="0"/><a:ext cx="$([long]$nbW)" cy="182880"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/>$templateRunXml</a:p>
  </p:txBody>
</p:sp>
"@
    $nbDoc = New-XmlDoc $nbXml
    $nbNode = $doc.ImportNode($nbDoc.DocumentElement, $true)
    $nbTextNode = $nbNode.SelectSingleNode('.//a:t', $nsm)
    $nbTextNode.InnerText = $fullText
    $nbRPr = $nbNode.SelectSingleNode('.//a:r/a:rPr', $nsm)
    if ($null -eq $nbRPr) {
        $nbRPr = $doc.CreateElement('a', 'rPr', $NS_A)
        [void]$nbNode.SelectSingleNode('.//a:r', $nsm).PrependChild($nbRPr)
    }
    $titleDefaults = Get-TitleStyleDefaults $zip $slideName $nsm
    if ((-not $nbRPr.GetAttribute('sz')) -and $titleDefaults.sz) { [void]$nbRPr.SetAttribute('sz', $titleDefaults.sz) }
    if ($titleDefaults.typeface -and (-not $nbRPr.SelectSingleNode('a:latin', $nsm))) {
        $latinEl = $doc.CreateElement('a', 'latin', $NS_A)
        [void]$latinEl.SetAttribute('typeface', $titleDefaults.typeface)
        [void]$nbRPr.AppendChild($latinEl)
    }
    [void]$spTree.AppendChild($nbNode)
    [void](Set-CaptionBoxHeight $nbNode $nsm)

    # Original caption frame keeps only the figure number.
    $capTextNode = $cap.SelectNodes('.//a:t', $nsm) | Select-Object -First 1
    $capTextNode.InnerText = $figNum
    $allParas = @($cap.SelectNodes('.//a:p', $nsm))
    for ($i = 1; $i -lt $allParas.Count; $i++) { [void]$allParas[$i].ParentNode.RemoveChild($allParas[$i]) }
    $firstPara = $allParas[0]
    $runsInFirst = @($firstPara.SelectNodes('.//a:r', $nsm))
    for ($i = 1; $i -lt $runsInFirst.Count; $i++) { [void]$runsInFirst[$i].ParentNode.RemoveChild($runsInFirst[$i]) }

    Set-ShapeOff $cap 'p:spPr' $nsm ([long]$capLeft) ([long]$TOP_START_EMU)
    Set-ShapeExt $cap 'p:spPr' $nsm ([long]$capW) $capExt.cy

    $capH = [double]$capExt.cy
    $estH = ($capSize * 1.22 + 2) * $PT
    if ($estH -lt $capH) { $capH = $estH }
    if ($capH -lt 20 * $PT) { $capH = 20 * $PT }
    Set-ShapeExt $cap 'p:spPr' $nsm ([long]$capW) ([long]$capH)

    # The new caption goes directly below the picture's own, unchanged
    # bottom edge - the picture itself is never touched.
    $picBottom = $picOff.y + $picExt.cy
    $nbTop = $picBottom + $GAP_BELOW
    $nbH = [double]$nbNode.SelectSingleNode('p:spPr/a:xfrm/a:ext', $nsm).GetAttribute('cy')

    $nbTop0 = $nbTop
    if (($nbTop + $nbH + 4 * $PT) -gt $creditTop) {
        # tight space: close the gap under the picture as far as needed, but never overlap it
        $tight = $creditTop - $nbH - 1 * $PT
        if ($tight -ge $picBottom) { $nbTop = $tight }
        elseif (($nbTop0 + 24 * $PT) -le $creditTop) {
            # still no room for the full height: use the space there is (as earlier versions did)
            $nbTop = $nbTop0; $nbH = $creditTop - $nbTop0 - 2 * $PT
            Set-ShapeExt $nbNode 'p:spPr' $nsm ([long]$nbW) ([long]$nbH)
        }
    }
    if (($nbTop + $nbH) -gt $creditTop) {
        # Not enough room between the picture and the credit line: put everything back.
        $capTextNode.InnerText = $fullText
        [void]$spTree.RemoveChild($nbNode)
        return 'needs-review'
    }

    Set-ShapeOff $nbNode 'p:spPr' $nsm ([long]$nbLeft) ([long]$nbTop)

    if ($hideTopTitle) { Set-ShapeDecorativeHidden $cap $nsm }

    return 'done'
}

function Set-ShapeDecorativeHidden([Xml.XmlElement]$elem, $nsm) {
    # extLst/ext under p:cNvPr use the drawingml ("a") namespace, not the
    # presentationml ("p") one - this matches the markup real PowerPoint
    # and publisher templates actually use for the decorative extension.
    $cNvPr = $elem.SelectSingleNode('.//p:cNvPr', $nsm)
    [void]$cNvPr.SetAttribute('hidden', '1')
    $doc = $elem.OwnerDocument
    $extLst = $cNvPr.SelectSingleNode('a:extLst', $nsm)
    if ($null -eq $extLst) {
        $extLst = $doc.CreateElement('a', 'extLst', $NS_A)
        [void]$cNvPr.AppendChild($extLst)
    }
    $ext = $doc.CreateElement('a', 'ext', $NS_A)
    [void]$ext.SetAttribute('uri', '{C183D7F6-B498-43B3-948B-1728B52AA6E4}')
    $adecNs = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
    $dec = $doc.CreateElement('adec', 'decorative', $adecNs)
    [void]$dec.SetAttribute('val', '1')
    [void]$ext.AppendChild($dec)
    [void]$extLst.AppendChild($ext)
}

function Invoke-ImbkMacro($zip, [string[]]$order, $docs, [double]$slideW, [double]$slideH, [bool]$hideTopTitle) {
    $doneCount = 0; $reviewList = @(); $alreadyCount = 0; $markedCount = 0; $doneSlides = @{}
    foreach ($name in $order) {
        $doc = $docs[$name].doc; $nsm = $docs[$name].nsm
        $result = Invoke-ImbkLayout $zip $name $doc $nsm $slideW $slideH $hideTopTitle
        if ($result -eq 'done') { $doneCount++; $doneSlides[$name] = $true }
        elseif ($result -eq 'needs-review') { $reviewList += ($order.IndexOf($name) + 1) }
        elseif ($result -eq 'already-processed') { $alreadyCount++ }
        elseif ($result -eq 'marked-only') { $markedCount++ }
    }
    return @{ done = $doneCount; review = $reviewList; already = $alreadyCount; markedOnly = $markedCount; doneSlides = $doneSlides }
}

# ---------- Step 5: loose caption / credit text boxes -> real layout placeholders ----------
function New-LayoutPh([string]$id, [string]$name, [string]$idx, [long]$x, [long]$y, [long]$cx, [long]$cy, [string]$sz, [string]$prompt) {
    return @"
<p:sp xmlns:p="$NS_P" xmlns:a="$NS_A"><p:nvSpPr><p:cNvPr id="$id" name="$name"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="$idx"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$cx" cy="$cy"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="$sz"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>$prompt</a:t></a:r></a:p></p:txBody></p:sp>
"@
}

function Invoke-CaptionPlaceholders($zip, [string[]]$order, [double]$slideW, [double]$slideH) {
    $converted = 0; $layoutsDone = @{}
    foreach ($name in $order) {
        $text = Read-Entry $zip $name
        $doc = New-XmlDoc $text; $nsm = New-Nsmgr $doc
        $spTree = $doc.SelectSingleNode('//p:cSld/p:spTree', $nsm)
        $shapes = @($spTree.ChildNodes | Where-Object { $_.LocalName -in @('sp','pic','grpSp') })
        $hasPic = $false
        foreach ($e in $shapes) { if (Test-IsPicture $e $nsm) { $hasPic = $true; break } }
        if (-not $hasPic) { continue }
        # already has caption placeholders on this slide? nothing to do
        $already = $doc.SelectSingleNode('//p:sp[.//p:ph[@idx="13" or @idx="14"]]', $nsm)
        if ($already) { continue }
        # loose text boxes with text, in slide order
        $boxes = @()
        foreach ($e in $shapes) {
            if ($e.LocalName -ne 'sp') { continue }
            if ($e.SelectSingleNode('.//p:cNvSpPr[@txBox="1"]', $nsm) -and (Test-HasText $e $nsm) -and -not (Get-PhInfo $e $nsm)) { $boxes += $e }
        }
        if ($boxes.Count -eq 0) { continue }
        $cap = $null
        foreach ($b in $boxes) {
            $t = (Get-SlideText $b $nsm).Trim()
            if ((Get-FigureNumber $t) -or ($t -match '^(Figure|Fig\.?|Table|Box|Exhibit)\s+\S+')) { $cap = $b; break }
        }
        if ($null -eq $cap) { continue }
        $credit = $null
        $capOff = Get-ShapeOff $cap 'p:spPr' $nsm
        foreach ($b in $boxes) {
            if ($b -eq $cap) { continue }
            $o = Get-ShapeOff $b 'p:spPr' $nsm
            if ($o -and $capOff -and $o.y -gt $capOff.y) { $credit = $b; break }
        }
        # make sure the layout has the two placeholders
        $names = Get-LayoutMasterNames $zip $name
        $layoutName = $names[0]
        if (-not $layoutsDone.ContainsKey($layoutName)) {
            $ltext = Read-Entry $zip $layoutName
            $ldoc = New-XmlDoc $ltext; $lnsm = New-Nsmgr $ldoc
            if (-not $ldoc.SelectSingleNode('//p:sp[.//p:ph[@idx="13"]]', $lnsm)) {
                $ltree = $ldoc.SelectSingleNode('//p:cSld/p:spTree', $lnsm)
                $x = [long]($slideW * 0.05); $w = [long]($slideW * 0.90)
                $capXml = New-LayoutPh '9013' 'Caption Placeholder' '13' $x ([long]($slideH * 0.845)) $w 330000 '1400' 'Click to add figure caption'
                $crXml  = New-LayoutPh '9014' 'Credit Placeholder'  '14' $x ([long]($slideH * 0.901)) $w 250000 '900'  'Click to add credit line'
                foreach ($xml in @($capXml, $crXml)) {
                    $nd = New-XmlDoc $xml
                    [void]$ltree.AppendChild($ldoc.ImportNode($nd.DocumentElement, $true))
                }
                Write-Entry $zip $layoutName $ldoc.OuterXml
            }
            $layoutsDone[$layoutName] = $true
        }
        # convert the text boxes in place (keeps their own position, size and text formatting)
        $pairs = @(,@($cap, 'Caption', '13'))
        if ($credit) { $pairs += ,@($credit, 'Credit', '14') }
        foreach ($pr in $pairs) {
            $sp = $pr[0]
            $cNvPr = $sp.SelectSingleNode('p:nvSpPr/p:cNvPr', $nsm)
            $cNvSpPr = $sp.SelectSingleNode('p:nvSpPr/p:cNvSpPr', $nsm)
            if ($cNvSpPr.HasAttribute('txBox')) { [void]$cNvSpPr.RemoveAttribute('txBox') }
            if (-not $cNvSpPr.SelectSingleNode('a:spLocks', $nsm)) {
                $lock = $doc.CreateElement('a', 'spLocks', $NS_A); [void]$lock.SetAttribute('noGrp', '1')
                [void]$cNvSpPr.AppendChild($lock)
            }
            $nvPr = $sp.SelectSingleNode('p:nvSpPr/p:nvPr', $nsm)
            $ph = $doc.CreateElement('p', 'ph', $NS_P)
            [void]$ph.SetAttribute('type', 'body'); [void]$ph.SetAttribute('sz', 'quarter'); [void]$ph.SetAttribute('idx', $pr[2])
            [void]$nvPr.AppendChild($ph)
            $converted++
        }
        Write-Entry $zip $name $doc.OuterXml
    }
    return $converted
}

function Get-SlideSize($zip) {
    $presText = Read-Entry $zip 'ppt/presentation.xml'
    $doc = New-XmlDoc $presText; $nsm = New-Nsmgr $doc
    $sz = $doc.SelectSingleNode('//p:sldSz', $nsm)
    return @{ w = [double]$sz.GetAttribute('cx'); h = [double]$sz.GetAttribute('cy') }
}

function Fix-One([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return "$path : file not found" }
    $path = (Resolve-Path -LiteralPath $path).Path
    $name = Split-Path $path -Leaf
    $script:BoxesResized = 0
    $bak = Backup-Name $path "_BEFORE_IMBKfix"
    Copy-Item $path $bak
    try {
        $zip = [IO.Compression.ZipFile]::Open($path, 'Update')
        try {
            $order = Get-SlideOrder $zip
            $sz = Get-SlideSize $zip
            $docs = @{}
            foreach ($slideName in $order) {
                $text = Read-Entry $zip $slideName
                $doc = New-XmlDoc $text; $nsm = New-Nsmgr $doc
                $docs[$slideName] = @{ doc = $doc; nsm = $nsm }
            }

            $imbkResult = Invoke-ImbkMacro $zip $order $docs $sz.w $sz.h $true
            # Save the IMBK layout changes before the next step re-reads from the zip.
            foreach ($slideName in $order) { Write-Entry $zip $slideName $docs[$slideName].doc.OuterXml }
        } finally { }

        # Number repeated titles, then rename shapes, then the language fix and document title,
        # exactly as SLID_Macro does, now working from the IMBK-updated slides.
        $order2 = Get-SlideOrder $zip
        $dedupResult = Invoke-TitleDedup $zip $order2
        Invoke-RenameShapes $zip $order2 $dedupResult.docs $imbkResult.doneSlides
        $phCount = Invoke-CaptionPlaceholders $zip $order2 $sz.w $sz.h
        $langCount = Invoke-LanguageFix $zip
        $firstDoc = $dedupResult.docs[$order2[0]].doc; $firstNsm = $dedupResult.docs[$order2[0]].nsm
        $firstTitle = Get-TitleText $firstDoc $firstNsm
        $subtitleText = Get-ShapeTextByName $firstDoc $firstNsm 'Subtitle'
        $titleSet = $false
        if ($firstTitle -and $firstTitle.text.Trim().Length -gt 0) {
            $docTitleText = $firstTitle.text
            if ($subtitleText -and $subtitleText.Trim().Length -gt 0) {
                $docTitleText = $docTitleText + ': ' + $subtitleText
            }
            $titleSet = Set-DocTitle $zip $docTitleText
        }
        [void]$zip.Dispose()
        Add-WorkLog 'IMBK Macro' $path $path

        $check = [IO.Compression.ZipFile]::OpenRead($path); [void]$check.Dispose()
        $titleMsg = if ($titleSet) { "set to '$docTitleText'" } else { "not changed (no title found on slide 1)" }
        $reviewMsg = if ($imbkResult.review.Count -gt 0) { ", slides to check by hand: " + ($imbkResult.review -join ',') } else { "" }
        return "$name : figures laid out $($imbkResult.done), caption boxes resized $($script:BoxesResized), already done $($imbkResult.already), accessibility marking added $($imbkResult.markedOnly)$reviewMsg, titles numbered $($dedupResult.changed), caption boxes made placeholders $phCount, language tags changed $langCount, document title $titleMsg. Backup: $(Split-Path $bak -Leaf)"
    }
    catch {
        try { [void]$zip.Dispose() } catch {}
        Copy-Item $bak $path -Force
        return "$name : FAILED, file put back from the backup. $($_.Exception.Message)"
    }
}

Write-Host "IMBK Macro - running version $IMBK_VERSION" -ForegroundColor Cyan
Write-Host ''
$files = if ($env:IMBK_FILES) { @($env:IMBK_FILES.Split(';') | Where-Object { $_ }) } elseif ($args.Count -gt 0) { $args } else { Select-Files }
if ($files.Count -eq 0) { Write-Host "No files selected."; exit }
$report = foreach ($f in $files) { Write-Host ('Fixing: ' + [IO.Path]::GetFileName($f)); Fix-One $f }
Write-Host ""
Write-Host "RESULT"
$report | ForEach-Object { Write-Host " - $_" }
Write-Host ""
if (-not $env:S4C_AUTO) { Read-Host "Press Enter to close" }
