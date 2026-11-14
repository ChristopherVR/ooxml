# Generates the Word documents behind native-reference.json: DOCPROPERTY, AUTONUM, LISTNUM and
# SYMBOL fields inserted and updated by Word 16 through its documented automation, saved as
# .docx and exported as PDF. Word saves AUTONUM, LISTNUM and SYMBOL without a result, so what
# they show is read from the exported PDFs (pdftotext -layout, and the fonts of the PDF's text
# spans for SYMBOL) into native-reference.json; the PDFs are not committed. DOCPROPERTY results
# are the cached results in doc-property.docx. The name Word records as the author on save is
# removed from each package.
#
# Each case line is `Style|text with {FIELD CODE} fields`; `Style@N` also sets outline level N.
# A line `===` starts a new document. Windows only; run from any folder:
#   powershell -ExecutionPolicy Bypass -File generate-word-computed-fields.ps1
param([switch]$PropertiesOnly)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$cases = @{
	'auto-number' = @'
Normal|{AUTONUM} a
Normal|{AUTONUM} b
Normal|{AUTONUM \s :} c
Normal|{AUTONUM \* roman} d
Normal|{AUTONUM} e {AUTONUM}
Heading 1|{AUTONUMLGL} H1
Heading 2|{AUTONUMLGL} H2
Heading 2|{AUTONUMLGL \e} H2e
Heading 3|{AUTONUMLGL \s -} H3
Heading 1|{AUTONUMLGL} H1b
Normal|{AUTONUMLGL} N
Heading 1|{AUTONUMOUT} O1
Heading 2|{AUTONUMOUT} O2
Heading 3|{AUTONUMOUT} O3
Heading 4|{AUTONUMOUT} O4
Heading 5|{AUTONUMOUT} O5
Heading 6|{AUTONUMOUT} O6
Heading 7|{AUTONUMOUT} O7
Heading 8|{AUTONUMOUT} O8
Heading 9|{AUTONUMOUT} O9
Heading 2|{AUTONUMOUT} O2b
Normal|{AUTONUMOUT} ON
Normal|{AUTONUM} after
===
Normal|{AUTONUM} n1
Normal|{AUTONUM} n2
Heading 1|{AUTONUM} h1
Normal|{AUTONUM} n3
Normal|{AUTONUMLGL} n4
Heading 1|{AUTONUMOUT} h1b
Heading 2|{AUTONUM} h2
Normal|{AUTONUM} n5
Normal|plain
Heading 2|{AUTONUMLGL} h2b
Heading 1|plain heading
Heading 2|{AUTONUMLGL} h2c
Heading 3|{AUTONUMOUT} h3
Normal|{AUTONUMLGL \s :} n6
===
Heading 2|{AUTONUMLGL} first h2
Heading 3|{AUTONUMOUT} h3
Normal|{AUTONUMOUT} n
Normal|{AUTONUM \s ")"} n
Normal|{AUTONUM \* ALPHABETIC \s ""} n
Normal|{AUTONUMLGL \e \s :} n
Normal|{AUTONUMOUT \s :} n
Normal|{AUTONUM \* Arabic} n
Normal|{AUTONUM \* CardText} n
Normal|{AUTONUM \* Upper} n
===
Normal|{AUTONUMLGL} n1
Normal|{AUTONUMOUT} n2
Normal|{AUTONUM} n3
Heading 2|{AUTONUMLGL} h2
Normal|{AUTONUMLGL} n4
Heading 1|{AUTONUMLGL} h1
Heading 1|{AUTONUMOUT} h1
Normal|{AUTONUMOUT} n5
Heading 4|{AUTONUMOUT} h4
Normal|{AUTONUMOUT} n6
Normal|{AUTONUMOUT \e} n7
Normal|{AUTONUMLGL \s "--"} n8
===
Heading 1|{AUTONUMLGL} h1
Normal|{AUTONUMLGL} n
Heading 2|{AUTONUMLGL} h2
Normal|{AUTONUMLGL} n
Heading 1|{AUTONUMLGL} h1
Normal|{AUTONUMLGL} n
Normal|{AUTONUMLGL} n
Heading 1|{AUTONUMLGL} h1
Normal|{AUTONUMLGL} n
Heading 1|plain
Normal|{AUTONUMLGL} n
Heading 2|{AUTONUMLGL} h2
Heading 3|{AUTONUMLGL} h3
Heading 2|{AUTONUMLGL} h2
Normal|{AUTONUMLGL} n
Heading 3|{AUTONUMLGL} h3
Normal|{AUTONUMLGL} n
===
Heading 1|{AUTONUMLGL} h1
Normal@2|{AUTONUMLGL} outline 2
Normal|{AUTONUMLGL} body
Normal@1|{AUTONUMOUT} outline 1
===
Heading 1|{AUTONUMLGL} h
Heading 2|plain
Normal|{AUTONUMLGL} n
'@
	'list-number' = @'
Normal|a {LISTNUM} b {LISTNUM} c {LISTNUM}
Normal|{LISTNUM} next para
Normal|{LISTNUM \l 2} l2 {LISTNUM \l 2} {LISTNUM \l 3}
Normal|{LISTNUM} back
Normal|{LISTNUM NumberDefault \l 4}.{LISTNUM \l 5}.{LISTNUM \l 6}.{LISTNUM \l 7}.{LISTNUM \l 8}.{LISTNUM \l 9}
Normal|{LISTNUM OutlineDefault} o1 {LISTNUM OutlineDefault \l 2} {LISTNUM OutlineDefault \l 3} {LISTNUM OutlineDefault \l 4} {LISTNUM OutlineDefault \l 5}.{LISTNUM OutlineDefault \l 6}.{LISTNUM OutlineDefault \l 7}.{LISTNUM OutlineDefault \l 8}.{LISTNUM OutlineDefault \l 9}
Normal|{LISTNUM LegalDefault} g1 {LISTNUM LegalDefault \l 2} {LISTNUM LegalDefault \l 3} {LISTNUM LegalDefault \l 2}
Normal|{LISTNUM \s 5} s5 {LISTNUM} {LISTNUM \l 2 \s 3} {LISTNUM \l 2}
Normal|{LISTNUM \* roman}x
List Number|item {LISTNUM} {LISTNUM}
List Number|item2 {LISTNUM}
List Number|item3 {LISTNUM \l 1} {LISTNUM NumberDefault}
Normal|{LISTNUM} after
Heading 1|{LISTNUM} head
===
Normal|{LISTNUM \l 3} x
Normal|{LISTNUM} y
===
Normal|{LISTNUM \l 2} {LISTNUM \l 3}
Normal|{LISTNUM} y
===
Normal|{LISTNUM \s 5} {LISTNUM} {LISTNUM}
Normal|{LISTNUM} y
===
Normal|{LISTNUM} {LISTNUM} z
Normal|{LISTNUM \l 2} {LISTNUM} {LISTNUM}
Normal|{LISTNUM} y
===
Normal|{LISTNUM LegalDefault} {LISTNUM} {LISTNUM}
Normal|{LISTNUM} y {LISTNUM}
===
Normal|x {LISTNUM \l 2} {LISTNUM}
Normal|x {LISTNUM} y
Normal|{LISTNUM} y
===
Normal|{LISTNUM} {LISTNUM \l 2}
Normal|{LISTNUM} back
===
Normal|{LISTNUM \l 3} x
Normal|{LISTNUM \* roman} back
===
Normal|{LISTNUM \l 3} x
Normal|{LISTNUM \s 5} back
===
Normal|{LISTNUM \l 2} x
Normal|x {LISTNUM} y
Normal|{LISTNUM} back
===
Normal|{LISTNUM \l 3} {LISTNUM}
Normal|{LISTNUM} back
===
Normal|{LISTNUM \l 2} {LISTNUM \l 2}
Normal|{LISTNUM} back
===
Normal|{LISTNUM \l 2} {LISTNUM \l 4} {LISTNUM \l 3}
Normal|{LISTNUM} back
===
Normal|{LISTNUM \l 1} {LISTNUM \l 3} {LISTNUM \l 3}
Normal|{LISTNUM} back
===
List Number|item
Normal|{LISTNUM} x
List Number|item2
===
List Bullet|b
Normal|{LISTNUM} y
===
List Number|item
Normal|{LISTNUM \l 2} x
List Number|item2
'@
	'symbol' = @'
Normal|x{SYMBOL 65}x
Normal|x{SYMBOL 0x41}x
Normal|x{SYMBOL 0X41}x
Normal|x{SYMBOL 183 \f Symbol}x
Normal|x{SYMBOL 61623 \f Symbol}x
Normal|x{SYMBOL 0xF0B7 \f "Symbol"}x
Normal|x{SYMBOL 8364 \u}x
Normal|x{SYMBOL 0x20AC \u}x
Normal|x{SYMBOL 128}x
Normal|x{SYMBOL 128 \a}x
Normal|x{SYMBOL 8364}x
Normal|x{SYMBOL 65 \s 20}x
Normal|x{SYMBOL 74 \f Wingdings \s 16 \h}x
Normal|x{SYMBOL 97 \* Upper}x
Normal|x{SYMBOL}x
Normal|x{SYMBOL abc}x
Normal|x{SYMBOL 65.7}x
Normal|x{SYMBOL -5}x
Normal|x{SYMBOL 0}x
Normal|x{SYMBOL 70000 \u}x
Normal|x{SYMBOL 300}x
Normal|x{SYMBOL 233 \f "Times New Roman" \s 9.5}x
Normal|x{SYMBOL 65 \f Arial \u}x
Normal|x{SYMBOL 147 \j}x
Normal|x{SYMBOL 33088 \j}x
Normal|x{SYMBOL 65 \s 11.5}{SYMBOL 65 \s 0}{SYMBOL 65 \s abc}x
Normal|x{SYMBOL 97 \f Wingdings}{SYMBOL 0x61 \f "Wingdings" \u}{SYMBOL 61537 \f Wingdings}{SYMBOL 97 \f "Courier New"}x
Normal|x{SYMBOL "65"}{SYMBOL 65 \f}{SYMBOL 65 \f Symbol}{SYMBOL 0xF041}x
'@
}

$docProperty = @(
	'DOCPROPERTY Title', 'DOCPROPERTY "Title" \* Upper', 'DOCPROPERTY title', 'DOCPROPERTY Subject',
	'DOCPROPERTY Author', 'DOCPROPERTY Keywords', 'DOCPROPERTY Comments', 'DOCPROPERTY Category',
	'DOCPROPERTY Manager', 'DOCPROPERTY Company', 'DOCPROPERTY HyperlinkBase', 'DOCPROPERTY LastSavedBy',
	'DOCPROPERTY Client', 'DOCPROPERTY client', 'DOCPROPERTY Count', 'DOCPROPERTY Rate', 'DOCPROPERTY Due',
	'DOCPROPERTY Flag', 'DOCPROPERTY Off', 'DOCPROPERTY Big', 'DOCPROPERTY Neg', 'DOCPROPERTY Missing',
	'DOCPROPERTY', 'DOCPROPERTY Count \* roman', 'DOCPROPERTY Due \@ "yyyy-MM-dd"',
	'DOCPROPERTY Client \* Caps', 'DOCPROPERTY Title \* FirstCap', 'DOCPROPERTY Template',
	'DOCPROPERTY NameofApplication', 'DOCPROPERTY "Hyperlink base"', 'DOCPROPERTY Pages',
	'DOCPROPERTY RevisionNumber', 'DOCPROPERTY Security', 'DOCPROPERTY ContentStatus',
	'DOCPROPERTY Count \* Roman', 'DOCPROPERTY Count \* ROMAN', 'DOCPROPERTY Count \* Alphabetic',
	'DOCPROPERTY Count \* alphabetic', 'DOCPROPERTY Count \* Hex', 'DOCPROPERTY Count \* Ordinal',
	'DOCPROPERTY Count \* OrdText', 'DOCPROPERTY Count \* CardText', 'DOCPROPERTY Count \* CardText \* Upper',
	'DOCPROPERTY Count \* Caps \* CardText', 'DOCPROPERTY Count \* Arabic', 'DOCPROPERTY Count \* ArabicDash',
	'DOCPROPERTY Count \* MERGEFORMAT', 'DOCPROPERTY Title \* roman', 'DOCPROPERTY Rate \* roman',
	'DOCPROPERTY Neg \* Arabic', 'DOCPROPERTY Flag \* Lower', 'DOCPROPERTY Due \@ "d MMMM yyyy"',
	'DOCPROPERTY Title \@ "yyyy"', 'DOCPROPERTY Title Extra', 'DOCPROPERTY "Client"',
	'DOCPROPERTY  Author  \* Upper \* MERGEFORMAT'
)

function Invoke-Com($target, [string]$member, [string]$kind, [object[]]$arguments) {
	[System.__ComObject].InvokeMember($member, $kind, $null, $target, $arguments)
}

# Removes the name Word records as the last author on save, so no user name is committed.
function Remove-LastAuthor([string]$path) {
	Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
	$zip = [System.IO.Compression.ZipFile]::Open($path, 'Update')
	try {
		$entry = $zip.GetEntry('docProps/core.xml')
		$reader = New-Object System.IO.StreamReader($entry.Open())
		$xml = $reader.ReadToEnd(); $reader.Close()
		$xml = $xml -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '' -replace '<dc:creator>(?!Ada Lovelace<)[^<]*</dc:creator>', ''
		$entry.Delete()
		$writer = New-Object System.IO.StreamWriter($zip.CreateEntry('docProps/core.xml').Open())
		$writer.Write($xml); $writer.Close()
	} finally { $zip.Dispose() }
}

function Add-Lines($word, $doc, [string[]]$lines) {
	$sel = $word.Selection
	foreach ($line in $lines) {
		$idx = $line.IndexOf('|')
		$style = $line.Substring(0, $idx); $outline = 0
		if ($style -match '^(.*)@(\d)$') { $style = $Matches[1]; $outline = [int]$Matches[2] }
		$sel.Style = $doc.Styles.Item($style)
		if ($outline) { $sel.ParagraphFormat.OutlineLevel = $outline }
		$parts = [regex]::Split($line.Substring($idx + 1), '(\{[^{}]*\})') | Where-Object { $_ -ne '' }
		foreach ($part in $parts) {
			if ($part.StartsWith('{')) {
				$null = $doc.Fields.Add($sel.Range, -1, $part.Substring(1, $part.Length - 2), $false)
				$null = $sel.EndKey(6)
			} else { $sel.TypeText($part) }
		}
		$sel.TypeParagraph()
	}
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
	foreach ($name in $(if ($PropertiesOnly) { @() } else { $cases.Keys })) {
		$blocks = $cases[$name] -split "`r?`n===`r?`n"
		for ($n = 0; $n -lt $blocks.Count; $n++) {
			$doc = $word.Documents.Add()
			Add-Lines $word $doc ($blocks[$n] -split "`r?`n" | Where-Object { $_ })
			$null = $doc.Fields.Update()
			$path = Join-Path $here ('{0}-{1}' -f $name, ($n + 1))
			$doc.SaveAs2("$path.docx", 16)
			$doc.ExportAsFixedFormat("$path.pdf", 17)
			$doc.Close(0)
			Remove-LastAuthor "$path.docx"
		}
	}
	$doc = $word.Documents.Add()
	$builtIn = $doc.BuiltInDocumentProperties
	$values = [ordered]@{ Title = 'my report title'; Subject = 'Field subject'; Author = 'Ada Lovelace';
		Keywords = 'alpha, beta'; Comments = 'Some comments'; Category = 'Cat A'; Manager = 'Boss';
		Company = 'Acme Ltd'; 'Hyperlink base' = 'http://example.com/' }
	foreach ($key in $values.Keys) {
		$property = Invoke-Com $builtIn 'Item' 'GetProperty' @($key)
		$null = Invoke-Com $property 'Value' 'SetProperty' @($values[$key])
	}
	$custom = $doc.CustomDocumentProperties
	# Office property types: 1 number, 2 yes/no, 3 date, 4 text, 5 float.
	foreach ($item in @(@('Client', 4, 'Contoso'), @('Count', 1, 42), @('Rate', 5, 3.5),
			@('Due', 3, [datetime]'2026-03-04'), @('Flag', 2, $true), @('Off', 2, $false),
			@('Big', 1, 1234567), @('Neg', 5, -2.25))) {
		$null = Invoke-Com $custom 'Add' 'InvokeMethod' @($item[0], $false, $item[1], $item[2])
	}
	$sel = $word.Selection
	foreach ($code in $docProperty) {
		$null = $doc.Fields.Add($sel.Range, -1, $code, $false)
		$null = $sel.EndKey(6)
		$sel.TypeParagraph()
	}
	$null = $doc.Fields.Update()
	$saved = Join-Path $env:TEMP 'doc-property.docx'
	$doc.SaveAs2([string]$saved, 16)
	$doc.Close(0)
	Move-Item -Force $saved (Join-Path $here 'doc-property.docx')
	Remove-LastAuthor (Join-Path $here 'doc-property.docx')
} finally { $word.Quit() }
