# Shared helpers for the Word layout probes (dot-sourced, not run alone): hand-written minimal
# documents (Courier New 11 pt, line 240 auto, no default spacing), a package writer and the wait
# for a running WINWORD.EXE. The spacing probe keeps its own copy of the same writer.
function Wait-NoWord([int]$WaitMinutes = 20) {
	for ($waited = 0; (Get-Process WINWORD -ErrorAction SilentlyContinue); $waited += 30) {
		if ($waited -ge $WaitMinutes * 60) { throw 'WINWORD.EXE kept running; giving up.' }
		Start-Sleep -Seconds 30
	}
}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
function Tw([double]$pt) { [int]($pt * 20) }
function Para([string]$label, [string]$spacing = '', [string]$extra = '', [string]$run = '') {
	$sp = if ($spacing) { "<w:spacing $spacing/>" } else { '' }
	"<w:p><w:pPr>$sp$extra</w:pPr><w:r><w:t>$label</w:t></w:r>$run</w:p>"
}
function Sect([string]$type = 'nextPage') {
	"<w:sectPr><w:type w:val=`"$type`"/><w:pgSz w:w=`"12240`" w:h=`"15840`"/><w:pgMar w:top=`"1440`" w:right=`"1440`" w:bottom=`"1440`" w:left=`"1440`" w:header=`"720`" w:footer=`"720`" w:gutter=`"0`"/></w:sectPr>"
}
$bodySect = Sect 'nextPage'
function Write-Docx([string]$path, [string]$body, [int]$mode, [bool]$noHtml, [string]$sect) {
	$flag = if ($noHtml) { '<w:doNotUseHTMLParagraphAutoSpacing/>' } else { '' }
	$files = [ordered]@{
		'[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
		'_rels/.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
		'word/_rels/document.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
		'word/styles.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:styles $W><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=`"Courier New`" w:hAnsi=`"Courier New`" w:cs=`"Courier New`"/><w:sz w:val=`"22`"/><w:szCs w:val=`"22`"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=`"0`" w:line=`"240`" w:lineRule=`"auto`"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=`"paragraph`" w:default=`"1`" w:styleId=`"Normal`"><w:name w:val=`"Normal`"/></w:style></w:styles>"
		'word/numbering.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:numbering $W><w:abstractNum w:abstractNumId=`"1`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"&#xF0B7;`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr><w:rPr><w:rFonts w:ascii=`"Symbol`" w:hAnsi=`"Symbol`" w:hint=`"default`"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId=`"2`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"o`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr><w:rPr><w:rFonts w:ascii=`"Courier New`" w:hAnsi=`"Courier New`"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId=`"3`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"decimal`"/><w:lvlText w:val=`"%1.`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=`"1`"><w:abstractNumId w:val=`"1`"/></w:num><w:num w:numId=`"2`"><w:abstractNumId w:val=`"2`"/></w:num><w:num w:numId=`"3`"><w:abstractNumId w:val=`"3`"/></w:num></w:numbering>"
		'word/settings.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:settings $W><w:compat>$flag<w:compatSetting w:name=`"compatibilityMode`" w:uri=`"http://schemas.microsoft.com/office/word`" w:val=`"$mode`"/></w:compat></w:settings>"
		'word/document.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:document $W><w:body>$body$sect</w:body></w:document>"
	}
	if (Test-Path $path) { Remove-Item $path }
	$zip = [IO.Compression.ZipFile]::Open($path, 'Create')
	try {
		foreach ($name in $files.Keys) {
			$entry = $zip.CreateEntry($name)
			$stream = $entry.Open()
			$bytes = [Text.UTF8Encoding]::new($false).GetBytes($files[$name])
			$stream.Write($bytes, 0, $bytes.Length); $stream.Close()
		}
	} finally { $zip.Dispose() }
}

