# Measures Word's document themes from observable output (manual; needs desktop Word on Windows,
# not run in CI).
#
#   pwsh -File src/core/scripts/docx-word-theme-capture.ps1 -Themes Atlas,Celestial,Depth [-Out <file.json>]
#
# For each named theme the script finds the `.thmx` Word offers in its Design > Themes gallery
# (the installed Document Themes folder or Word's downloaded Document Themes), applies it to a new
# blank document with the documented `Document.ApplyDocumentTheme`, saves the document with
# `Document.SaveAs2` and reads `word/theme/theme1.xml` from the saved file. It also reads the
# documented `ThemeColorScheme` and `ThemeFontScheme` objects of the open document, so the two
# observations can be compared. Recorded per theme: the twelve palette values (a system colour's
# `lastClr`), the Latin major and minor faces, the `a:fmtScheme` XML, and whether a fill style is a
# picture fill (`a:blipFill`). Theme packages, thumbnails and pictures are not copied.
# The values go to `drawingml/theme-gallery-legacy.ts` and `theme-formats-legacy.ts` (see
# PROVENANCE.md, "Word Design gallery data"). Always quits Word in a finally block.
param(
	[Parameter(Mandatory)][string[]]$Themes,
	[string]$Out = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-theme-capture.json'),
	[int]$WaitMinutes = 15
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
for ($waited = 0; (Get-Process WINWORD -ErrorAction SilentlyContinue); $waited += 30) {
	if ($waited -ge $WaitMinutes * 60) { throw 'WINWORD.EXE kept running; giving up.' }
	Start-Sleep -Seconds 30
}
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-theme-capture'
New-Item -ItemType Directory -Force $work | Out-Null

$folders = @(
	(Join-Path $env:ProgramFiles 'Microsoft Office\root\Document Themes 16'),
	(Join-Path $env:APPDATA 'Microsoft\Templates\LiveContent\16\Managed\Document Themes'),
	(Join-Path $env:APPDATA 'Microsoft\Templates\LiveContent\16\User\Document Themes'),
	(Join-Path $env:APPDATA 'Microsoft\Templates\Document Themes')
) | Where-Object { Test-Path $_ }

function Find-Theme([string]$name) {
	foreach ($folder in $folders) {
		$found = Get-ChildItem -Path $folder -Recurse -Filter '*.thmx' -ErrorAction SilentlyContinue |
			Where-Object { $_.BaseName -eq $name -or $_.BaseName.EndsWith("[[fn=$name]]") } |
			Select-Object -First 1
		if ($found) { return $found.FullName }
	}
	return $null
}

function Read-ThemeXml([string]$docx) {
	$zip = [IO.Compression.ZipFile]::OpenRead($docx)
	try {
		$entry = $zip.GetEntry('word/theme/theme1.xml')
		$reader = [IO.StreamReader]::new($entry.Open())
		try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
	} finally { $zip.Dispose() }
}

$slots = 'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'
# WdThemeColorIndex values for the same slots (documented in the Word object model).
$comSlots = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
$results = [ordered]@{}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	foreach ($name in $Themes) {
		$path = Find-Theme $name
		if (-not $path) {
			$results[$name] = [ordered]@{ missing = $true }
			Write-Warning "No .thmx named $name on this machine; skipped."
			continue
		}
		$doc = $word.Documents.Add()
		try {
			$doc.ApplyDocumentTheme($path)
			$com = [ordered]@{}
			for ($i = 0; $i -lt $comSlots.Count; $i++) {
				# RGB is stored as 0x00BBGGRR.
				$value = $doc.DocumentTheme.ThemeColorScheme.Colors($comSlots[$i] + 1).RGB
				$com[$slots[$i]] = '{0:X2}{1:X2}{2:X2}' -f ($value -band 0xFF), (($value -shr 8) -band 0xFF), (($value -shr 16) -band 0xFF)
			}
			$fonts = $doc.DocumentTheme.ThemeFontScheme
			$comMajor = $fonts.MajorFont.Item(1).Name
			$comMinor = $fonts.MinorFont.Item(1).Name
			$saved = Join-Path $work ("$name.docx" -replace '\s', '-')
			$doc.SaveAs2($saved, 12)
		} finally { $doc.Close(0) }
		[xml]$theme = Read-ThemeXml $saved
		$ns = New-Object Xml.XmlNamespaceManager $theme.NameTable
		$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
		$colors = foreach ($slot in $slots) {
			$node = $theme.SelectSingleNode("//a:clrScheme/a:$slot/*", $ns)
			if ($node.LocalName -eq 'sysClr') { $node.GetAttribute('lastClr') } else { $node.GetAttribute('val') }
		}
		$format = $theme.SelectSingleNode('//a:fmtScheme', $ns)
		$results[$name] = [ordered]@{
			file = Split-Path -Leaf $path
			colors = @($colors)
			comColors = @($com.Values)
			majorFont = $theme.SelectSingleNode('//a:majorFont/a:latin', $ns).GetAttribute('typeface')
			minorFont = $theme.SelectSingleNode('//a:minorFont/a:latin', $ns).GetAttribute('typeface')
			comMajorFont = $comMajor
			comMinorFont = $comMinor
			pictureFills = $theme.SelectNodes('//a:fmtScheme//a:blipFill', $ns).Count
			fmtScheme = $format.OuterXml
		}
	}
} finally {
	$word.Quit(0)
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
Write-Output "Wrote $Out"
