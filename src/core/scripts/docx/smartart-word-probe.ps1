# Probe (manual, Windows, desktop Word): how Word treats a SmartArt whose data part, cached drawing
# or definitions changed behind its back. Evidence for docs/agnostic-core-plan.md ("Word and an
# edited data part").
#
#   pwsh -File scripts/docx/smartart-word-probe.ps1 [-OutDir <folder>] [-MakeFixture]
#
# -MakeFixture first has Word insert a Basic Block List SmartArt with the texts Alpha..Epsilon and
# save it as docx/__fixtures__/smartart-word.docx (the committed fixture was made this way).
# Each variant built by smartart-word-variants.mjs is opened read-only; the script prints the node
# texts Word reports, which of the probe words appear in the picture Word draws (the EMF of the
# inline shape), and the shapes and texts of the drawing Word writes when it saves the file again.
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-smartart-probe'),
	[switch]$MakeFixture
)
$ErrorActionPreference = 'Stop'
$core = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
$fixture = Join-Path $core 'docx/__fixtures__/smartart-word.docx'
New-Item -ItemType Directory -Force $OutDir | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	if ($MakeFixture) {
		$doc = $word.Documents.Add()
		$doc.Content.Text = 'Before the diagram.'
		$layouts = $word.SmartArtLayouts
		$basic = $null
		for ($i = 1; $i -le $layouts.Count; $i++) {
			if ($layouts.Item($i).Id -eq 'urn:microsoft.com/office/officeart/2005/8/layout/default') { $basic = $layouts.Item($i); break }
		}
		$doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
		$shape = $doc.InlineShapes.AddSmartArt($basic, $doc.Paragraphs.Item(2).Range)
		$texts = @('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon')
		$nodes = $shape.SmartArt.AllNodes
		for ($i = 1; $i -le $nodes.Count; $i++) { $nodes.Item($i).TextFrame2.TextRange.Text = $texts[$i - 1] }
		$doc.SaveAs2($fixture, 16)
		$doc.Close(0)
	}
	Push-Location $core
	try { & bun scripts/docx/smartart-word-variants.mjs $fixture $OutDir; if ($LASTEXITCODE) { throw 'variants failed' } }
	finally { Pop-Location }

	$resaved = Join-Path $OutDir 'resaved'
	New-Item -ItemType Directory -Force $resaved | Out-Null
	foreach ($file in Get-ChildItem $OutDir -Filter '*.docx' | Sort-Object Name) {
		$doc = $null
		try {
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$inline = $doc.InlineShapes.Item(1)
			$nodes = $inline.SmartArt.AllNodes
			$texts = @(); for ($i = 1; $i -le $nodes.Count; $i++) { $texts += $nodes.Item($i).TextFrame2.TextRange.Text }
			$emf = [Text.Encoding]::Unicode.GetString([byte[]]$inline.Range.EnhMetaFileBits)
			$drawn = @('Alpha', 'Beta', 'EDITED', 'Gamma', 'Cacheonly', 'Zeta') | Where-Object { $emf.Contains($_) }
			$copy = Join-Path $resaved $file.Name
			$doc.SaveAs2($copy, 16)
			$doc.Close(0); $doc = $null
			$zip = [IO.Compression.ZipFile]::OpenRead($copy)
			try {
				$entry = $zip.Entries | Where-Object { $_.FullName -like 'word/diagrams/drawing*.xml' } | Select-Object -First 1
				$xml = if ($entry) { (New-Object IO.StreamReader($entry.Open())).ReadToEnd() } else { '' }
			} finally { $zip.Dispose() }
			$shapes = ([regex]::Matches($xml, '<dsp:sp ')).Count
			$saved = ([regex]::Matches($xml, '<a:t>([^<]*)') | ForEach-Object { $_.Groups[1].Value }) -join '|'
			Write-Host "$($file.BaseName): nodes [$($texts -join ', ')]; drawn [$($drawn -join ', ')]; re-saved drawing $shapes shapes [$saved]"
		} catch {
			Write-Host "$($file.BaseName): FAILED $($_.Exception.Message)" -ForegroundColor Red
		} finally { if ($doc) { $doc.Close(0) } }
	}
} finally { $word.Quit() }
