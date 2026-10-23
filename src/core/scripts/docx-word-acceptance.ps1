# Manual acceptance check for SmartArt editing in the docx writer (not run in CI; needs desktop
# Word on Windows).
#
#   pwsh -File scripts/docx-word-acceptance.ps1 [-OutDir <folder>] [-SkipBuild]
#
# 1. `bun scripts/docx-acceptance-build.ts` writes documents with `saveDocx`: the SmartArt fixtures
#    untouched, after node edits, and with SmartArt inserted from presets.
# 2. Each file is opened in Word through COM. A file fails when Word throws (an unattended Word
#    refuses a package it would have to repair), when the caption says "Repaired", when a SmartArt
#    reports different node texts (`SmartArt.AllNodes`) or layout than manifest.json, or when a node's text is
#    missing from the picture Word draws for the diagram (`Range.EnhMetaFileBits`, the EMF Word
#    renders: evidence that Word laid the diagram out from the data, not from a stale drawing; inline
#    SmartArt only).
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-acceptance'),
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.docx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Push-Location $root
	try {
		& bun scripts/docx-acceptance-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "docx-acceptance-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$results = @()
$failures = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.docx' | Sort-Object Name) {
		$status = 'ok'
		$detail = ''
		$doc = $null
		try {
			# ConfirmConversions off, ReadOnly, no recent-files entry.
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$caption = [string]$doc.ActiveWindow.Caption
			$want = @($manifest.($file.Name))
			$seen = @()
			$problems = @()
			$index = 0
			# Inline SmartArt first, then floating (wp:anchor) SmartArt: the order the fixtures use.
			$graphics = @($doc.InlineShapes | Where-Object { $_.HasSmartArt }) + @($doc.Shapes | Where-Object { $_.HasSmartArt })
			foreach ($shape in $graphics) {
				$nodes = $shape.SmartArt.AllNodes
				$texts = @()
				for ($i = 1; $i -le $nodes.Count; $i++) { $texts += [string]$nodes.Item($i).TextFrame2.TextRange.Text }
				$seen += , $texts
				$expect = @($want[$index].texts)
				$layout = [string]$shape.SmartArt.Layout.Id
				if ($want[$index].layout -and $layout -ne $want[$index].layout) {
					$problems += "SmartArt $($index + 1): Word uses layout $layout, expected $($want[$index].layout)"
				}
				if (($texts -join '|') -ne ($expect -join '|')) {
					$problems += "SmartArt $($index + 1): Word has [$($texts -join ', ')], expected [$($expect -join ', ')]"
				}
				# Only an inline shape has a range whose picture is the diagram; a floating one is not checked.
				$emf = $null
				try { $emf = [Text.Encoding]::Unicode.GetString([byte[]]$shape.Range.EnhMetaFileBits) } catch { $emf = $null }
				foreach ($text in $(if ($emf) { $texts } else { @() })) {
					foreach ($token in ($text -split '\s+' | Where-Object { $_ })) {
						if (-not $emf.Contains($token)) { $problems += "SmartArt $($index + 1): '$token' is not drawn" }
					}
				}
				$index++
			}
			if ($caption -match 'Repaired') { $status = 'REPAIRED'; $detail = $caption }
			elseif ($index -ne $want.Count) { $status = 'CONTENT'; $detail = "Word sees $index SmartArt, expected $($want.Count)" }
			elseif ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
			else { $detail = "$index SmartArt: " + (($seen | ForEach-Object { "[$($_ -join ', ')]" }) -join ' ') }
		} catch {
			$status = 'FAILED'
			$detail = $_.Exception.Message
		} finally {
			if ($doc) { $doc.Close(0) | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
		}
		$results += [pscustomobject]@{ File = $file.Name; Status = $status; Detail = $detail }
		if ($status -ne 'ok') { $failures += $file.Name }
	}
} finally {
	$word.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
	[GC]::Collect()
}
$results | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
Write-Host "Documents are in $OutDir for inspection."
if ($failures.Count) {
	Write-Host "FAILED: $($failures -join ', ')" -ForegroundColor Red
	exit 1
}
Write-Host "All $($results.Count) documents opened in Word without repair, with the expected SmartArt." -ForegroundColor Green
