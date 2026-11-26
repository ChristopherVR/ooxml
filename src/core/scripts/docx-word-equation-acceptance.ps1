# Manual acceptance check for the Word Equation tab's saved Office Math (not run in CI; needs
# desktop Word on Windows).
#
#   pwsh -File scripts/docx-word-equation-acceptance.ps1 [-OutDir <folder>] [-SkipBuild]
#
# 1. `bun scripts/docx-word-equation-build.ts` writes documents with `saveDocx`: every structure of
#    the Equation tab's galleries (display and inline, placeholders left empty), the built-in
#    equations, Linear-format equations, Unicode linear input built up and empty equations.
# 2. Each file is opened in Word through COM. A file fails when Word throws (an unattended Word
#    refuses a package it would have to repair), when the caption says "Repaired", or when Word
#    finds a different number of equations (`Document.OMaths`) than manifest.json lists.
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-equation-acceptance'),
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.docx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Push-Location $root
	try {
		& bun scripts/docx-word-equation-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "docx-word-equation-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$failures = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.docx' | Sort-Object Name) {
		$doc = $null
		try {
			# ConfirmConversions off, ReadOnly, no recent-files entry.
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$caption = [string]$doc.ActiveWindow.Caption
			$count = [int]$doc.OMaths.Count
			$want = [int]$manifest.($file.Name)
			if ($caption -match 'Repaired') { $failures += "$($file.Name): Word repaired the file" }
			elseif ($count -ne $want) { $failures += "$($file.Name): Word finds $count equations, expected $want" }
			else { Write-Host "ok   $($file.Name) ($count equations)" }
		} catch {
			$failures += "$($file.Name): $($_.Exception.Message)"
		} finally {
			if ($doc) { $doc.Close($false) | Out-Null }
		}
	}
} finally {
	$word.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
if ($failures.Count) {
	$failures | ForEach-Object { Write-Host "FAIL $_" }
	exit 1
}
Write-Host "All $($manifest.PSObject.Properties.Count) documents open in Word without repair."
