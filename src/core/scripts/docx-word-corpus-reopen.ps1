# Reopen check for the Word reference corpus (manual; needs desktop Word on Windows).
#
#   pwsh -File src/core/scripts/docx-word-corpus-reopen.ps1 [-OutDir <folder>] [-SkipBuild]
#
# 1. `bun scripts/docx-word-corpus-reopen-build.ts` loads each corpus document with the core, applies
#    one trivial edit and saves it with `saveDocx`.
# 2. Each saved file is opened in Word through COM. A file fails when Word throws (an unattended Word
#    refuses a package it would have to repair), when the caption says "Repaired", when the edit
#    marker is missing from the text, or when the page count differs from the one Word recorded for
#    the original by more than one page (the marker adds a few characters, which may legitimately
#    spill one line; a larger change means content was lost or reflowed).
# Failures are listed per document; the script exits 1 when any file fails.
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-corpus-reopen'),
	[switch]$SkipBuild,
	[int]$WaitMinutes = 15
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.docx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Push-Location $root
	try {
		& bun scripts/docx-word-corpus-reopen-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "docx-word-corpus-reopen-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json

for ($waited = 0; (Get-Process WINWORD -ErrorAction SilentlyContinue); $waited += 30) {
	if ($waited -ge $WaitMinutes * 60) { throw 'WINWORD.EXE kept running; giving up.' }
	Start-Sleep -Seconds 30
}

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
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$want = $manifest.($file.Name)
			$caption = [string]$doc.ActiveWindow.Caption
			$pages = [int]$doc.ComputeStatistics(2)
			$text = [string]$doc.Content.Text
			if ($caption -match 'Repaired') { $status = 'REPAIRED'; $detail = $caption }
			elseif ($want.edited -and -not $text.Contains($want.marker)) { $status = 'CONTENT'; $detail = "edit marker '$($want.marker)' missing" }
			elseif ([Math]::Abs($pages - $want.pageCount) -gt 1) { $status = 'PAGES'; $detail = "Word has $pages page(s), original had $($want.pageCount)" }
			else { $detail = "$pages page(s), original $($want.pageCount), marker $(if ($want.edited) { 'present' } else { 'not applicable' })" }
		} catch {
			$status = 'FAILED'
			$detail = $_.Exception.Message
		} finally {
			if ($doc) { try { $doc.Close(0) } catch { } }
		}
		$results += "{0,-8} {1}  {2}" -f $status, $file.Name, $detail
		if ($status -ne 'ok') { $failures += $file.Name }
	}
} finally {
	try { $word.Quit(0) } catch { }
}
$results
if ($failures.Count) { Write-Error "Word rejected or changed: $($failures -join ', ')"; exit 1 }
Write-Host "All $($results.Count) corpus documents reopen in Word without repair."
