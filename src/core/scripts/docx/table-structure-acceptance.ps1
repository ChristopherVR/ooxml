# Manual acceptance check (Windows, desktop Word; not run in CI): files whose merged or nested
# table was restructured by the engine and saved by the writer must open in Word without repair
# and show the same grid: the same cells per row with the same text.
#
#   pwsh -File table-structure-acceptance.ps1 [-OutDir <folder>] [-SkipBuild]
#
# `bun scripts/docx/table-structure-acceptance-build.ts` writes the documents and manifest.json
# (text of every cell, per row); a nested table's file also lists its outer table under `<file>#outer`
# (empty text means "not checked").
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-table-structure-acceptance'),
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
if (-not $SkipBuild) {
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Get-ChildItem $OutDir -Filter '*.docx' | Remove-Item -Force
	Push-Location $root
	try {
		& bun scripts/docx/table-structure-acceptance-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "table-structure-acceptance-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json
$deadline = (Get-Date).AddMinutes(15)
while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
	if ((Get-Date) -gt $deadline) { throw 'Word stays busy' }
	Start-Sleep -Seconds 30
}

# Rows of cell texts as Word reports them (end-of-cell marks removed), through Cells in row order.
function Read-Cells($table) {
	$rows = @{}
	foreach ($cell in $table.Range.Cells) {
		# A cell holding a nested table also reports that table's text; keep the cell's own paragraphs.
		$text = ([string]$cell.Range.Text) -replace "[\r\a]+$", ''
		$key = [int]$cell.RowIndex
		if (-not $rows.ContainsKey($key)) { $rows[$key] = @() }
		$rows[$key] += $text
	}
	return @($rows.Keys | Sort-Object | ForEach-Object { , @($rows[$_]) })
}
function Compare-Rows($want, $got, $skipUnchecked) {
	$problems = @()
	if (@($want).Count -ne @($got).Count) { return @("Word has $(@($got).Count) rows, expected $(@($want).Count)") }
	for ($r = 0; $r -lt @($want).Count; $r++) {
		$wantRow = @($want[$r]); $gotRow = @($got[$r])
		if ($wantRow.Count -ne $gotRow.Count) { $problems += "row $($r + 1): Word has $($gotRow.Count) cells, expected $($wantRow.Count)"; continue }
		for ($c = 0; $c -lt $wantRow.Count; $c++) {
			if ($skipUnchecked -and $wantRow[$c] -eq '') { continue }
			$expected = ([string]$wantRow[$c]) -replace "`n", "`r"
			if ($gotRow[$c] -ne $expected) { $problems += "row $($r + 1) cell $($c + 1): Word has '$($gotRow[$c])', expected '$expected'" }
		}
	}
	return $problems
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$results = @()
$failures = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.docx' | Sort-Object Name) {
		$status = 'ok'; $detail = ''; $doc = $null
		try {
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$caption = [string]$doc.ActiveWindow.Caption
			$problems = @()
			$table = $doc.Tables.Item(1)
			$nested = $file.Name.StartsWith('nested')
			$target = if ($nested) { $table.Tables.Item(1) } else { $table }
			$problems += Compare-Rows $manifest.($file.Name) (Read-Cells $target) $false
			if ($manifest.PSObject.Properties["$($file.Name)#outer"]) {
				$problems += Compare-Rows $manifest.("$($file.Name)#outer") (Read-Cells $table) $true
			}
			if ($caption -match 'Repaired') { $status = 'REPAIRED'; $detail = $caption }
			elseif ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
			else { $detail = "$($target.Rows.Count) rows, grid as expected" }
		} catch {
			$status = 'FAILED'; $detail = $_.Exception.Message
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
if ($failures.Count) {
	Write-Host "FAILED: $($failures -join ', ')" -ForegroundColor Red
	exit 1
}
Write-Host "All $($results.Count) documents opened in Word without repair, with the expected grid." -ForegroundColor Green
