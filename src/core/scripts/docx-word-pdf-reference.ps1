# Save the PDF export reference documents as PDF with desktop Word (Windows, manual): opens each
# `.docx` of `docx/pdf-export/__fixtures__/word-reference`, counts its pages and calls
# `ExportAsFixedFormat`. `docx/pdf-export/__fixtures__/word-reference/read-word-pdfs.mjs` then reads
# the line breaks of the PDFs into `word-lines.json`. The documents name free fonts (Carlito,
# Caladea, Liberation Sans and Serif), which must be installed. Only the Word process this script
# starts is closed; one it could not stop is reported by PID.
param(
	[string]$InputDir = (Join-Path $PSScriptRoot '../docx/pdf-export/__fixtures__/word-reference'),
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-pdf-reference')
)
$ErrorActionPreference = 'Stop'
$files = @(Get-ChildItem -LiteralPath $InputDir -Filter '*.docx' -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No DOCX files in $InputDir" }
$null = New-Item -ItemType Directory -Force $OutDir
$output = (Resolve-Path -LiteralPath $OutDir).Path

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3
$results = @()
try {
	Write-Host "Word $($word.Version), build $($word.Build)"
	foreach ($file in $files) {
		$doc = $null
		try {
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$doc.Repaginate()
			$doc.ExportAsFixedFormat((Join-Path $output "$($file.BaseName).pdf"), 17)
			$results += [pscustomobject]@{ File = $file.Name; Pages = $doc.ComputeStatistics(2) }
		} finally {
			if ($doc) {
				$doc.Close(0)
				[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
			}
		}
	}
} finally {
	# Quit only an instance this script started: another session may be using a running Word.
	if ($started.Count -gt 0) { $word.Quit(0) }
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
	foreach ($id in $started) {
		$process = Get-Process -Id $id -ErrorAction SilentlyContinue
		if ($process -and -not $process.WaitForExit(15000)) {
			Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
			if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Write-Warning "Could not stop WINWORD.EXE $id" }
		}
	}
}
$results | ConvertTo-Json | Set-Content -Encoding utf8 (Join-Path $output 'pages.json')
$results | Format-Table
