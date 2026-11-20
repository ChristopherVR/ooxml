# Records where desktop Word draws floating objects (`wp:anchor`) anchored in table cells, and
# floating charts, and how the cell text and the row height follow them (manual; needs desktop
# Word on Windows and PyMuPDF, not CI). From src/core:
#
#   pwsh -File scripts/docx-word-cell-float-probe.ps1 [-Work <dir>]
#
# 1. `bun scripts/docx-word-cell-float-cases.ts` writes the documents of
#    `docx/__fixtures__/cell-floats/cases.ts` to <Work>/docs (and the demonstration document to
#    <Work>/demo).
# 2. Each is opened in an owned, invisible Word with OpenAndRepair off (so a package Word would
#    have to repair throws), exported to PDF and closed without saving.
# 3. `scripts/docx-word-cell-float-pdf.py` reads the case PDFs into
#    `docx/__fixtures__/cell-floats/word-positions.json`.
# No user documents are opened.
param(
	[string]$Work = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-cell-float-probe')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$docs = Join-Path $Work 'docs'
$demo = Join-Path $Work 'demo'
$pdf = Join-Path $Work 'pdf'
New-Item -ItemType Directory -Force $docs, $demo, $pdf | Out-Null
Push-Location $root
try {
	& bun scripts/docx-word-cell-float-cases.ts $docs
	if ($LASTEXITCODE -ne 0) { throw "docx-word-cell-float-cases.ts failed ($LASTEXITCODE)" }
	& bun scripts/docx-word-cell-float-cases.ts $demo demo
	if ($LASTEXITCODE -ne 0) { throw "docx-word-cell-float-cases.ts failed ($LASTEXITCODE)" }
} finally { Pop-Location }

function Open-Document($app, [string]$path) {
	# Named arguments through InvokeMember: PowerShell 7 cannot pass COM's "missing" positionally.
	$documents = $app.Documents
	$names = [string[]]@('FileName', 'ConfirmConversions', 'ReadOnly', 'AddToRecentFiles', 'OpenAndRepair', 'NoEncodingDialog')
	$values = [object[]]@($path, $false, $true, $false, $false, $true)
	return , [System.__ComObject].InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $documents, $values, $null, $null, $names)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0 # wdAlertsNone
Set-Content -Encoding utf8 (Join-Path $pdf 'word-build.txt') $word.Build
try {
	foreach ($file in @(Get-ChildItem $docs -Filter '*.docx') + @(Get-ChildItem $demo -Filter '*.docx')) {
		$doc = Open-Document $word ([string]$file.FullName)
		try {
			$target = if ($file.DirectoryName -eq $demo) { $demo } else { $pdf }
			$doc.ExportAsFixedFormat([string](Join-Path $target "$($file.BaseName).pdf"), 17) # wdExportFormatPDF
			Write-Host "$($file.BaseName) ok"
		} finally {
			$doc.Saved = $true
			$doc.Close($false)
		}
	}
} finally {
	$word.Quit($false)
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
& python -I (Join-Path $PSScriptRoot 'docx-word-cell-float-pdf.py') $pdf (Join-Path $root 'docx/__fixtures__/cell-floats/word-positions.json')
