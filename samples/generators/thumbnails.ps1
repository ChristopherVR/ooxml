# Renders the gallery thumbnails (samples/thumbnails/<id>.webp) from the sample files.
#
#   pwsh -File samples/generators/thumbnails.ps1 [-Only <id>[,<id>...]] [-WorkDir <folder>]
#
# Manual, Windows only: it needs desktop Word, Excel, PowerPoint and Visio, plus Python with
# PyMuPDF and Pillow (pip install pymupdf pillow). Each sample is opened read-only through the
# application's documented automation and exported (a slide or a Visio page to PNG, a Word page or
# an Excel sheet to PDF); thumbnails.py then crops and scales the export to 800 x 450 WebP.
# thumbnails.json says which slide, page or sheet stands for each sample. Nothing is saved back
# into a sample, and only the application instances this script starts are closed.
#
# Thumbnails are pictures of what Office draws, so they are not part of the deterministic build:
# rerun this after a sample's first slide, page or sheet changes, and commit the result.
param(
	[string[]]$Only = @(),
	[string]$WorkDir = (Join-Path ([IO.Path]::GetTempPath()) 'ooxml-sample-thumbnails')
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$samples = Split-Path -Parent $here
$manifest = Get-Content -Raw (Join-Path $samples 'manifest.json') | ConvertFrom-Json
$hints = Get-Content -Raw (Join-Path $here 'thumbnails.json') | ConvertFrom-Json
New-Item -ItemType Directory -Force $WorkDir | Out-Null
New-Item -ItemType Directory -Force (Join-Path $samples 'thumbnails') | Out-Null

function Get-Hint($id, $name, $default) {
	$hint = $hints.$id
	if ($null -ne $hint -and $null -ne $hint.$name) { return $hint.$name }
	return $default
}

function Export-Deck($sample, $source, $target) {
	$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
	$app = New-Object -ComObject PowerPoint.Application
	try {
		# ReadOnly, not untitled, no window: nothing is shown and nothing can be saved back.
		$deck = $app.Presentations.Open($source, -1, 0, 0)
		try {
			$deck.Slides.Item([int](Get-Hint $sample.id 'slide' 1)).Export($target, 'PNG', 1600, 900)
		} finally { $deck.Close() }
	} finally {
		if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
	}
}

function Export-Document($sample, $source, $target) {
	$app = New-Object -ComObject Word.Application
	try {
		$app.Visible = $false
		$app.DisplayAlerts = 0
		$missing = [Type]::Missing
		# FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, four passwords and Revert,
		# Format, Encoding, Visible, OpenAndRepair: read-only, hidden, and never repaired.
		$doc = $app.Documents.Open($source, $false, $true, $false, $missing, $missing, $missing,
			$missing, $missing, $missing, $missing, $false, $false)
		try {
			# wdExportFormatPDF = 17. The whole document; thumbnails.py takes the page it needs.
			$doc.ExportAsFixedFormat($target, 17)
		} finally { $doc.Close(0) }
	} finally {
		$app.Quit()
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
	}
}

function Export-Workbook($sample, $source, $target) {
	$app = New-Object -ComObject Excel.Application
	try {
		$app.Visible = $false
		$app.DisplayAlerts = $false
		# UpdateLinks 0, ReadOnly.
		$book = $app.Workbooks.Open($source, 0, $true)
		try {
			$name = Get-Hint $sample.id 'sheet' $null
			$sheet = if ($name) { $book.Worksheets.Item($name) } else { $book.Worksheets.Item(1) }
			# The page setup is changed in memory only (the workbook is never saved): one landscape
			# page, as wide as the range, without the sheet's own headers and footers.
			$setup = $sheet.PageSetup
			$range = Get-Hint $sample.id 'range' $null
			if ($range) { $setup.PrintArea = $range }
			$setup.Orientation = 2
			$setup.Zoom = $false
			$setup.FitToPagesWide = 1
			$setup.FitToPagesTall = 1
			$setup.CenterHeader = ''; $setup.LeftHeader = ''; $setup.RightHeader = ''
			$setup.CenterFooter = ''; $setup.LeftFooter = ''; $setup.RightFooter = ''
			$setup.PrintTitleRows = ''; $setup.PrintTitleColumns = ''
			$setup.PrintGridlines = $false
			$setup.PrintHeadings = $false
			$setup.CenterHorizontally = $false
			$setup.LeftMargin = 18; $setup.RightMargin = 18; $setup.TopMargin = 18; $setup.BottomMargin = 18
			# xlTypePDF = 0, xlQualityStandard = 0.
			$sheet.ExportAsFixedFormat(0, $target, 0, $false, $false, 1, 1, $false)
		} finally { $book.Close($false) }
	} finally {
		$app.Quit()
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
	}
}

function Export-Drawing($sample, $source, $target) {
	$app = New-Object -ComObject Visio.InvisibleApp
	try {
		$app.AlertResponse = 7
		# visOpenRO (2) + visOpenDontList (8) + visOpenMacrosDisabled (128).
		$doc = $app.Documents.OpenEx($source, 138)
		try {
			# visRasterUseCustomResolution = 3 at 144 pixels per inch.
			$app.Settings.SetRasterExportResolution(3, 144, 144, 0)
			$doc.Pages.Item([int](Get-Hint $sample.id 'page' 1)).Export($target)
		} finally { $doc.Close() }
	} finally {
		$app.Quit()
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
	}
}

$made = @()
foreach ($sample in $manifest.samples) {
	if ($Only.Count -gt 0 -and $Only -notcontains $sample.id) { continue }
	$source = (Resolve-Path (Join-Path $samples $sample.file)).Path
	$kind = if ($sample.product -in 'docx', 'xlsx') { 'pdf' } else { 'png' }
	$export = Join-Path $WorkDir "$($sample.id).$kind"
	if (Test-Path $export) { Remove-Item -Force $export }
	Write-Host "Rendering $($sample.file)"
	switch ($sample.product) {
		'pptx' { Export-Deck $sample $source $export }
		'docx' { Export-Document $sample $source $export }
		'xlsx' { Export-Workbook $sample $source $export }
		'visio' { Export-Drawing $sample $source $export }
	}
	if (-not (Test-Path $export)) { throw "$($sample.id): no export was written" }
	$anchor = Get-Hint $sample.id 'anchor' 'top'
	# Only a Word export has more than one page.
	$page = if ($sample.product -eq 'docx') { [int](Get-Hint $sample.id 'page' 1) } else { 1 }
	$made += "$($sample.id)|$export|$anchor|$page"
}
& python (Join-Path $here 'thumbnails.py') (Join-Path $samples 'thumbnails') @made
if ($LASTEXITCODE -ne 0) { throw "thumbnails.py failed ($LASTEXITCODE)" }
