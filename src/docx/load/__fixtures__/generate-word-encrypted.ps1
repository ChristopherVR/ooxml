# Generates word-encrypted.docx: a document Word saved with the password to open 'open sesame'
# (agile encryption), used by the encrypted-package load test. Requires desktop Word (COM):
#   pwsh -File src/docx/load/__fixtures__/generate-word-encrypted.ps1
# The package is encrypted, so it cannot be scrubbed afterwards: RemovePersonalInformation makes
# Word drop the author and last-modified-by names on save.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$path = Join-Path $here 'word-encrypted.docx'
if (Test-Path $path) { Remove-Item -Force $path }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	$doc = $word.Documents.Add()
	$doc.Content.Text = 'Top secret paragraph.'
	$doc.RemovePersonalInformation = $true
	$doc.Password = 'open sesame'
	# 16 = wdFormatXMLDocument (.docx)
	$doc.SaveAs2($path, 16)
	$doc.Close(0)
} finally {
	$word.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output "wrote $path"
