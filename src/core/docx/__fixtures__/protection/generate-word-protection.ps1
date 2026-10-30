# Generates the protected documents protection.word.test.ts checks: one per editing restriction,
# protected by desktop Word (Document.Protect) with the password below, plus one saved in Word 2007
# compatibility mode. Word always writes the SHA-512 hash (cryptAlgorithmSid 14), even in compatibility
# modes, so no real Word fixture of the older hashes exists.
# Provenance: produced by Microsoft Word through COM on the machine that committed the files.
# Requires desktop Word. Run from any folder:
#   pwsh -File src/core/docx/__fixtures__/protection/generate-word-protection.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$password = 'Word-Pass-1'
$wdFormatXMLDocument = 12
$wdFormatDocument97 = 0
$modes = [ordered]@{ readOnly = 3; comments = 1; trackedChanges = 0; forms = 2 }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
function New-Protected($name, $type, $compat, $viaDoc) {
	$doc = $word.Documents.Add()
	$doc.Content.Text = "Protected $name"
	if ($compat) { $doc.SetCompatibilityMode(12) }
	if ($viaDoc) {
		$tmp = Join-Path $here "$name.doc"
		$doc.SaveAs2($tmp, $wdFormatDocument97)
		$doc.Close(0)
		$doc = $word.Documents.Open($tmp)
	}
	$doc.Protect($type, $false, $password)
	$path = Join-Path $here "$name.docx"
	$doc.SaveAs2($path, $wdFormatXMLDocument)
	$doc.Close(0)
	if ($viaDoc) { Remove-Item $tmp -ErrorAction SilentlyContinue }
}
try {
	foreach ($mode in $modes.Keys) { New-Protected $mode $modes[$mode] $false $false }
	New-Protected 'readOnly-compat12' 3 $true $false
} finally {
	$word.Quit(0)
}

# readOnly-iso.docx: Word's readOnly.docx with its protection element rewritten to the ISO 29500
# spelling (SHA-512 spin, this password), then opened in Word and unprotected with it, which shows
# Word accepts the plain spin in that form.
$isoPassword = 'Iso-Pass-1'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$salt = [Text.Encoding]::ASCII.GetBytes('0123456789abcdef')
$sha = [Security.Cryptography.SHA512]::Create()
$hash = $sha.ComputeHash([byte[]]($salt + [Text.Encoding]::Unicode.GetBytes($isoPassword)))
for ($i = 0; $i -lt 100000; $i++) { $hash = $sha.ComputeHash([byte[]]($hash + [BitConverter]::GetBytes([uint32]$i))) }
$iso = Join-Path $here 'readOnly-iso.docx'
Copy-Item (Join-Path $here 'readOnly.docx') $iso -Force
$archive = [IO.Compression.ZipFile]::Open($iso, 'Update')
try {
	$entry = $archive.GetEntry('word/settings.xml')
	$reader = New-Object IO.StreamReader($entry.Open())
	$xml = $reader.ReadToEnd(); $reader.Close()
	$element = '<w:documentProtection w:edit="readOnly" w:enforcement="1" w:algorithmName="SHA-512" w:hashValue="' + [Convert]::ToBase64String($hash) + '" w:saltValue="' + [Convert]::ToBase64String($salt) + '" w:spinCount="100000"/>'
	$xml = [regex]::Replace($xml, '<w:documentProtection[^>]*/>', $element)
	$entry.Delete()
	$writer = New-Object IO.StreamWriter($archive.CreateEntry('word/settings.xml').Open())
	$writer.Write($xml); $writer.Close()
} finally { $archive.Dispose() }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	$doc = $word.Documents.Open($iso, $false, $true)
	$doc.Unprotect($isoPassword)
	if ($doc.ProtectionType -ne -1) { throw 'Word did not accept the ISO form hash' }
	$doc.Close(0)
} finally { $word.Quit(0) }
