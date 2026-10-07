# Opens only generated fixtures read-only in an owned hidden Word instance.
# Each method starts from a fresh document and calls real Field.Update or Fields.Update.
param([Parameter(Mandatory)][string]$FixtureDirectory, [Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
$fixtures = (Resolve-Path -LiteralPath $FixtureDirectory).Path
$application = $null
$document = $null
$fields = $null
function Read-Fields($collection) {
    for ($index = 1; $index -le $collection.Count; $index++) {
        $field = $collection.Item($index)
        $code = $field.Code
        $result = $field.Result
        try { [ordered]@{ instruction = ([string]$code.Text).Trim(); locked = [bool]$field.Locked; result = [string]$result.Text } }
        finally { foreach ($com in @($result, $code, $field)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
    }
}
function Read-Dirty($owner) {
    [xml]$xml = $owner.WordOpenXML
    $namespaces = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $namespaces.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    foreach ($node in $xml.SelectNodes("//w:fldSimple | //w:fldChar[@w:fldCharType='begin']", $namespaces)) {
        if ($node.HasAttribute('dirty', $namespaces.LookupNamespace('w'))) { $node.GetAttribute('dirty', $namespaces.LookupNamespace('w')) }
        else { 'absent' }
    }
}
try {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $cases = @()
    foreach ($kind in @('simple', 'complex')) {
        foreach ($method in @('individual', 'collection')) {
            try {
                $document = $application.Documents.Open((Join-Path $fixtures "$kind.docx"), $false, $true, $false)
                $fields = $document.Fields
                $openedDirty = @(Read-Dirty $document)
                # Dirty fields may recalculate on open. Seed stale caches only after Word loads them,
                # then record actual update calls separately; no Range API is used as an update oracle.
                $opened = @(Read-Fields $fields)
                $caches = @('42', '9', 'Old title', 'Old title', '9', '9')
                if ($fields.Count -ne $caches.Count) { throw 'Unexpected generated field count.' }
                for ($index = 1; $index -le $fields.Count; $index++) {
                    $field = $fields.Item($index)
                    $result = $field.Result
                    try { $result.Text = $caches[$index - 1] }
                    finally { foreach ($com in @($result, $field)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
                }
                $before = @(Read-Fields $fields)
                $beforeDirty = @(Read-Dirty $document)
                if ($method -eq 'individual') {
                    for ($index = 1; $index -le $fields.Count; $index++) {
                        $field = $fields.Item($index)
                        try { [void]$field.Update() }
                        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($field) }
                    }
                } else { [void]$fields.Update() }
                $after = @(Read-Fields $fields)
                $afterDirty = @(Read-Dirty $document)
                $cases += [ordered]@{
                    kind = $kind; method = $method
                    instructions = @($after | ForEach-Object { $_.instruction })
                    locked = @($after | ForEach-Object { $_.locked })
                    opened = @($opened | ForEach-Object { $_.result })
                    before = @($before | ForEach-Object { $_.result })
                    after = @($after | ForEach-Object { $_.result })
                    openedDirty = $openedDirty; beforeDirty = $beforeDirty; afterDirty = $afterDirty
                }
            } finally {
                if ($fields) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($fields); $fields = $null }
                if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null }
            }
        }
    }
    [ordered]@{
        wordVersion = [string]$application.Version
        wordBuild = [string]$application.Build
        recordedDate = (Get-Date).ToString('yyyy-MM-dd')
        scope = 'installed perpetual Word; not current Microsoft 365 subscription certification'
        generator = 'bun scripts/generate-word-field-locks.mjs <directory>; fieldLockFixture defines 6 dirty fields including locked SEQ cache 42'
        recorder = 'scripts/record-word-field-locks.ps1 -FixtureDirectory <directory> -OutputPath <reference>; open dirty fields, seed only Field.Result.Text with original stale caches, then real Field.Update and Fields.Update on separate fresh documents'
        dirtyObservation = 'WordOpenXML omits all source dirty=true flags already after opening, before reseeding or explicit update. Their absence after update cannot establish that explicit update cleared them.'
        cases = $cases
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding utf8
} finally {
    if ($fields) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($fields) }
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($application) { $application.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
}
