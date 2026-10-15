# Opens saved documents in desktop Word (Windows, manual) and reports the shapes, content controls and
# equations Word found, so a save that Word must repair or that loses objects shows up. Word is started
# hidden; a document that needs repair fails to open instead of being silently fixed.
param([Parameter(Mandatory = $true)][string[]]$Path)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$results = @()
try {
    foreach ($file in $Path) {
        $full = (Resolve-Path -LiteralPath $file).Path
        $document = $null
        try {
            # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            $document = $word.Documents.Open($full, $false, $true, $false)
            $types = @{}
            foreach ($control in $document.ContentControls) { $key = [string]$control.Type; $types[$key] = 1 + [int]$types[$key] }
            $presets = @()
            foreach ($shape in $document.Shapes) { $presets += [string]$shape.AutoShapeType }
            $results += [ordered]@{
                file = Split-Path $full -Leaf
                opened = $true
                floatingShapes = $document.Shapes.Count
                inlineShapes = $document.InlineShapes.Count
                shapeAutoTypes = $presets
                contentControls = $document.ContentControls.Count
                contentControlTypes = $types
                equations = $document.OMaths.Count
                paragraphs = $document.Paragraphs.Count
            }
        } catch {
            $results += [ordered]@{ file = Split-Path $full -Leaf; opened = $false; error = $_.Exception.Message }
        } finally {
            if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
        }
    }
} finally {
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}
$results | ConvertTo-Json -Depth 5
