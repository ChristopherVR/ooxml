param([Parameter(Mandatory=$true)][string]$OutputDirectory)
# Owns its invisible Visio instance and writes only named test artifacts.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
try {
    $doc = $app.Documents.Add('')
    try {
        $shape = $doc.Pages.Item(1).DrawRectangle(1,1,3,2)
        $shape.Title = 'Native title'
        # This Visio build fails its second metadata property setter. Populate the documented
        # native User row through ShapeSheet, then verify both public property getters below.
        $shape.AddNamedRow(242, 'visAltDescription', 0) | Out-Null
        $shape.CellsU('User.visAltDescription').FormulaU = '"Native description"'
        if ([string]$shape.AlternativeText -ne 'Native description') { throw 'Native description setup failed.' }
        $page = $doc.Pages.Item(1)
        $page.Title = 'Native page title'
        $page.PageSheet.AddNamedRow(242, 'visAltDescription', 0) | Out-Null
        $page.PageSheet.CellsU('User.visAltDescription').FormulaU = '"Native page description"'
        $doc.SaveAs((Join-Path $OutputDirectory 'native.vsdx'))
    } finally { $doc.Saved=$true; $doc.Close() }
    & bun (Join-Path $PSScriptRoot 'visio-alt-text-acceptance.ts') $OutputDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Acceptance generation failed.' }
    $results = @()
    foreach ($name in @('edited','cleared')) {
        $doc = $app.Documents.OpenEx((Join-Path $OutputDirectory "$name.vsdx"),66)
        try {
            $shape = $doc.Pages.Item(1).Shapes.Item(1)
            $title = [string]$shape.Title
            $description = [string]$shape.AlternativeText
            if ($name -eq 'edited' -and ($title -ne 'Service status' -or $description -ne "Green rectangle`nAll services operational.")) { throw 'Edited native metadata differs.' }
            if ($name -eq 'cleared' -and ($title -ne '' -or $description -ne '')) { throw 'Cleared metadata remains.' }
            $pageTitle = [string]$doc.Pages.Item(1).Title
            $pageDescription = [string]$doc.Pages.Item(1).AlternativeText
            if ($name -eq 'edited' -and ($pageTitle -ne 'Process overview' -or $pageDescription -ne "Requests arrive`nthen proceed to approval.")) { throw 'Edited page metadata differs.' }
            if ($name -eq 'cleared' -and ($pageTitle -ne '' -or $pageDescription -ne '')) { throw 'Cleared page metadata remains.' }
            $results += [ordered]@{name=$name;title=$title;description=$description;pageTitle=$pageTitle;pageDescription=$pageDescription}
        } finally { $doc.Saved=$true; $doc.Close() }
    }
    [ordered]@{source="Microsoft Visio $($app.Version) COM, $(Get-Date -Format yyyy-MM-dd)";cases=$results} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
    $results | ConvertTo-Json -Depth 4
} finally {
    $app.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
}
