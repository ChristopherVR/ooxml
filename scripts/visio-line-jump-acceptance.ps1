param([Parameter(Mandatory=$true)][string]$OutputDirectory)
# Creates only test artifacts and owns its invisible Visio instance.
$ErrorActionPreference = 'Stop'
& bun (Join-Path $PSScriptRoot 'visio-line-jump-acceptance.ts') $OutputDirectory
if ($LASTEXITCODE -ne 0) { throw 'Acceptance input generation failed.' }
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$results = @()
try {
    foreach ($name in @('off','gap','size','reverse')) {
        $doc = $app.Documents.OpenEx((Join-Path $OutputDirectory "$name.vsdx"),66)
        try {
            $svgPath = Join-Path $OutputDirectory "$name.svg"
            $doc.Pages.Item(1).Export($svgPath)
            $svg = Get-Content -LiteralPath $svgPath -Raw
            $group = [regex]::Match($svg,'(?s)<g id="shape1-[^"]*".*?</g>').Value
            $paths = @([regex]::Matches($group,'<path d="([^"]*)"') | ForEach-Object {$_.Groups[1].Value})
            if (!$paths.Count) { throw "No native paths for $name." }
            if ($name -eq 'off' -and ($paths.Count -ne 1 -or $paths[0] -match 'A' -or ([regex]::Matches($paths[0],'L').Count -ne 1))) { throw 'Off retained stale geometry.' }
            if ($name -eq 'gap' -and ($paths.Count -ne 2 -or ($paths -match 'A'))) { throw 'Gap did not reopen as two paths.' }
            if ($name -eq 'size' -and $paths[0] -notmatch 'A9 9 .* 1 ') { throw 'Size did not reopen with a nine-point arc.' }
            if ($name -eq 'reverse' -and $paths[0] -notmatch 'A3(?:\.\d+)? 3(?:\.\d+)? -?\d+ [01] 0 ') { throw 'Reverse did not reopen with reversed sweep.' }
            $results += [ordered]@{name=$name;paths=$paths}
        } finally { $doc.Saved=$true; $doc.Close() }
    }
    [ordered]@{source="Microsoft Visio $($app.Version) COM SVG export, $(Get-Date -Format yyyy-MM-dd)";cases=$results} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
    $results | ConvertTo-Json -Depth 4
} finally {
    $app.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
}
