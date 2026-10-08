# Native logical-text oracle. Creates and closes only its own invisible Visio document.
param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$ReopenCore)
$ErrorActionPreference = 'Stop'
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
try {
    $application.AlertResponse = 7
    if ($ReopenCore) {
        $cases = @(Get-Content -LiteralPath (Join-Path $directory 'core-cases.json') -Raw | ConvertFrom-Json)
        $document = $application.Documents.OpenEx((Join-Path $directory 'core.vsdx'),202)
        $observations = @()
        foreach ($case in $cases) {
            $shape = $document.Pages.Item(1).Shapes.ItemFromID([int]$case.id)
            if ($shape.Text -cne $case.text) { throw "Logical text mismatch: $($case.mode)/$($case.id)" }
            if ($shape.Characters.Text -cne $case.text) { throw "Character text mismatch: $($case.mode)/$($case.id)" }
            $observations += [ordered]@{ id=$case.id; mode=$case.mode; text=$shape.Text }
        }
        [void]$document.SaveAs((Join-Path $directory 'core-resaved.vsdx'))
        $record = 'reopen-evidence.json'
    } else {
        $nativePath = Join-Path $directory 'native.vsdx'
        if (Test-Path -LiteralPath $nativePath) { throw 'Choose a new oracle output directory.' }
        $inputs = @('',"`n","`n`n","A`n","A`n`n",'Hello',"Line 1`nLine 2","😀`nLast`n"," spaces `t ",'<literal> & "quotes"')
        $document = $application.Documents.Add('')
        $observations = @()
        foreach ($value in $inputs) {
            $shape = $document.Pages.Item(1).DrawRectangle(1,2,3,3)
            $shape.Text = $value
            $observations += [ordered]@{ id=[string]$shape.ID; text=$value }
        }
        [void]$document.SaveAs($nativePath)
        $document.Saved = $true
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        $document = $application.Documents.OpenEx($nativePath,202)
        foreach ($case in $observations) {
            $shape = $document.Pages.Item(1).Shapes.ItemFromID([int]$case.id)
            if ($shape.Text -cne $case.text) { throw "Native saved text differs: $($case.id)" }
        }
        $record = 'native-evidence.json'
    }
    [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$observations } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory $record) -Encoding utf8
    Write-Output $directory
} finally {
    if ($null -ne $document) {
        $document.Saved = $true
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
