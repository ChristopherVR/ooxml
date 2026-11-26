param(
    [Parameter(Mandatory = $true)][string]$Path
)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Visio.InvisibleApp
try {
    $app.AlertResponse = 7
    # visOpenRO + visOpenDontList + visOpenHidden + visOpenMacrosDisabled.
    $stencil = $app.Documents.OpenEx((Resolve-Path -LiteralPath $Path).Path, 2 + 8 + 64 + 128)
    try {
        if ($stencil.Masters.Count -lt 1) { throw 'Expected at least one master.' }
        $master = $stencil.Masters.Item(1)
        $shape = $master.Shapes.Item(1)
        [pscustomobject]@{
            Visio = $app.Version
            Masters = $stencil.Masters.Count
            Master = $master.NameU
            Width = $shape.CellsU('Width').ResultIU
            Connects = $master.Connects.Count
            ConnectionPoints = $shape.RowCount(7)
            ShapeData = $shape.RowCount(243)
        }
    } finally {
        $stencil.Close()
    }
} finally {
    $app.Quit()
}
