param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$RenameTo
)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Visio.InvisibleApp
try {
    $app.AlertResponse = 7
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    if ($RenameTo) {
        $drawing = $app.Documents.Add('')
        # visOpenDocked + visOpenDontList + visOpenRW + visOpenMacrosDisabled.
        $flags = 4 + 8 + 32 + 128
        $stencil = $app.Documents.OpenEx($resolved, $flags)
        try {
            if ($stencil.ReadOnly) { throw 'The stencil opened read-only.' }
            $original = $stencil.Masters.Item(1).NameU
            if ($original -eq $RenameTo) { throw 'Choose a different master name.' }
            $stencil.Masters.Item(1).Name = $RenameTo
            $stencil.Masters.Item(1).NameU = $RenameTo
            $stencil.Save()
        } finally {
            $stencil.Close()
        }
        $stencil = $app.Documents.OpenEx($resolved, $flags)
        try {
            $master = $stencil.Masters.Item(1)
            if ($master.Name -cne $RenameTo -or $master.NameU -cne $RenameTo) {
                throw 'The renamed master did not survive reopening.'
            }
            [pscustomobject]@{
                Visio = $app.Version
                ReadOnly = $stencil.ReadOnly
                OpenFlags = $flags
                Original = $original
                ReopenedName = $master.Name
                ReopenedNameU = $master.NameU
            }
        } finally {
            $stencil.Close()
            $drawing.Close()
        }
        return
    }
    # visOpenRO + visOpenDontList + visOpenHidden + visOpenMacrosDisabled.
    $stencil = $app.Documents.OpenEx($resolved, 2 + 8 + 64 + 128)
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
