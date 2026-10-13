param([Parameter(Mandatory)][string]$Directory, [switch]$Render, [int]$ExpectedEffectCount = 4)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint process before this owned fixture check.' }
$folder = (Get-Item -LiteralPath $Directory).FullName
$app = $null
$presentation = $null
$records = @()
function Read-PackageSound([string]$Path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $media = @{}
        foreach ($entry in $zip.Entries | Where-Object { $_.FullName.StartsWith('ppt/media/') -and $_.Length -gt 0 }) {
            $stream = $entry.Open()
            try { $media[$entry.FullName] = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
            finally { $stream.Dispose() }
        }
        $reader = [IO.StreamReader]::new($zip.GetEntry('ppt/slides/_rels/slide1.xml.rels').Open())
        try { [xml]$rels = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $audio = @($rels.Relationships.Relationship | Where-Object { $_.Type.EndsWith('/audio') } | ForEach-Object { $_.Id+':'+$_.Target })
        return @{ Media=$media; Audio=$audio }
    } finally { $zip.Dispose() }
}
$sourceSound = Read-PackageSound (Join-Path $folder 'source.pptx')
try {
    $app = New-Object -ComObject PowerPoint.Application
    $cases = @(@{ Name='source'; ClearShape=0 }, @{ Name='dimToColor-cleared'; ClearShape=2 }, @{ Name='dimScheme-cleared'; ClearShape=3 }, @{ Name='hideAfterAnimation-cleared'; ClearShape=4 }, @{ Name='hideOnNextClick-cleared'; ClearShape=5 })
    foreach ($case in $cases) {
        $path = (Get-Item -LiteralPath (Join-Path $folder ($case.Name+'.pptx'))).FullName
        $sound = Read-PackageSound $path
        foreach ($name in $sourceSound.Media.Keys) {
            if ($sound.Media[$name] -ne $sourceSound.Media[$name]) { throw 'Package media bytes changed.' }
        }
        if (($sound.Audio -join '|') -ne ($sourceSound.Audio -join '|')) { throw 'Package audio relationships changed.' }
        $presentation = $app.Presentations.Open($path, $false, $false, $false)
        $effects = $presentation.Slides.Item(1).TimeLine.MainSequence
        if ($effects.Count -ne $ExpectedEffectCount) { throw 'Native effect count changed.' }
        $states = @()
        for ($i=1; $i -le $effects.Count; $i++) {
            $effect = $effects.Item($i)
            $states += @{ ShapeId=$effect.Shape.Id; AfterEffect=$effect.EffectInformation.AfterEffect; EffectType=$effect.EffectType; SoundType=$effect.EffectInformation.SoundEffect.Type }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($effect)
        }
        if ($case.ClearShape -ne 0 -and ($states | Where-Object ShapeId -eq $case.ClearShape).AfterEffect -ne 0) { throw 'Native after-effect clear was not recognized.' }
        $saved = Join-Path $folder ($case.Name+'-native.pptx')
        $presentation.SaveCopyAs($saved)
        if ($Render -and $case.Name -in 'source','dimToColor-cleared') {
            $video = Join-Path $folder ($case.Name+'.mp4')
            $presentation.CreateVideo($video, $true, 12, 540, 10, 85)
            $deadline = [DateTime]::UtcNow.AddSeconds(90)
            while ($presentation.CreateVideoStatus -ne 3 -and [DateTime]::UtcNow -lt $deadline) {
                if ($presentation.CreateVideoStatus -eq 4) { throw 'Native video export failed.' }
                Start-Sleep -Milliseconds 250
            }
            if ($presentation.CreateVideoStatus -ne 3) { throw 'Native video export timed out.' }
        }
        $presentation.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
        $presentation = $null
        $presentation = $app.Presentations.Open($saved, $false, $false, $false)
        $reopened = $presentation.Slides.Item(1).TimeLine.MainSequence
        if ($reopened.Count -ne $ExpectedEffectCount) { throw 'Native reopened effect count changed.' }
        for ($i=1; $i -le $reopened.Count; $i++) {
            $effect = $reopened.Item($i)
            $before = $states | Where-Object ShapeId -eq $effect.Shape.Id
            if ($effect.EffectInformation.AfterEffect -ne $before.AfterEffect) { throw 'Native reopened after-state changed.' }
            if ($effect.EffectInformation.SoundEffect.Type -ne $before.SoundType) { throw 'Native reopened sound changed.' }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($effect)
        }
        $records += @{ Case=$case.Name; Effects=$states; NativeSaved=$saved; NativeReopened=$true; PackageSound=$sound }
        $presentation.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
        $presentation = $null
    }
    [IO.File]::WriteAllText((Join-Path $folder 'native-after-effect-evidence.json'), ($records | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
} finally {
    if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
