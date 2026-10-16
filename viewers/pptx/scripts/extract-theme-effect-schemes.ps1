param([string]$out)
New-Item -ItemType Directory -Force $out | Out-Null
Add-Type -A System.IO.Compression.FileSystem
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Add(0)
$s = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(2))
$p.SaveCopyAs("$out\Office.pptx")
$dir = 'C:\Program Files\Microsoft Office\root\Document Themes 16\Theme Effects'
foreach ($f in Get-ChildItem $dir -Filter *.eftx) {
  $p.SlideMaster.Theme.ThemeEffectScheme.Load($f.FullName)
  $p.SaveCopyAs("$out\" + $f.BaseName + ".pptx")
}
$p.Saved = -1
$p.Close()
foreach ($f in Get-ChildItem $out -Filter *.pptx) {
  $z = [IO.Compression.ZipFile]::OpenRead($f.FullName)
  $e = $z.GetEntry('ppt/theme/theme1.xml')
  $r = New-Object IO.StreamReader($e.Open()); $t = $r.ReadToEnd(); $r.Close()
  $has = $z.Entries | ? { $_.FullName -like 'ppt/theme/_rels/*' -or $_.FullName -like 'ppt/media/*' }
  $z.Dispose()
  $i = $t.IndexOf('<a:fmtScheme'); $j = $t.IndexOf('</a:fmtScheme>') + 14
  [IO.File]::WriteAllText("$out\" + $f.BaseName + ".fmt.xml", $t.Substring($i, $j - $i), (New-Object Text.UTF8Encoding($false)))
  $f.BaseName + ' len=' + ($j - $i) + ' extra=' + (($has | % { $_.FullName }) -join ',')
}
