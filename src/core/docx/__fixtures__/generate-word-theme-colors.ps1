# Records Word's cached theme colours for native text and page backgrounds.
$ErrorActionPreference='Stop'
$parityWord=New-Object -ComObject Word.Application
$parityWord.Visible=$false
$parityWord.DisplayAlerts=0
$parityDoc=$null
$parityPath=Join-Path ([IO.Path]::GetTempPath()) ("ooxml-word-colors-"+[guid]::NewGuid().ToString('N')+'.docx')
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-ParityPart($path,$name) {
 $archive=[IO.Compression.ZipFile]::OpenRead($path)
 try {
  $reader=[IO.StreamReader]::new($archive.GetEntry($name).Open())
  try {return [xml]$reader.ReadToEnd()} finally {$reader.Dispose()}
 } finally {$archive.Dispose()}
}
function Read-ParityColor($element,$ns) {
 $record=[ordered]@{themeColor=$element.GetAttribute('themeColor',$ns);expected=$element.GetAttribute('color',$ns)}
 if(-not $record.expected){$record.expected=$element.GetAttribute('val',$ns)}
 foreach($name in @('themeTint','themeShade')) {
  $value=$element.GetAttribute($name,$ns)
  if($value){$record[$name]=$value}
 }
 return $record
}
try {
 $parityCases=@()
 foreach($theme in 0..9) {
  foreach($tint in @(0,0.2,0.4,0.5,0.6,0.8,-0.25,-0.5,-0.75)) {
   $parityCases+=@{theme=$theme;tint=$tint}
  }
 }
 $parityDoc=$parityWord.Documents.Add()
 $parityDoc.Content.Text=($parityCases|ForEach-Object {"Theme $($_.theme) tint $($_.tint)"}) -join "`r"
 for($index=0;$index -lt $parityCases.Count;$index++) {
  $range=$parityDoc.Paragraphs.Item($index+1).Range
  $range.Font.TextColor.ObjectThemeColor=$parityCases[$index].theme
  $range.Font.TextColor.TintAndShade=[single]$parityCases[$index].tint
 }
 $parityDoc.SaveAs2($parityPath,16)
 $parityDoc.Close(0)
 $parityDoc=$null
 $wordNs='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
 $drawingNs='http://schemas.openxmlformats.org/drawingml/2006/main'
 $xml=Read-ParityPart $parityPath 'word/document.xml'
 $themeXml=Read-ParityPart $parityPath 'word/theme/theme1.xml'
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'word-theme-native.xml'),$themeXml.OuterXml,[Text.UTF8Encoding]::new($false))
 $themeColors=[ordered]@{}
 foreach($slot in $themeXml.GetElementsByTagName('clrScheme',$drawingNs)[0].ChildNodes) {
  $color=$slot.FirstChild
  $value=$color.GetAttribute('val')
  if($color.LocalName -eq 'sysClr'){$value=$color.GetAttribute('lastClr')}
  $themeColors[$slot.LocalName]=$value
 }
 $colors=@()
 foreach($paragraph in $xml.GetElementsByTagName('p',$wordNs)) {
  $color=$paragraph.GetElementsByTagName('color',$wordNs)[0]
  if($color){$colors+=Read-ParityColor $color $wordNs}
 }
 $backgrounds=@()
 foreach($tint in @(0,0.5,-0.5)) {
  $parityDoc=$parityWord.Documents.Open($parityPath,$false,$false)
  $parityDoc.Background.Fill.Solid()
  $parityDoc.Background.Fill.ForeColor.ObjectThemeColor=5
  $parityDoc.Background.Fill.ForeColor.TintAndShade=[single]$tint
  $parityDoc.Save()
  $parityDoc.Close(0)
  $parityDoc=$null
  $xml=Read-ParityPart $parityPath 'word/document.xml'
  $backgrounds+=Read-ParityColor $xml.GetElementsByTagName('background',$wordNs)[0] $wordNs
 }
 $roundtripColors=@()
 foreach($paragraph in $xml.GetElementsByTagName('p',$wordNs)) {
  $color=$paragraph.GetElementsByTagName('color',$wordNs)[0]
  if($color){$roundtripColors+=Read-ParityColor $color $wordNs}
 }
 if(($colors|ConvertTo-Json -Depth 4) -ne ($roundtripColors|ConvertTo-Json -Depth 4)){throw 'Native theme colours changed on reopen.'}
 $result=[ordered]@{source="Microsoft Word $($parityWord.Version)";theme=$themeColors;colors=$colors;backgrounds=$backgrounds}
 $result|ConvertTo-Json -Depth 6|Set-Content (Join-Path $PSScriptRoot 'word-theme-colors.json')
 Write-Output "Recorded $($colors.Count) text colours and $($backgrounds.Count) backgrounds."
} finally {
 if($parityDoc){$parityDoc.Close(0)}
 $parityWord.Quit()
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($parityWord)|Out-Null
 if(Test-Path -LiteralPath $parityPath){Remove-Item -LiteralPath $parityPath}
}
