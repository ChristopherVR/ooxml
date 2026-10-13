param([string]$OutputDirectory = (Join-Path $env:TEMP ('ooxml-visio-print-inheritance-'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
$out=$OutputDirectory;New-Item -ItemType Directory -Path $out|Out-Null
$names=@('PaperKind','PrintPageOrientation','ScaleX','ScaleY','OnPage','CenterX','CenterY','PagesX','PagesY','PageLeftMargin','PageRightMargin','PageTopMargin','PageBottomMargin','PageWidth','PageHeight','PageScale','DrawingScale')
function Read-Cells($sheet){$r=[ordered]@{};foreach($n in $names){$c=$sheet.CellsU($n);$r[$n]=@{formula=$c.FormulaU;value=$c.ResultIU}};return $r}
$cases=@(@{id='root-paper';target='root';values=@{PaperKind='9'}},@{id='document-paper';target='document';values=@{PaperKind='9'}},@{id='page-paper';target='page';values=@{PaperKind='9'}},@{id='root-fit';target='root';values=@{OnPage='1'}},@{id='document-fit';target='document';values=@{OnPage='1'}},@{id='root-zoom';target='root';values=@{ScaleX='.5';ScaleY='.5'}},@{id='document-zoom';target='document';values=@{ScaleX='.5';ScaleY='.5'}})
$app=New-Object -ComObject Visio.InvisibleApp;$records=@()
try{
 $app.AlertResponse=7;$app.EventsEnabled=0
 foreach($case in $cases){
  $doc=$app.Documents.Add('');$page=$doc.Pages.Item(1);$sheet=$page.PageSheet;$root=$doc.Styles.ItemFromID(0)
  foreach($n in @('DrawingResizeType','DrawingSizeType','PageScale','DrawingScale','PageWidth','PageHeight','PageLeftMargin','PageRightMargin','PageTopMargin','PageBottomMargin')){$value=switch($n){DrawingResizeType{'0'}DrawingSizeType{'3'}PageWidth{'17 in'}PageHeight{'22 in'}PageScale{'1 in'}DrawingScale{'1 in'}default{'.25 in'}};$sheet.CellsU($n).FormulaU=$value}
  $target=switch($case.target){root{$root}document{$doc.DocumentSheet}page{$sheet}}
  $writeRefusal=$null;try{foreach($n in $case.values.Keys){$target.CellsU($n).FormulaU=$case.values[$n]}}catch{$writeRefusal=$_.Exception.Message;Write-Output ($case.id+': native write refused')}
  $native=Join-Path $out ($case.id+'.vsdx');$doc.SaveAs($native)|Out-Null
  $count=$page.PrintTileCount
  $records+=,@{id=$case.id;target=$case.target;writeRefusal=$writeRefusal;requested=$case.values;resolved=(Read-Cells $sheet);count=$count;native=$native}
  $records|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $out 'cases.json') -Encoding utf8
  Write-Output ($case.id+': paper '+$sheet.CellsU('PaperKind').ResultIU+', fit '+$sheet.CellsU('OnPage').ResultIU+', zoom '+$sheet.CellsU('ScaleX').ResultIU+', count '+$count)
 }
 Write-Output ('Evidence: '+$out)
}finally{$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
