# Records direct drawing-size and independent printer settings in an owned invisible Visio process.
# Application.Settings and the system clipboard are never read or changed.
param(
 [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-page-size-' + [guid]::NewGuid().ToString('N'))),
 [switch]$NoPageDependencies,
 [switch]$SizeModes,
 [string]$CoreOutputDirectory,
 [string]$FactorySourcePath
)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$ErrorActionPreference='Stop'
$names=@('PageWidth','PageHeight','PageScale','DrawingScale','DrawingScaleType','DrawingSizeType','DrawingResizeType','PrintPageOrientation','PaperKind','PaperSource','PageLeftMargin','PageRightMargin','PageTopMargin','PageBottomMargin','ScaleX','ScaleY','PagesX','PagesY')
function Read-Cells($Shape,$Names){
 $values=[ordered]@{}
 foreach($name in $Names){if($Shape.CellExistsU($name,0)){$c=$Shape.CellsU($name);$values[$name]=@{formula=$c.FormulaU;value=$c.ResultIU}}}
 return $values
}
function Read-Page($Page){
 $shapes=@()
 foreach($s in $Page.Shapes){
  $shapes+=,@{id=$s.ID;name=$s.NameU;text=$s.Text;oneD=$s.OneD;cells=(Read-Cells $s @('PinX','PinY','Width','Height','Angle','LocPinX','LocPinY','BeginX','BeginY','EndX','EndY'));connects=$s.Connects.Count}
 }
 $backPageName=if($Page.BackPage){$Page.BackPage.NameU}else{$null}
 return @{id=$Page.ID;name=$Page.NameU;background=$Page.Background;backPage=$backPageName;cells=(Read-Cells $Page.PageSheet $names);shapes=$shapes;connects=$Page.Connects.Count}
}
function Assert-Near($actual,$expected,$label) {
 if([math]::Abs([double]$actual-[double]$expected) -gt 1e-10){throw "$label differs: $actual != $expected"}
}
function Assert-Page($actual,$expected) {
 $expectedBack=if($null -eq $expected.backPage){$null}elseif($expected.backPage -is [string]){$expected.backPage}else{'Background'}
 if($actual.id -ne $expected.id -or $actual.name -ne $expected.name -or $actual.background -ne $expected.background -or $actual.backPage -ne $expectedBack -or $actual.connects -ne $expected.connects){throw 'Page identities/connections differ'}
 foreach($property in $expected.cells.PSObject.Properties){Assert-Near $actual.cells[$property.Name].value $property.Value.value $property.Name}
 if(@($actual.shapes).Count -ne @($expected.shapes).Count){throw 'Shape count differs'}
 foreach($shape in $expected.shapes){
  $found=@($actual.shapes | Where-Object {$_.id -eq $shape.id})
  if($found.Count -ne 1 -or $found[0].text -cne $shape.text -or $found[0].name -cne $shape.name -or $found[0].connects -ne $shape.connects -or $found[0].oneD -ne $shape.oneD){throw 'Shape identity/text/connections differ'}
  foreach($property in $shape.cells.PSObject.Properties){Assert-Near $found[0].cells[$property.Name].value $property.Value.value ('Shape '+$shape.id+' '+$property.Name)}
 }
}
$app=New-Object -ComObject Visio.InvisibleApp;$doc=$null;$records=@()
try{
 $app.AlertResponse=7
 $app.EventsEnabled=0
 if ($CoreOutputDirectory) {
  $evidence=Get-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Raw | ConvertFrom-Json
  foreach($item in $evidence.cases) {
   if($item.operation -notlike 'drawing-*'){continue}
   $name=[IO.Path]::GetFileName($item.beforeFile).Replace('-before.vsdx','-core.vsdx')
   $doc=$app.Documents.OpenEx((Join-Path $CoreOutputDirectory $name),202)
   $foreground=Read-Page ($doc.Pages.ItemU($item.after.foreground.name));$background=if($item.after.background){Read-Page ($doc.Pages.ItemU($item.after.background.name))}else{$null}
   Assert-Page $foreground $item.after.foreground
   if($item.after.background){Assert-Page $background $item.after.background}
   $records+=,@{scale=$item.scale;operation=$item.operation;foreground=$foreground;background=$background;matched=$true}
   $doc.Saved=$true;$doc.Close();$doc=$null
  }
  @{version=$app.Version;cases=$records;limitations='Owned core output COM reopen against direct native cell references; no interactive ribbon or Auto Size UI claim.'}|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $CoreOutputDirectory 'core-native-reopened.json') -Encoding utf8
  Write-Output ('Core page-size outputs reopened and matched: '+$records.Count)
  return
 }
 if($FactorySourcePath){
  $doc=$app.Documents.OpenEx((Resolve-Path -LiteralPath $FactorySourcePath).Path,202);$fg=$doc.Pages.Item(1)
  $beforePath=Join-Path $OutputDirectory 'scale-1-drawing-custom-before.vsdx';$afterPath=Join-Path $OutputDirectory 'scale-1-drawing-custom-after.vsdx'
  Copy-Item -LiteralPath $FactorySourcePath -Destination $beforePath
  $before=@{foreground=(Read-Page $fg);background=$null}
  $fg.PageSheet.CellsU('DrawingSizeType').FormulaU='3';$fg.PageSheet.CellsU('DrawingResizeType').FormulaU='0'
  $fg.PageSheet.CellsU('PageWidth').ResultIU=6.25;$fg.PageSheet.CellsU('PageHeight').ResultIU=4.75
  $doc.SaveAs($afterPath)|Out-Null
  $after=@{foreground=(Read-Page $fg);background=$null}
  $records+=,@{scale=1;operation='drawing-custom';beforeFile=$beforePath;afterFile=$afterPath;before=$before;after=$after}
  @{version=$app.Version;cases=$records;factorySource=$FactorySourcePath;limitations='Direct fixed/custom page-size transition from actual core factory bytes, no template or interactive UI claim.'}|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
  return
 }
 $scales=if($SizeModes){@(1)}else{@(1,.5,2)}
 $operations=if($SizeModes){foreach($size in @(0,1,2,3,4)){foreach($resize in @(0,1,2)){'drawing-mode-'+$size+'-'+$resize}}}else{@('print-landscape','printer-a4','drawing-letter','drawing-a4','drawing-custom','drawing-swap','fit')}
 foreach($scale in $scales){
  foreach($operation in $operations){
   if ($NoPageDependencies -and $operation -notlike 'drawing-*') { continue }
   $doc=$app.Documents.Add('');$fg=$doc.Pages.Item(1);$fg.NameU='Foreground';$bg=$doc.Pages.Add();$bg.NameU='Background';$bg.Background=-1
   foreach($p in @($fg,$bg)){
    $p.PageSheet.CellsU('DrawingResizeType').FormulaU='0';$p.PageSheet.CellsU('DrawingSizeType').FormulaU='3';$p.PageSheet.CellsU('DrawingScaleType').FormulaU='3'
    $p.PageSheet.CellsU('PageScale').FormulaU='1 in';$p.PageSheet.CellsU('DrawingScale').FormulaU=$scale.ToString([cultureinfo]::InvariantCulture)+' in'
    $p.PageSheet.CellsU('PageWidth').ResultIU=[double](9*$scale);$p.PageSheet.CellsU('PageHeight').ResultIU=[double](6*$scale)
   }
   $bg.PageSheet.CellsU('PageWidth').ResultIU=[double](7.5*$scale);$bg.PageSheet.CellsU('PageHeight').ResultIU=[double](5*$scale)
   if($SizeModes){
    $mode=$operation.Split('-')
    $fg.PageSheet.CellsU('DrawingSizeType').FormulaU=$mode[2]
    $fg.PageSheet.CellsU('DrawingResizeType').FormulaU=$mode[3]
   }
   $fg.BackPage='Background'
   $rotated=$fg.DrawRectangle(1,1,3,2);$rotated.NameU='Rotated';$rotated.CellsU('Angle').FormulaU='30 deg';$rotated.Text='Unchanged rotated text'
   $outside=$fg.DrawRectangle(11,-1,12,0);$outside.NameU='Outside'
   $line=$fg.DrawLine(2,1.5,11.5,-.5);$line.NameU='Connector';$line.CellsU('GlueType').FormulaU='2';$line.CellsU('BeginX').GlueTo($rotated.CellsU('PinX'));$line.CellsU('EndX').GlueTo($outside.CellsU('PinX'))
   if (-not $NoPageDependencies) { $dependent=$fg.DrawRectangle(4,3,5,4);$dependent.NameU='Dependent';$dependent.CellsU('PinX').FormulaU='ThePage!PageWidth/2';$dependent.CellsU('Width').FormulaU='ThePage!PageWidth/10' }
   $back=$bg.DrawRectangle(1,1,3,2);$back.NameU='BackgroundShape';$back.Text='Background stays independent';$back.CellsU('Width').FormulaU='ThePage!PageWidth/5'
   $prefix='scale-'+$scale.ToString([cultureinfo]::InvariantCulture)+'-'+$operation
   $beforePath=Join-Path $OutputDirectory ($prefix+'-before.vsdx');$afterPath=Join-Path $OutputDirectory ($prefix+'-after.vsdx')
   $doc.SaveAs($beforePath)|Out-Null
   $before=@{foreground=(Read-Page $fg);background=(Read-Page $bg)}
   if($SizeModes){
    $fg.PageSheet.CellsU('DrawingSizeType').FormulaU='3';$fg.PageSheet.CellsU('DrawingResizeType').FormulaU='0'
    $fg.PageSheet.CellsU('PageWidth').ResultIU=[double](6.25*$scale);$fg.PageSheet.CellsU('PageHeight').ResultIU=[double](4.75*$scale)
   }
   switch($operation){
    'print-landscape'{$fg.PageSheet.CellsU('PrintPageOrientation').FormulaU='2'}
    'printer-a4'{$fg.PageSheet.CellsU('PaperKind').FormulaU='9'}
    'drawing-letter'{$fg.PageSheet.CellsU('PageWidth').ResultIU=[double](8.5*$scale);$fg.PageSheet.CellsU('PageHeight').ResultIU=[double](11*$scale)}
    'drawing-a4'{$fg.PageSheet.CellsU('PageWidth').ResultIU=[double](210/25.4*$scale);$fg.PageSheet.CellsU('PageHeight').ResultIU=[double](297/25.4*$scale)}
    'drawing-custom'{$fg.PageSheet.CellsU('PageWidth').ResultIU=[double](6.25*$scale);$fg.PageSheet.CellsU('PageHeight').ResultIU=[double](4.75*$scale)}
    'drawing-swap'{$w=$fg.PageSheet.CellsU('PageWidth').ResultIU;$h=$fg.PageSheet.CellsU('PageHeight').ResultIU;$fg.PageSheet.CellsU('PageWidth').ResultIU=[double]$h;$fg.PageSheet.CellsU('PageHeight').ResultIU=[double]$w}
    'fit'{$fg.ResizeToFitContents()}
   }
   $doc.SaveAs($afterPath)|Out-Null
   $after=@{foreground=(Read-Page $fg);background=(Read-Page $bg)}
   $records+=,@{scale=$scale;operation=$operation;beforeFile=$beforePath;afterFile=$afterPath;before=$before;after=$after}
   $doc.Saved=$true;$doc.Close();$doc=$null
  }
 }
 @{application=$app.Name;version=$app.Version;visible=$app.Visible;cases=$records;limitations='Direct per-page native cells and ResizeToFitContents API; no GUI orientation or page-size defaults measured. Application.Settings never changed.'}|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
}finally{if($doc){$doc.Saved=$true;$doc.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)};$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
