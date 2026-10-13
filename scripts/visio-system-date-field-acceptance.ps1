param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory|Out-Null
$outDir=(Resolve-Path -LiteralPath $OutputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 if($app.Documents.Count -ne 0){throw 'Expected an owned empty Visio instance.'}
 $doc=$app.Documents.Add('');$cases=@();$masters=@{}
 foreach($code in @(200,211,212)) {
  $master=$doc.Masters.Add();$master.NameU='System date '+$code
  $copy=$master.Open();$shape=$copy.DrawRectangle(0,0,3,0.4)
  $shape.AddNamedRow(242,'DateValue',0)|Out-Null
  $shape.CellsU('User.DateValue').FormulaU='DATE(2026,1,1)+TIME(1,2,3)'
  $shape.Characters.AddCustomFieldU('User.DateValue',$code);$copy.Close()
  $masters[$code]=$master
 }
 foreach($code in @(200,211,212)) {
  foreach($formula in @('DATE(2026,10,13)+TIME(14,15,16)','DATE(2024,2,29)','DATE(1900,2,28)+TIME(0,0,1)','DATE(1900,3,1)+TIME(12,0,0)','DATE(1899,12,30)','DATE(1899,12,29)','DATE(2026,10,13)+TIME(23,59,59.999)','DATE(2026,12,31)+TIME(23,59,59)','DATE(2000,2,29)+TIME(1,2,3)','DATE(2100,3,1)+TIME(9,5,6)')) {
   $instance=$doc.Pages.Item(1).Drop($masters[$code],2,1+$cases.Count*0.5)
   $instance.CellsU('User.DateValue').FormulaU=$formula
   $cases += [ordered]@{id=$instance.ID;code=$code;formula=$formula;text=$instance.Characters.Text;valueFormula=$instance.CellsU('Fields.Value').FormulaU;value=$instance.CellsU('Fields.Value').ResultStrU('')}
  }
 }
 $doc.SaveAs((Join-Path $outDir 'field-system-date-native.vsdx'))|Out-Null
 $evidence=[ordered]@{version=$app.Version;culture=[Globalization.CultureInfo]::CurrentCulture.Name;shortDate=[Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern;cases=$cases}
 [IO.File]::WriteAllText((Join-Path $outDir 'field-system-date-native.json'),($evidence|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
 $cases|ConvertTo-Json -Depth 4
} finally {
 foreach($doc in @($app.Documents)){$doc.Saved=$true;$doc.Close()}
 $app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null
}
