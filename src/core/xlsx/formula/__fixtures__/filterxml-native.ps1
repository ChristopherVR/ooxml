# Regenerate the native FILTERXML corpus with desktop Excel on Windows.
$ErrorActionPreference='Stop'
$parityExcel=New-Object -ComObject Excel.Application
$parityExcel.Visible=$false
$parityExcel.DisplayAlerts=$false
$parityBook=$null
try {
 $parityBook=$parityExcel.Workbooks.Add()
 $paritySheet=$parityBook.Worksheets.Item(1)
 $paritySheet.Columns('C').ColumnWidth=60
 $parityXml='<r><a id="one">001</a><a id="two"> 2.5 </a><a>TRUE</a><a/><a>hi<b>nested</b>tail</a><a>1,234</a><a>12%</a><a>2026-10-31</a></r>'
 $parityCases=@(
 @($parityXml,'//a'),@($parityXml,'//a/@id'),@($parityXml,'//a/text()'),@($parityXml,'//a[2]'),@($parityXml,'//missing'),@($parityXml,'count(//a)'),@($parityXml,'string(//a)'),@($parityXml,'//a[1] | //a[3]'),
 @('<r><a>  hi  </a><a> </a><a><![CDATA[  CDATA  ]]></a></r>','//a'),
 @('<r xmlns="urn:x"><a>namespaced</a></r>','//a'),
 @('<r xmlns:p="urn:x"><p:a>prefixed</p:a></r>','//p:a'),
 @('<r xmlns:p="urn:x"><p:a>prefixed</p:a></r>','//*[local-name()="a"]'),
 @('<r><a>broken</r>','//a'),@('<r><a></a></r>','//a'),@('<r><a>true</a><a>FALSE</a><a>#N/A</a><a>1E3</a><a>-0.5</a></r>','//a'),
 @('<!DOCTYPE r [<!ENTITY t "expanded">]><r><a>&t;</a></r>','//a'),
 @('<r><!--comment--><?test data?><a>A</a><a>B</a></r>','//comment() | //processing-instruction()'),
 @('<r><a>A</a><a>B</a></r>','/'),@('<r><a>A</a><a>B</a></r>','//@*'),@('<r/>',''),@('','//a'),
 @('<r><a x=""/><a x="  hi  "/><a x=" TRUE "/><a x="001"/></r>','//@x'),
 @('<r><a> first <b>ignored</b> last </a><a xml:space="preserve">  hi  </a><a> hi   there </a><a>&#32; hi &#32;</a></r>','//a'),
 @('<r xmlns:p="urn:x"><p:a>prefixed</p:a><a>plain</a></r>','//a'),
 @('<r xmlns="urn:x"><a>yes</a></r>','//*[namespace-uri()="urn:x"]'),
 @('<r><a xmlns:p="urn:x"><p:b>nested binding</p:b></a></r>','//p:b'),
 @('<r><a x=unquoted>invalid</a></r>','//a'),
 @('<r><a>&unknown;</a></r>','//a'),
 @('<r><a>True</a><a> true </a><a>1,23</a><a>1/2/2026</a><a>12:30</a><a>+5</a><a>(4)</a><a>$2.50</a></r>','//a'),
 @('<r xmlns:p="urn:u" xmlns:q="urn:u"><p:a>p</p:a><q:a>q</q:a><a xmlns:p="urn:v"><p:a>nested</p:a></a></r>','//p:a'),
 @('<r><a>  hi  </a><a>hi</a></r>','//a[.="  hi  "]'),
 @('<r><a x="  hi  "/><a x="hi"/></r>','//a[@x="  hi  "]/@x'),
 @('<r> <a> hi </a> </r>','//text()'),
 @('<r xml:space="preserve"><a>  hi  </a><a>  TRUE  </a><a> 2 </a></r>','//a'),
 @('<r><a>1/2/2026</a><a>13/2/2026</a><a>2/13/2026</a><a>1-2-2026</a><a>2026/1/2</a><a>2026-01-02</a><a>2-Jan-2026</a><a>Jan 2, 2026</a><a>01/02/26</a><a>2026-01-02T12:30:00</a><a>2026-01-02 12:30</a></r>','//a')
 )
 $parityResults=@()
 foreach($parityCase in $parityCases) {
  $paritySheet.Range('A1').Value2=$parityCase[0]
  $paritySheet.Range('B1').Value2=$parityCase[1]
  $paritySheet.Range('C1').Formula2='=FILTERXML(A1,B1)'
  $parityExcel.CalculateFull()
  $parityRange=$paritySheet.Range('C1')
  if($parityRange.HasSpill) {$parityRange=$parityRange.SpillingToRange}
  $parityValues=@()
  foreach($parityCell in $parityRange.Cells) {
   if($paritySheet.Evaluate("ISERROR($($parityCell.Address()))")) {$parityValues+=@{error=[string]$parityCell.Text}}
   else {$parityValues+=,$parityCell.Value2}
  }
  $parityResults += [ordered]@{formula="=FILTERXML(A1,B1)";xml=$parityCase[0];xpath=$parityCase[1];values=$parityValues}
 }
 $parityResults | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $PSScriptRoot 'filterxml-native.json')
 Write-Output "Recorded $($parityResults.Count) FILTERXML cases from Excel $($parityExcel.Version)"
} finally {
 if($parityBook){$parityBook.Close($false)}
 $parityExcel.Quit()
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($parityExcel)|Out-Null
}
