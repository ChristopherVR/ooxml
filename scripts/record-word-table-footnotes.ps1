param([string]$OutputDirectory = (Join-Path $env:TEMP ('ooxml-native-word-table-notes-' + [guid]::NewGuid().ToString('N'))), [string[]]$CaseNames = @())
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$cases = @(
 @{name='nested-split-early';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;nested=$true;nestedLines=20},
 @{name='nested-split-late';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=15;noteLines=3;nested=$true;nestedLines=20},
 @{name='nested-second-cell-late';rows=1;cellLines=1;noteRow=0;noteCell=1;noteLine=15;noteLines=3;nested=$true;nestedLines=20},
 @{name='padded-split-early';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=3;padding=4},
 @{name='padded-split-late';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=15;noteLines=3;padding=4},
 @{name='vertical-merge-reference';rows=2;cellLines=12;noteRow=1;noteCell=0;noteLine=2;noteLines=3;merge=$true},
 @{name='nested-table-reference';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;nested=$true},
 @{name='short-table';rows=2;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3},
 @{name='partial-late-row';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;opening=15},
 @{name='late-table-row-group';rows=15;cellLines=1;noteRow=14;noteCell=0;noteLine=0;noteLines=3},
 @{name='late-row';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;opening=17},
 @{name='split-late-reference';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=15;noteLines=3},
 @{name='split-early-reference';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=3},
 @{name='split-long-note';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=80},
 @{name='split-late-long-note';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=15;noteLines=80},
 @{name='mode15-split-long-note-columns';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=80;columns=2},
 @{name='mode14-split-long-note-columns';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=80;columns=2;mode=14},
 @{name='mode12-split-long-note-columns';rows=1;cellLines=20;noteRow=0;noteCell=0;noteLine=0;noteLines=80;columns=2;mode=12},
 @{name='padded-long-note';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=80;padding=4},
 @{name='minimum-height-long-note';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=80;minimumHeight=60},
 @{name='long-note';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=80},
 @{name='keep-row';rows=1;cellLines=3;noteRow=0;noteCell=0;noteLine=1;noteLines=3;opening=14;cantSplit=$true},
 @{name='exact-clipped-reference';rows=1;cellLines=4;noteRow=0;noteCell=0;noteLine=3;noteLines=3;exact=$true},
 @{name='repeated-header';rows=2;cellLines=30;noteRow=0;noteCell=0;noteLine=0;noteLines=3;header=$true},
 @{name='two-cell-notes';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;twoNotes=$true},
 @{name='mode15-two-column-table';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;columns=2},
 @{name='mode14-two-column-table';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;columns=2;mode=14},
 @{name='mode12-two-column-table';rows=1;cellLines=1;noteRow=0;noteCell=0;noteLine=0;noteLines=3;columns=2;mode=12}
)
if($CaseNames.Count){$cases=@($cases | Where-Object {$CaseNames -contains $_.name})}
$app=$null;$owned=$false;$records=@()
try {
 $app=New-Object -ComObject Word.Application
 $owned=$app.Documents.Count -eq 0
 if(-not $owned){throw 'New Word application contains documents; refusing to touch them.'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases) {
  $doc=$null
  try {
   $doc=$app.Documents.Add()
   $doc.PageSetup.PageWidth=400;$doc.PageSetup.PageHeight=300
   $doc.PageSetup.TopMargin=40;$doc.PageSetup.BottomMargin=40
   $doc.PageSetup.LeftMargin=40;$doc.PageSetup.RightMargin=40
   if($case.columns){$doc.PageSetup.TextColumns.SetCount($case.columns);$doc.PageSetup.TextColumns.Spacing=20}
   $opening=if($case.opening){(0..($case.opening-1) | ForEach-Object {'Opening'+$_.ToString('D2')}) -join ([string][char]11)}else{''}
   if($case.columns){$opening='Left body'+"`r"+[string][char]14}
   $doc.Content.Text=$opening+$(if($opening){"`r"}else{''})
   $range=$doc.Range($doc.Content.End-1,$doc.Content.End-1)
   $table=$doc.Tables.Add($range,$case.rows,2)
   $table.AllowAutoFit=$false;$table.PreferredWidthType=3
   $table.PreferredWidth=if($case.columns){150}else{320}
   $table.Columns.SetWidth($table.PreferredWidth/2,0)
   $table.TopPadding=0;$table.BottomPadding=0;$table.LeftPadding=0;$table.RightPadding=0
   $table.Borders.Enable=0
   if($case.padding){$table.TopPadding=$case.padding;$table.BottomPadding=$case.padding}
   if($case.minimumHeight){$table.Rows.HeightRule=1;$table.Rows.Height=$case.minimumHeight}
   for($row=0;$row-lt$case.rows;$row++) {
    $count=if($case.header-and$row-eq0){1}else{$case.cellLines}
    for($cell=0;$cell-lt2;$cell++) {
     $lines=0..($count-1) | ForEach-Object {
      $text='T0R'+$row+'C'+$cell+'L'+$_.ToString('D2')
      if(-not$case.nested-and$row-eq$case.noteRow-and$cell-eq$case.noteCell-and$_-eq$case.noteLine){$text+=' REF0'}
      if($case.twoNotes-and$row-eq0-and$cell-eq1-and$_-eq0){$text+=' REF1'}
      $text
     }
     $table.Cell($row+1,$cell+1).Range.Text=$lines -join ([string][char]11)
    }
   }
   if($case.merge){$table.Cell(1,1).Merge($table.Cell(2,1))}
   if($case.nested){
    $range=$table.Cell(1,1).Range.Duplicate
    $range.Text='Outer cell'
    $range=$table.Cell(1,1).Range.Duplicate;$range.Collapse(1)
    $inner=$doc.Tables.Add($range,1,2)
    $inner.AllowAutoFit=$false;$inner.Columns.SetWidth(80,0)
    $inner.TopPadding=0;$inner.BottomPadding=0;$inner.LeftPadding=0;$inner.RightPadding=0;$inner.Borders.Enable=0
    $innerCount=if($case.nestedLines){$case.nestedLines}else{1}
    foreach($innerCell in 0..1){
     $innerLines=0..($innerCount-1)|ForEach-Object {'T1R0C'+$innerCell+'L'+$_.ToString('D2')+$(if($innerCell-eq$case.noteCell-and$_-eq$case.noteLine){' REF0'}else{''})}
     $inner.Cell(1,$innerCell+1).Range.Text=$innerLines -join ([string][char]11)
    }
   }
   $table.Rows.AllowBreakAcrossPages=-not$case.cantSplit
   if($case.header){$table.Rows.Item(1).HeadingFormat=$true}
   if($case.exact){$table.Rows.Item(1).HeightRule=2;$table.Rows.Item(1).Height=12}
   $doc.Content.Font.Name='Times New Roman';$doc.Content.Font.Size=10
   $doc.Content.ParagraphFormat.SpaceBefore=0;$doc.Content.ParagraphFormat.SpaceAfter=0
   $doc.Content.ParagraphFormat.LineSpacingRule=4;$doc.Content.ParagraphFormat.LineSpacing=12
   $doc.Content.ParagraphFormat.WidowControl=0
   $notes=if($case.twoNotes){2}else{1}
   foreach($number in 0..($notes-1)) {
    $range=$doc.Content.Duplicate
    if(-not$range.Find.Execute('REF'+$number)){throw 'Missing synthetic cell reference'}
    $range.Text=''
    $text=(0..($case.noteLines-1) | ForEach-Object {'N'+$number+'L'+$_.ToString('D2')+' source'}) -join "`r"
    $note=$doc.Footnotes.Add($range,[Type]::Missing,$text)
    $note.Range.Font.Name='Times New Roman';$note.Range.Font.Size=10
    $note.Range.ParagraphFormat.SpaceBefore=0;$note.Range.ParagraphFormat.SpaceAfter=0
    $note.Range.ParagraphFormat.LineSpacingRule=4;$note.Range.ParagraphFormat.LineSpacing=12
   }
   foreach($separator in @($doc.Footnotes.Separator,$doc.Footnotes.ContinuationSeparator)) {
    $separator.Font.Name='Times New Roman';$separator.Font.Size=10
    $separator.ParagraphFormat.SpaceBefore=0;$separator.ParagraphFormat.SpaceAfter=0
    $separator.ParagraphFormat.LineSpacingRule=4;$separator.ParagraphFormat.LineSpacing=12
   }
   $doc.Footnotes.Separator.InsertBefore('Sources ');$doc.Footnotes.ContinuationSeparator.InsertBefore('Continued ')
   $docPath=Join-Path $OutputDirectory ($case.name+'.docx')
   if($case.mode){$doc.SaveAs2($docPath,16,$false,'',$false,'',$false,$false,$false,$false,$false,65001,$false,$false,0,$false,[int]$case.mode)}else{$doc.SaveAs2($docPath,16)}
   $doc.Repaginate()
   $doc.ExportAsFixedFormat((Join-Path $OutputDirectory ($case.name+'.pdf')),17)
   $records+=@{name=$case.name;noteCount=$notes;linesPerNote=$case.noteLines;compatibilityMode=$doc.CompatibilityMode;nativePages=$doc.ComputeStatistics(2);wordVersion=$app.Version;wordBuild=$app.Build;references=@(1..$notes | ForEach-Object {@{id='N'+($_-1);page=$doc.Footnotes.Item($_).Reference.Information(3)}})}
   Write-Output ('Recorded '+$case.name)
  } finally {if($null-ne$doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 ConvertTo-Json -InputObject @($records) -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'native-cases.json')
 Write-Output $OutputDirectory
} finally {
 if($null-ne$app-and$owned){$app.Quit(0)}
 if($null-ne$app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
