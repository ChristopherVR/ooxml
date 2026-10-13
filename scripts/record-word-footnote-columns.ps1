param([string]$OutputDirectory = (Join-Path $env:TEMP ('ooxml-native-word-footnote-columns-' + [guid]::NewGuid().ToString('N'))), [string[]]$CaseNames = @())
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$cases = @(
 @{name='default-two-short';body=2;note=0;lines=1;notes=1},
 @{name='one-two-short';body=2;note=1;lines=1;notes=1},
 @{name='two-two-short';body=2;note=2;lines=1;notes=1},
 @{name='three-two-short';body=2;note=3;lines=1;notes=1},
 @{name='default-two-twenty';body=2;note=0;lines=20;notes=1},
 @{name='one-two-twenty';body=2;note=1;lines=20;notes=1},
 @{name='two-two-twenty';body=2;note=2;lines=20;notes=1},
 @{name='three-two-twenty';body=2;note=3;lines=20;notes=1},
 @{name='default-three-twenty';body=3;note=0;lines=20;notes=1},
 @{name='two-one-twenty';body=1;note=2;lines=20;notes=1},
 @{name='unequal-two-twenty';body=2;note=0;lines=20;notes=1;unequal=$true},
 @{name='default-two-eighty';body=2;note=0;lines=80;notes=1},
 @{name='two-two-eighty';body=2;note=2;lines=80;notes=1},
 @{name='three-two-eighty';body=2;note=3;lines=80;notes=1},
 @{name='two-two-multiple';body=2;note=2;lines=12;notes=3},
 @{name='two-two-custom';body=2;note=2;lines=20;notes=1;custom=$true},
 @{name='two-two-custom-eighty';body=2;note=2;lines=80;notes=1;custom=$true},
 @{name='mode14-two-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=14},
 @{name='mode12-two-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=12}

 @{name='mode14-explicit-two-eighty';body=2;note=2;lines=80;notes=1;custom=$true;mode=14},
 @{name='mode12-explicit-two-eighty';body=2;note=2;lines=80;notes=1;custom=$true;mode=12},
 @{name='mode14-two-custom-twenty';body=2;note=0;lines=20;notes=1;custom=$true;mode=14},
 @{name='mode12-two-custom-twenty';body=2;note=0;lines=20;notes=1;custom=$true;mode=12},
 @{name='mode14-left-custom-sixty';body=2;note=0;lines=60;notes=1;custom=$true;mode=14;left=$true},
 @{name='mode12-left-custom-sixty';body=2;note=0;lines=60;notes=1;custom=$true;mode=12;left=$true},
 @{name='mode14-left-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=14;left=$true},
 @{name='mode12-left-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=12;left=$true},
 @{name='mode14-unequal-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=14;unequal=$true},
 @{name='mode12-unequal-custom-eighty';body=2;note=0;lines=80;notes=1;custom=$true;mode=12;unequal=$true},
 @{name='mode14-two-standard-eighty';body=2;note=0;lines=80;notes=1;mode=14},
 @{name='mode12-two-standard-eighty';body=2;note=0;lines=80;notes=1;mode=12}
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
   $doc.PageSetup.PageWidth=400;$doc.PageSetup.PageHeight=500
   $doc.PageSetup.TopMargin=40;$doc.PageSetup.BottomMargin=40
   $doc.PageSetup.LeftMargin=40;$doc.PageSetup.RightMargin=40
   $doc.PageSetup.TextColumns.SetCount($case.body)
   if($case.body -gt 1){$doc.PageSetup.TextColumns.Spacing=20}
   if($case.unequal){$doc.PageSetup.TextColumns.EvenlySpaced=$false;$doc.PageSetup.TextColumns.Item(1).Width=100;$doc.PageSetup.TextColumns.Item(2).Width=200}
   $doc.Content.Text='Left body text.' + "`r" + ([string][char]14)*($case.body-1) + 'Reference body ' + ((0..($case.notes-1) | ForEach-Object {'REF'+$_}) -join ' ') + '.'
   if($case.left){$doc.Content.Text='Reference body ' + ((0..($case.notes-1) | ForEach-Object {'REF'+$_}) -join ' ') + '.' + "`r" + ([string][char]14)*($case.body-1) + 'Right body text.'}
   $doc.Content.Font.Name='Times New Roman';$doc.Content.Font.Size=10
   $doc.Content.ParagraphFormat.SpaceBefore=0;$doc.Content.ParagraphFormat.SpaceAfter=0
   $doc.Content.ParagraphFormat.LineSpacingRule=4;$doc.Content.ParagraphFormat.LineSpacing=12
   $doc.Sections.Item(1).Range.FootnoteOptions.LayoutColumns=$case.note
   foreach($number in 0..($case.notes-1)) {
    $range=$doc.Content.Duplicate
    if(-not $range.Find.Execute('REF'+$number)){throw 'Missing synthetic reference'}
    $range.Text=''
    $text=(0..($case.lines-1) | ForEach-Object {'N'+$number+'L'+$_.ToString('D2')+' source'}) -join "`r"
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
   if($case.custom){$doc.Footnotes.Separator.InsertBefore('Sources ');$doc.Footnotes.ContinuationSeparator.InsertBefore('Continued ')}
   $doc.Repaginate()
   $docPath=Join-Path $OutputDirectory ($case.name+'.docx')
   if($case.mode){$doc.SaveAs2($docPath,16,$false,'',$false,'',$false,$false,$false,$false,$false,65001,$false,$false,0,$false,[int]$case.mode)}else{$doc.SaveAs2($docPath,16)}
   $doc.Repaginate()
   $doc.ExportAsFixedFormat((Join-Path $OutputDirectory ($case.name+'.pdf')),17)
   $records+=@{name=$case.name;bodyColumns=$case.body;footnoteColumns=$case.note;linesPerNote=$case.lines;noteCount=$case.notes;compatibilityMode=$doc.CompatibilityMode;nativePages=$doc.ComputeStatistics(2);wordVersion=$app.Version;wordBuild=$app.Build;bodyWidths=@(1..$case.body | ForEach-Object {$doc.PageSetup.TextColumns.Item($_).Width})}
   Write-Output ('Recorded '+$case.name)
  } finally {if($null-ne$doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'native-cases.json')
 Write-Output $OutputDirectory
} finally {
 if($null-ne$app-and$owned){$app.Quit(0)}
 if($null-ne$app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
