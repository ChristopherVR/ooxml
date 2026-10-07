param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-annotations-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $sheet.Range('A1').Value2 = 2
    $sheet.Range('A2').Formula = '=B2+$J$1'
    $sheet.Range('J1').Value2 = 3
    [void]$sheet.Range('A1').AddComment('source A1')
    [void]$sheet.Range('B1').AddComment('source blank B1')
    [void]$sheet.Range('A2').AddComment('source formula A2')
    $sheet.Range('A1').Validation.Add(1,1,1,'1','5')
    $sheet.Range('B1').Validation.Add(7,1,1,'=A1>0')
    $sheet.Range('B2').Validation.Add(3,1,1,'x,y')
    $sheet.Range('A1').Validation.InputTitle = 'source input'
    $sheet.Range('A1').Validation.InputMessage = 'enter a number'
    $sheet.Range('A1').Validation.ErrorTitle = 'source error'
    $sheet.Range('A1').Validation.ErrorMessage = 'invalid number'
    $modes = [ordered]@{all=-4104;values=-4163;formulas=-4123;formats=-4122;noBorders=7;comments=-4144;validation=6}
    $cases = @()
    foreach ($mode in $modes.Keys) {
        foreach ($transpose in @($false,$true)) {
            foreach ($skip in @($false,$true)) {
                foreach ($operation in @('none','add')) {
                    [void]$sheet.Range('D4:E5').Clear()
                    $sheet.Range('D4:E5').Value2 = 9
                    foreach ($ref in @('D4','E4','D5','E5')) {
                        [void]$sheet.Range($ref).AddComment("dest $ref")
                        $sheet.Range($ref).Validation.Add(1,1,1,'0','99')
                    }
                    [void]$sheet.Range('A1:B2').Copy()
                    $op = if ($operation -eq 'add') {2} else {-4142}
                    [void]$sheet.Range('D4').PasteSpecial($modes[$mode],$op,$skip,$transpose)
                    $cells = @()
                    foreach ($ref in @('D4','E4','D5','E5')) {
                        $cell = $sheet.Range($ref)
                        $comment = $null
                        if ($null -ne $cell.Comment) { $comment = [ordered]@{author=$cell.Comment.Author;text=$cell.Comment.Text()} }
                        $validation = $null
                        try {
                            $dv = $cell.Validation
                            $type = $dv.Type
                            if ($null -eq $type) { throw 'No validation' }
                            $validation = [ordered]@{type=$type;operator=$dv.Operator;formula1=[string]$dv.Formula1;formula2=[string]$dv.Formula2
                                allowBlank=$dv.IgnoreBlank;showInputMessage=$dv.ShowInput;showErrorMessage=$dv.ShowError
                                errorStyle=$dv.AlertStyle;errorTitle=[string]$dv.ErrorTitle;error=[string]$dv.ErrorMessage
                                promptTitle=[string]$dv.InputTitle;prompt=[string]$dv.InputMessage}
                            if ($type -eq 3) {$validation.showDropDown=$dv.InCellDropdown}
                        } catch {}
                        $cells += [ordered]@{row=$cell.Row-1;col=$cell.Column-1;value=$cell.Value2;comment=$comment;validation=$validation}
                    }
                    $cases += [ordered]@{mode=$mode;transpose=$transpose;skipBlanks=$skip;operation=$operation;cells=$cells}
                }
            }
        }
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 10) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) annotation cases"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
