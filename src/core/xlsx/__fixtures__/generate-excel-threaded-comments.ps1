# Generates a native thread with two replies through Excel 16.
$ErrorActionPreference = 'Stop'
$commentExcel = $null
$commentBook = $null
try {
	$commentExcel = New-Object -ComObject Excel.Application
	$commentExcel.Visible = $false
	$commentExcel.DisplayAlerts = $false
	$commentBook = $commentExcel.Workbooks.Add()
	$commentSheet = $commentBook.Worksheets.Item(1)
	$commentSheet.Name = 'Comments'
	$commentRoot = $commentSheet.Range('B2').AddCommentThreaded('Original comment')
	[void]$commentRoot.AddReply('First reply')
	[void]$commentRoot.AddReply('Second reply')
	$commentBook.SaveAs((Join-Path $PSScriptRoot 'excel-threaded-comments.xlsx'), 51)
} finally {
	if ($null -ne $commentBook) {
		$commentBook.Close($false)
		[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($commentBook)
	}
	if ($null -ne $commentExcel) {
		$commentExcel.Quit()
		[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($commentExcel)
	}
}
# Keep the native comment structure while removing machine/account identity from the fixture.
$commentArchive = [IO.Compression.ZipFile]::Open((Join-Path $PSScriptRoot 'excel-threaded-comments.xlsx'), 'Update')
try {
	foreach ($commentPath in @('xl/persons/person.xml', 'docProps/core.xml')) {
		$commentEntry = $commentArchive.GetEntry($commentPath)
		if ($null -eq $commentEntry) { continue }
		$commentReader = [IO.StreamReader]::new($commentEntry.Open())
		[xml]$commentXml = $commentReader.ReadToEnd()
		$commentReader.Dispose()
		foreach ($commentPerson in $commentXml.SelectNodes('//*[local-name()="person"]')) {
			$commentPerson.SetAttribute('displayName', 'Parity Author')
			$commentPerson.SetAttribute('userId', 'Parity Author')
			$commentPerson.SetAttribute('providerId', 'None')
		}
		foreach ($commentCreator in $commentXml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) {
			$commentCreator.InnerText = 'Parity Author'
		}
		$commentEntry.Delete()
		$commentWriter = [IO.StreamWriter]::new($commentArchive.CreateEntry($commentPath).Open(), [Text.UTF8Encoding]::new($false))
		$commentWriter.Write($commentXml.OuterXml)
		$commentWriter.Dispose()
	}
} finally {
	$commentArchive.Dispose()
}
