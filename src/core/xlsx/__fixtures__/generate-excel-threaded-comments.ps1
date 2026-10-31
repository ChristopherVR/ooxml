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
	$commentSheet.Range('B2').Value2 = 'Threaded comment cell'
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
$commentPersonIds = @{}
$commentPersonNumber = 0
try {
	foreach ($commentPath in @('xl/persons/person.xml', 'docProps/core.xml')) {
		$commentEntry = $commentArchive.GetEntry($commentPath)
		if ($null -eq $commentEntry) { continue }
		$commentReader = [IO.StreamReader]::new($commentEntry.Open())
		[xml]$commentXml = $commentReader.ReadToEnd()
		$commentReader.Dispose()
		foreach ($commentPerson in $commentXml.SelectNodes('//*[local-name()="person"]')) {
			$commentPersonNumber += 1
			$commentPersonId = '{AAAAAAAA-0000-4000-8000-' + $commentPersonNumber.ToString('D12') + '}'
			$commentPersonIds[$commentPerson.GetAttribute('id')] = $commentPersonId
			$commentPerson.SetAttribute('id', $commentPersonId)
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
	foreach ($commentEntry in @($commentArchive.Entries | Where-Object { $_.FullName -like 'xl/threadedComments/*.xml' })) {
		$commentPath = $commentEntry.FullName
		$commentReader = [IO.StreamReader]::new($commentEntry.Open())
		[xml]$commentXml = $commentReader.ReadToEnd()
		$commentReader.Dispose()
		foreach ($commentNode in $commentXml.SelectNodes('//*[@personId or @mentionpersonId]')) {
			foreach ($commentAttribute in @('personId', 'mentionpersonId')) {
				$commentOldId = $commentNode.GetAttribute($commentAttribute)
				if ($commentPersonIds.ContainsKey($commentOldId)) { $commentNode.SetAttribute($commentAttribute, $commentPersonIds[$commentOldId]) }
			}
		}
		$commentEntry.Delete()
		$commentWriter = [IO.StreamWriter]::new($commentArchive.CreateEntry($commentPath).Open(), [Text.UTF8Encoding]::new($false))
		$commentWriter.Write($commentXml.OuterXml)
		$commentWriter.Dispose()
	}
} finally {
	$commentArchive.Dispose()
}
