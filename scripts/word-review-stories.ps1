# Read only the main, note, header and footer stories in the caller's owned document.
function Get-WordReviewStories($Document) {
    $stories = @()
    foreach ($type in @(1, 2, 3, 6, 7, 8, 9, 10, 11)) {
        $range = $null
        try { $range = $Document.StoryRanges.Item($type) } catch { continue }
        $index = 0
        while ($range) {
            $stories += [ordered]@{ type = $type; index = $index++; text = [string]$range.Text; revisions = [int]$range.Revisions.Count }
            $next = $range.NextStoryRange
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
            $range = $next
        }
    }
    return $stories
}
