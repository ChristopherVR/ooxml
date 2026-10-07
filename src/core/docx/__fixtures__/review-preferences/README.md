# Native Word review preferences

`scripts/record-word-review-preferences.ps1` generated these owned synthetic
documents using hidden desktop Word build 16.0.20430. `preferences.docx` enables
Track Changes and disables Track Formatting and Track Moves. `edited.docx`
applies bold to "Preference" and inserts "!". Word reports just the insertion;
Reject All removes it and retains bold. The JSON records these native results.

The core export reference reopens the package with both preferences enabled and
the standalone package with both disabled. Native Word reads the expected flags,
records bold only for the enabled case, and rejects it back to the baseline.
This follows the documented
[formatting flag](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.donottrackformatting?view=openxml-3.0.1)
and [move flag](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.donottrackmoves?view=openxml-3.0.1).

Tests preserve and edit the OOXML settings, import their values into the editor,
share preferences through Yjs, and verify recording and export. This is desktop
interoperability evidence, not full current Microsoft 365 certification.
