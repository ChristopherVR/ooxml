# Style-inherited page break reference

Recorded 2026-10-07 in a separate hidden Word instance, build 16.0.20430
(Click-to-Run 16.0.20430.20140). Installed licenses are perpetual editions;
this reference does not certify current M365 subscription behavior.

`scripts/record-word-page-break-style.ps1` creates three synthetic paragraphs.
The second and third use a style with `w:pageBreakBefore`; the third explicitly
disables it with `w:val="0"`. Word places them on pages 1, 2 and 2. The native
DOCX and `reference.json` retain that evidence. Core tests cover source parsing,
style resolution, explicit off, editor attributes, text edits and layout page
assignments. Source styles remain unchanged during package export. Existing
documents and Office profile settings are untouched.

`core-export-reference.json` records Word reopening a package export after adding
`!` to the explicitly disabled paragraph. It retains page assignments 1, 2, 2
and the effective off override.

`tracked-core-export-reference.json` records Word opening a tracked direct off
override on the inherited-break paragraph: one paragraph revision (10), all
paragraphs on page 1, and native rejection restoring page assignments 1, 2, 2.
The fixture includes Word's calculated `lastRenderedPageBreak` cache marker.
Empty cache markers do not block edits and are invalidated when their run is
rewritten. Untouched package saves retain their original bytes. Unknown cache
attributes and other unsupported inline content remain guarded.

Microsoft documents the cache marker as the
[position of the last calculated page break](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.lastrenderedpagebreak?view=openxml-3.0.1).
