# Native opaque run properties

Recorded 2026-10-07 in a separate hidden Word instance, build 16.0.20430
(Click-to-Run 16.0.20430.20140). Installed Office licenses are perpetual editions;
this reference does not certify current M365 subscription behavior.

`scripts/record-word-opaque-run-properties.ps1` creates `Outlined text` in Arial
12 pt, with outline and shadow on. `reference.json` records Word's values.
Existing documents and Office profile settings are untouched.

`core-export-reference.json` records Word reopening both package-preserving and
standalone exports with a tracked `!` inserted after four characters. All three
runs retain their complete source property basis. Word sees one text insertion
and outline/shadow still on; Reject All restores the original text, keeps both
properties on and leaves zero revisions. The adjacent core regression reproduces
these exports. Editor/Yjs and six browser bindings also cover tracked splits,
undo/redo and export with opaque property bases.

Outline and shadow rendering remain unsupported. This verifies preservation
during text edits, not visual parity or arbitrary editing of opaque properties.
