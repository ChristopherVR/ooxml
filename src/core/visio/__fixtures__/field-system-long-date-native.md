# Inherited English system long dates

`scripts/visio-system-date-field-acceptance.ps1 -LongDate` records forty-four
owned Visio 16.0 instances, with master fields for `FIELDPICTURE(21)` and
`FIELDPICTURE(201)` and instance DATE overrides. The master placeholder is
`Thursday, January 1, 2026`. Valid cases cover all twelve months and seven
weekdays, leap/non-leap centuries, the 1899-12-30 epoch and time fractions.

Only the exact stored English `dddd, MMMM d, yyyy` profile is admitted. The
existing token formatter validates the complete cached calendar date and its
weekday, then formats the instance's DATE value. No language or ordering is
inferred from machine culture. German/French, abbreviated/reordered, invalid
leap/day/weekday, leading-zero and pre-epoch cache controls remain unchanged.
The renderer does not alter Value, unit, Formula, Format or inherited text XML.

Core tests compare forty-two supported native displays and exact page/master
XML through rename/export/reopen; two pre-epoch source cases retain the master
placeholder. Browser tests cover every case, rename undo/redo and export/reopen.
The owned native reopen checker validates all forty-four value/picture formulas
and forty-two supported texts through core-export open, save and reopen.

Native negative-date behavior is explicitly unsupported: `DATE(1899,12,29)`
produced a year-30177 long date before save. Reopening both the untouched native
source and the core export changed picture 21's weekday and made picture 201
blank. The recorded `unsupportedReopen` entries preserve that evidence; these
two cases are not used to infer a rendering rule or claimed native text parity.

This does not establish other languages, English date ordering/punctuation,
historical calendars or other internal date pictures. The companion system
date corpus covers the separately guarded ISO date and English time profiles.
