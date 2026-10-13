# Native explicit ISO date picture

The owned Visio 16 corpus is regenerated with
`scripts/visio-iso-date-field-acceptance.ps1 -MsoIsoPicture`.
`FIELDPICTURE(204)` is the explicit ISO date picture (`visFmtMsoDateISO`),
separate from the system short-date picture tested in the companion corpus.
See Microsoft's [format enumeration](https://learn.microsoft.com/en-us/office/vba/api/visio.visfieldformats).

Eleven supported Gregorian dates match native ISO displays in core and browser
save/reopen, including leap days and midnight fractional seconds. Locale-shaped,
invalid or empty stored placeholders do not constrain this explicit picture.
Shape XML remains unchanged through an unrelated core page edit.

The negative pre-epoch diagnostic `DATE(1899,12,29)` produces an empty native
ISO display and a five-digit-year cache. It remains unsupported in core, which
keeps the inherited master placeholder and preserves its raw XML. Native
edit/save/reopen retains all twelve source displays. Other date pictures and
formula evaluation remain outside the admitted profile.