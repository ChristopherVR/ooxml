# Native numeric inherited fields

`numeric-field-native.vsdx` and its JSON evidence are synthetic fixtures generated
with Microsoft Visio 16 through `scripts/visio-numeric-field-acceptance.ps1`.
No Microsoft stencil or document content is copied. Each master has a numeric
field with value 1; its instance overrides the User cell and Field.Value cache.
The master's stored text therefore cannot stand in for the instance's text.

All 64 cases cover FIELDPICTURE codes 2 through 9: zero to three fixed decimals,
with or without units. Values cover scalar, negative scalar, inches, millimeters,
centimeters, feet, degrees and radians. V stores internal inches or radians;
U supplies the display unit. Core parsing and unrelated page rename/save/reopen
match the captured native Characters.Text in every case. A browser regression
checks every rendered field before and after browser export/reopen.

The implementation reuses the existing core numeric format-picture formatter.
It deliberately leaves page-dependent unit tags, fractions, date/time pictures,
other precision codes and unverified units at their stored text. This does not
claim complete numeric formatting or locale parity.

Microsoft documents the internal-picture mechanism in
[FIELDPICTURE](https://learn.microsoft.com/en-us/office/client-developer/visio/fieldpicture-function)
and explicit numeric pictures in
[FORMAT](https://learn.microsoft.com/en-us/office/client-developer/visio/format-function).
The concrete code mapping and unit conversion here are verified by the native
fixture rather than inferred from those examples.