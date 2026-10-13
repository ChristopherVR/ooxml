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
It deliberately leaves page-dependent unit tags, compound feet/inches, date/time pictures,
other precision codes and unverified units at their stored text. This does not
claim complete numeric formatting or locale parity.

Microsoft documents the internal-picture mechanism in
[FIELDPICTURE](https://learn.microsoft.com/en-us/office/client-developer/visio/fieldpicture-function)
and explicit numeric pictures in
[FORMAT](https://learn.microsoft.com/en-us/office/client-developer/visio/format-function).
The concrete code mapping and unit conversion here are verified by the native
fixture rather than inferred from those examples.

The boundary fixture adds 104 native cases for positive/negative half-decimal
rounding, zero, nonzero values smaller than the fixed-decimal precision and
15-digit numbers. Regenerate it with the same script and -BoundaryCases.
Native uses scientific notation for nonzero values that would round to zero;
unit-bearing picture codes use a lower-case exponent. Integer overflow beyond
15 digits remains unverified and keeps stored text.

The general fixture adds 46 native cases for FIELDPICTURE 0 and 1. These use up
to four optional decimal places, scientific notation for very small values,
and the same explicit units. General magnitudes above 1e15 remain unverified
and use stored text. Regenerate with `-GeneralCases`. Both extended corpora
include scientific mantissa ties and exponent carries; native fixed pictures
expand their mantissa precision after a carry.

All 214 fields match core parsing, unrelated page rename/save/reopen and browser
export/reopen. Core-edited drawings also match in native Visio before and after
native resave/reopen. Run `scripts/verify-visio-numeric-fields.ps1` against a
directory containing the three core-edited drawings and their JSON evidence.
These checks compare field text, not general document geometry or locale parity.

The fractional fixture adds 96 native cases for FIELDPICTURE 15 through 18:
nearest one/two-digit denominators, with or without explicit units. Positive
and negative values, zero, small fractions, whole-number carries, near-half
boundaries, thirds and all seven admitted units are included. The bounded core
formatter searches denominators up to 9 or 99, without allocating candidate
arrays. Values above 1e15 and nonfinite values remain unsupported.

Regenerate with `-FractionCases`. All four corpora (310 fields) pass core
edit/save/reopen, browser export/reopen and native resave/reopen. Compound
feet/inches pictures 10/13/14 and explicit FORMAT fraction strings are separate
unsupported profiles; the new formatter does not claim those behaviors.
