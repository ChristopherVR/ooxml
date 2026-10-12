# Native inherited string fields

The synthetic fixture was generated with owned Visio 16 automation using
`scripts/visio-inherited-field-acceptance.ps1`. It uses a new master containing
`User.ValueText = "Abc"` and a custom text field with `FIELDPICTURE(0)`. Four
instances override the user cell with strings `42`, `Ready`, `73.50` and empty.
The adjacent JSON records native `Characters.Text` for each instance.

The master retains `<fld>Abc</fld>` while the instance Field rows cache their
individual values as `U="STR"`. The parser must display those instance caches,
including the empty string, without reformatting numeric-looking strings.
This also occurs in native data graphic callouts. The fixture contains only
synthetic shapes; no built-in stencil content was copied.

The regression verifies parsing, unrelated page rename and reopen, and honest
fallback for an unimplemented internal picture. Numeric, date and other internal
field pictures are outside this acceptance case.

References: [FIELDPICTURE](https://learn.microsoft.com/en-us/office/client-developer/visio/fieldpicture-function)
and [AddCustomFieldU](https://learn.microsoft.com/en-us/office/vba/api/visio.characters.addcustomfieldu).
