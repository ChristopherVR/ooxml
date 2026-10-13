# Inherited system date profiles

`scripts/visio-system-date-field-acceptance.ps1` records thirty owned Visio 16.0
cases with inherited master fields and instance `User.DateValue` overrides.
The machine's culture is en-US, but its configured system short-date output is
ISO `yyyy-MM-dd`. Culture alone is insufficient to choose a system picture.

The renderer previously retained each master's January placeholder despite
the instance's cached DATE value. This slice admits picture 200 only with a
valid stored ISO date; pictures 211/212 require a valid stored ISO date with
the exact measured English twelve-hour AM/PM minute/second suffix. It reuses
the existing DATE cache parser and date-token formatter. Other locale caches,
invalid dates, mismatched suffix forms and pre-epoch source values remain
unchanged. Explicit ISO picture 204 retains its existing unconditional profile.

The corpus covers midnight/noon, single-digit hours, leap dates, the 1899-12-30
epoch, pre-epoch refusal and fractional seconds near midnight. Core tests
compare native text and preserve exact page/master XML through an unrelated
page rename and reopen. Browser tests check all thirty fields, rename history,
export and reopen. An owned native app also opened, saved and reopened that
core export: all thirty text, value-formula and picture-formula checks matched.
The reproduction checker is `scripts/visio-system-date-reopen.ps1`.

This corpus does not establish other system date/time locales or other date
picture codes. The companion `field-system-long-date-native.md` records the
separately validated English long-date pictures 21/201. It does not change field insertion
choices or claim general Visio formatting parity.
