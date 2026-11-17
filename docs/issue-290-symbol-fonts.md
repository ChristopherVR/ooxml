# Issue #290: missing symbol bullets

Investigation on 17 November 2026, following
[Tom's comments](https://github.com/ChristopherVR/ooxml/issues/290#issuecomment-6945570223).

## Why the square still showed as a section sign

The Unicode mapping commits `a5dc3a0dc`, `b9a333032` and `82589a0b0` were on the
local `fix/safari-symbol-bullets` branch, absent from `origin/main` at
`164a9ac10` and from the `pptx-react-viewer@4.80.0` release. The earlier issue
reply described code that had not shipped. This change integrates those commits.

Wingdings maps the stored character `§` to a small square. Georgia interprets
the same character as a section sign. Missing Unicode glyph coverage cannot
cause that reversal: the translation itself must happen first.

Georgia coverage varies by font version. Tom reports that his version has
U+25A0 but lacks U+2751, U+2756, U+2B9A and U+2713. The installed Georgia tested
here lacks all five in its public Unicode cmap. Neither observation establishes
coverage of Apple's Georgia. A browser may also show a character through a
system fallback even when Georgia itself lacks it.

## PowerPoint embedding experiment

Using the documented PowerPoint COM object model, create a blank presentation,
add a text box, and save as PPTX with `SaveAs(path, 24, -1)` (embedding enabled):

| Content                                   | `Presentation.Fonts`             | Saved embedded font list |
| ----------------------------------------- | -------------------------------- | ------------------------ |
| Georgia text with a Wingdings `§` bullet  | Georgia and Aptos, not Wingdings | Georgia                  |
| Ordinary text explicitly set to Wingdings | Wingdings and Aptos              | Empty                    |

PowerPoint's `Font.Embeddable` returned true for Wingdings. Its installed font's
public OS/2 `fsType` was 8 (editable embedding). This reproduces the omission
without an iPad and does not support blaming a restrictive licence flag. It
also shows that bullet-only font enumeration is not the entire explanation.
The exact internal reason PowerPoint skips this font remains unconfirmed.
No Microsoft implementation internals were inspected.

Microsoft documents [font embedding permissions](https://learn.microsoft.com/en-us/typography/fonts/font-faq)
and [PowerPoint's Embeddable property](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Font.Embeddable),
but neither guarantees that every eligible font is actually written to a deck.
The experiment used synthetic files, not Tom's linked sample.

## Implemented fallback

1. Keep the original symbol font when the browser can paint its glyph.
2. Otherwise translate known Symbol/Wingdings characters to Unicode, keeping
   the paragraph's text font first in the CSS family list.
3. If that font lacks the glyph, use the bundled Noto Sans Symbols 2 subset
   before the generic system fallback. No download or persistence is involved.

The subset covers U+2000-2BFF and includes all of Tom's listed symbols. Its
licence ships in the UI package's NOTICE; the pinned source and reproducible
generation command are in PROVENANCE.md. The same inline font rule is collected
by the existing image/print export code. Source text and OOXML remain unchanged.

This supplies approximate Unicode equivalents, not exact Wingdings outlines.
Unmapped bullet characters fall back to a plain bullet; full Wingdings 2/3 and
Webdings mappings remain tracked in the parity checklist. Real iPadOS testing
is still needed in addition to automated browser coverage.

## Validation

- Core and UI builds and typechecks passed.
- Full UI suite: 17,282 passed, 11 skipped (1,687 files passed, 2 skipped).
- Core symbol mapping: 7 tests passed.
- Chromium and WebKit: the regression passed in all five framework demos,
  with symbol-font absence simulated and Google font endpoints blocked.
- Lint, formatting, and the framework-neutral browser-spec check passed.
