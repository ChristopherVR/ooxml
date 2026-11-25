# Word PDF export: PDF/A

`exportDocxPdf(model, { pdfA: true })` (`src/ui/src/docx/pdf/export-pdf.ts`, File > Export >
Create a PDF with "PDF/A compliant") writes an ISO 19005-2 (PDF/A-2) file. This page says what
that means, how it is checked, and what it does not claim.

## What the file contains

- XMP metadata (`src/ui/src/export/pdf-archive.ts`), uncompressed, with `pdfaid:part` 2 and the
  title, author, subject, keywords, producer and dates of the document information dictionary.
  Both copies are written from the same cleaned value (`infoText`): characters XML cannot hold
  (control characters, U+FFFE, U+FFFF) are dropped and unpaired surrogates become U+FFFD, so the
  dictionary and the XMP always agree (ISO 19005-2 6.6.2.3).
- An sRGB IEC61966-2.1 output intent (`pdf-srgb-icc.ts`): every colour the export writes is
  DeviceRGB or DeviceGray, and the output intent makes them device-independent. Transparency
  (watermarks, translucent fills, PNG alpha as soft masks) is allowed in PDF/A-2 and uses it.
- The trailer's file identifier, PDF 1.7 header, Flate compression only (no LZW), no encryption,
  no JavaScript, no external content.
- Fonts embedded as TrueType subsets (`pdf-embedded-font.ts`) with widths taken from the font
  program and a ToUnicode map for every character.
- Notes (comments) with their own appearance stream and the print flag; links with the print flag.

## Conformance level and fonts

PDF/A requires every font to be embedded, and the export can only embed font files the host
supplies (`fonts`, see [font-loading.md](./font-loading.md)). With `pdfA`, text in a family no
supplied font is named for (a heading in Segoe UI, a chart in Aptos Narrow) is written with a
supplied font that has its characters, reported in `report.substitutedFonts`, instead of an
unembedded standard PDF font. Without `pdfA` those families keep the standard fonts, as before.

The level the file declares follows what it contains (`pdfALevel`):

| Text in the file                      | Declares     | `report.pdfA`                                  |
| ------------------------------------- | ------------ | ---------------------------------------------- |
| All in embedded fonts                 | PDF/A-2**u** | `{ conformant: true, level: 'U', issues: [] }` |
| Some in standard fonts (not embedded) | PDF/A-2b     | `conformant: false`, the families in `issues`  |

Level U is level B plus a Unicode mapping for every character; every embedded font carries a
ToUnicode map, so a file that meets B meets U. Passing `pdfALevel: 'A'` checks the additional
PDF/A-2a requirements (structure tags, a document language and alternative text on every figure)
and declares level A only when all of them are met, otherwise falling back to level U or B and
listing the unmet requirements in `report.pdfA.issues`.

Passing `pdfUa: true` checks PDF/UA-1 (ISO 14289-1) conformance (structure tags, embedded fonts,
document title, language and alternative text on every figure). When all requirements are met, it
writes the PDF/UA-1 identification schema (`pdfuaid:part` = 1) in XMP and sets
`report.pdfUa.conformant` to `true`; otherwise it omits the claim and lists the issues in
`report.pdfUa.issues`.

## Validating with veraPDF

`scripts/pdfa-validate.mjs` exports a corpus with `pdfA: true` and validates it with
[veraPDF](https://verapdf.org), the open-source reference validator. The corpus
(`src/ui/src/docx/pdf/__fixtures__/pdfa-corpus.ts`) covers embedded fonts with accented and
non-WinAnsi characters, Unicode and XML-special metadata, a PNG with alpha and a JPEG, a table,
hyperlinks and links to bookmarks, the outline, the structure tree, headers, footers, a footnote,
tracked changes and comment notes (markup), a chart, shapes and a text box, a translucent text
watermark and a gradient fill. A control file in standard fonts must fail.

The script never downloads veraPDF. To install it (it needs Java 11 or later):

1. Download `https://software.verapdf.org/releases/verapdf-installer.zip` and unzip it.
2. Install the command line unattended, `java -jar verapdf-izpack-installer-<version>.jar
   auto-install.xml`, with an `auto-install.xml` that selects the "veraPDF CLI" pack (see
   [automated installation](https://docs.verapdf.org/install/#automated-installation)).
3. Run it:

```sh
VERAPDF=/path/to/verapdf node scripts/pdfa-validate.mjs
# options: --out <dir>  --flavours 2b,2u,2a,ua1  --fonts <a.ttf:b.ttf> | --fonts synthetic
```

`VERAPDF` names the `verapdf` (or `verapdf.bat`) command or the folder it is installed in. The
fonts are those given with `--fonts` or `PDFA_FONTS`, else a system sans-serif font when one is
found, else the corpus's synthetic TrueType font. The script exits with 1 when a corpus file fails
PDF/A-2b or the control file passes it; the other flavours are informational. Without veraPDF,
the corpus still runs in the unit tests (`pdfa-corpus.test.ts`) and checks the report and the
declared level.

### Results

veraPDF 1.30.4, corpus of seven documents plus the control file, with Arial and with the
synthetic font:

| Profile  | Before this check                                                                              | Now                                  |
| -------- | ---------------------------------------------------------------------------------------------- | ------------------------------------ |
| PDF/A-2b | 6 of 7 pass (chart text, and later comment balloons, in unembedded standard fonts, 6.2.11.4.1) | 7 of 7 pass                          |
| PDF/A-2u | 0 of 7 (files declared B, 6.6.4)                                                               | 7 of 7 pass (declared U)             |
| PDF/A-2a | 0 of 7 (declared B)                                                                            | only 6.6.4 fails (A is not declared) |
| PDF/UA-1 | informational                                                                                  | informational, see below             |

The control file fails 2b (6.2.11.4.1, fonts not embedded) and reports `conformant: false`.

PDF/UA-1 is not claimed. The remaining findings are the missing PDF/UA identification schema in
the XMP (5), which the export does not write because it does not claim PDF/UA, and figures with
no alternative text in the source document (7.3: a drawing shape without a description). Page
borders, line numbers, watermarks and floats behind the header and footer are pagination
artifacts, so every content item is tagged or an artifact (7.1).
