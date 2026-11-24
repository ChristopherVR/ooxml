# PDF import fixtures (#319)

| File | Made by |
| --- | --- |
| `source/rapport.docx` | `generate.mjs` (run from `src/core`): a two-page report with headings, justified paragraphs, bold, italic, underlined, coloured and superscript text, a centred PNG picture, a header and a footer with PAGE and NUMPAGES fields. Generic content. |
| `rapport.pdf` | `export-word.ps1`: Microsoft Word 16 exported `source/rapport.docx` through its documented COM object model (`Documents.Open`, `ExportAsFixedFormat`), with document properties and structure tags. |
| `source/tableaux.docx`, `tableaux.pdf` | `generate.mjs` and `export-word.ps1` (Word 16): a ruled four-column table with a shaded header row (`w:tblHeader`, so Word repeats it on page 2), a cell spanning two columns and a vertically merged cell, crossing the page break; then a borderless three-column table. |
| `source/listes.docx`, `listes.pdf` | The same: a bulleted list on two levels (Symbol bullet, Courier New "o") with a wrapped item, and a numbered list ("1.", "a)") on two levels, both real Word numbering. |
| `source/colonnes.docx`, `colonnes.pdf` | The same: a one-column introduction, a continuous two-column section (balanced), then one column again. |
| `lettre.pdf` | The Word PDF export of this repository (`exportDocxPdf`) from `src/core/docx/layout/fixtures/table-reports/t10-tableau-entete-pied.docx`: standard fonts, no embedding, a few kilobytes. |
| `lettre-rc4-owner-only.pdf`, `lettre-aes128.pdf`, `lettre-aes256.pdf` | `encrypt.py` (pypdf, an independent implementation of the standard security handler): RC4-128 with an empty user password, AES-128 and AES-256 with the user password `secret` and the owner password `proprietaire`; printing and copying allowed, editing not. |

`export-word.ps1 -Names tableaux,listes` exports only the named documents (the committed PDFs keep their bytes otherwise).

`src/core/pdf/reader.test.ts`, `../import.test.ts` and `../structure.test.ts` read them; `e2e/docx/pdf-open.spec.ts` opens `rapport.pdf` in the editor.
