# Native continuous-section page fields

Generated with `scripts/record-word-continuous-page-fields.ps1` using desktop
Word `16.0.20430.20140` (COM reports `16.0.20430`), then measured from Word's
PDF exports with `scripts/record-word-page-field-text.py` and pypdf 6.10.0.
Installed licenses: Professional 2021/ProPlus 2024. This records desktop Word
behavior without claiming M365 subscription certification.

All documents use Arial 12 pt, exact 12 pt body lines, 0 before/after spacing,
Letter pages and 72 pt margins. They have separate first/default/even headers
and footers for sections A and B. Footer fields are PAGE and SECTIONPAGES.
Section B begins continuously on an odd or even physical page, then overflows.

The twelve cases cover default continuation, different-first-page headers,
odd/even headers, decimal/Roman numbering and restarts at 1, 10 and 11. The
visible header/footer on a shared page belongs to its first section. Both
sections count the shared page for SECTIONPAGES. With odd/even headers, Word
aligns a restart with the shared page's parity; an incompatible start is
decremented for that shared page. Later pages continue the resulting number.

To regenerate, run the PowerShell script, then:

```text
python scripts/record-word-page-field-text.py <output-directory>
```

Retain the DOCX sources and evidence.json. PDFs remain diagnostic outputs,
rather than being committed or shipped in the package. Tests compare actual
core page facts and rendered headers/footers with the native PDF text.
