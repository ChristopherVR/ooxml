# Browser test fixtures

- `excel-encrypted.xlsx`: Excel 16 encrypted workbook, password `open sesame`. Copied from
  `ChristopherVR/ooxml`, `src/xlsx/__fixtures__/encrypted/excel-encrypted.xlsx`.
- `excel-smartart.xlsx`: Excel 16 Basic Block List with Plan, Build and Ship. Copied from
  `ChristopherVR/ooxml`, `src/xlsx/__fixtures__/excel-smartart.xlsx`.
  Both fixtures were copied from commit `4b4c7a006b22fe915ab9e682c37b20621c4869db`.

- `legacy-97.xls`: the demo sample workbook (`demos/demo-vanilla/public/sample.xlsx`) saved by
  Excel 16 as Excel 97-2003 (BIFF8, `FileFormat 56`). Regenerate with
  `python scripts/sample-workbook.py` then `pwsh scripts/sample-workbook-excel.ps1` (Windows with
  Excel installed). Specs load it through the landing page's file input; `pack:smoke` and the
  vanilla entry test open it through every package.
