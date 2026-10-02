# Browser test fixtures

- `legacy-97.xls`: the demo sample workbook (`demos/demo-vanilla/public/sample.xlsx`) saved by
  Excel 16 as Excel 97-2003 (BIFF8, `FileFormat 56`). Regenerate with
  `python scripts/sample-workbook.py` then `pwsh scripts/sample-workbook-excel.ps1` (Windows with
  Excel installed). Specs load it through the landing page's file input; `pack:smoke` and the
  vanilla entry test open it through every package.
