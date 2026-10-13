# Independent native precedence controls

Seven owned synthetic Visio 16.0.20430.20146 drawings independently test page, direct DocumentSheet and Style API mutations. scripts/visio-print-inheritance-acceptance.ps1 reproduces the controls without opening user drawings, changing printer configuration or printing.

Direct DocumentSheet paper, fit and zoom writes persist in source but do not affect the resolved page or native count (nine). A page paper override selects A4 and changes the native count to six. Attempts to write root-style print cells through this COM Style API were refused. This is an API refusal, not proof that the entire style is protected.

Raw VSDX bytes, resolved page cells, requests, native counts and SHA-256 values are retained separately from the original 79-case corpus. No native reopen or printer-output acceptance is claimed. Core no-op and unrelated rename regressions preserve all other package parts, including document print defaults.
