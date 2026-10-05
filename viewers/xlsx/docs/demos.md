# The live demos

The same demo app is built once per framework adapter and published next to these docs. Each one is the real `<xlsx-editor>` running entirely in your browser: workbooks are parsed, edited and saved in the tab, nothing is uploaded, and there is no server behind them.

| Demo                                         | Adapter                            |
| -------------------------------------------- | ---------------------------------- |
| [React](/demo/){target="_self"}              | `@christophervr/xlsx-react-viewer` |
| [Vue](/demo-vue/){target="_self"}            | `xlsx-vue-viewer`                  |
| [Angular](/demo-angular/){target="_self"}    | `xlsx-angular-viewer`              |
| [Svelte](/demo-svelte/){target="_self"}      | `xlsx-svelte-viewer`               |
| [Solid](/demo-solid/){target="_self"}        | `xlsx-solid-viewer`                |
| [Vanilla JS](/demo-vanilla/){target="_self"} | `xlsx-vanilla-viewer`              |

The home page embeds them with a framework switcher: [open the live demo section](/#live-demo).

## What you can do in a demo

- Drop in an `.xlsx`, `.xlsm`, legacy `.xls` or `.csv` file, start a new workbook, or open the sample workbook.
- Add `?sample=1` to open the sample workbook straight away, and `?locale=de` (any supported tag) to start in another interface language.
- Use the demo's theme button or this site's theme toggle: the demos follow the `vitepress-theme-appearance` preference live (see [theming](/theming)).

## Collaboration

There is no collaboration demo for Excel because co-authoring is not implemented for spreadsheets. See [collaboration](/collaboration) for what works today (single editor, `readOnly` viewers).

## Embedded here

<iframe
	src="/xlsx-viewer/demo/?sample=1"
	title="xlsx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Run them locally

```bash
bun install
bun run demo
```

`bun run docs:build` (from `docs/`) builds the documentation and all six demos exactly as the Pages workflow does.
