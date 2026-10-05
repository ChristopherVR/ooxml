# The live demos

The same demo app is built once per framework adapter and published next to these docs. Each one is the real `<docx-editor>` running entirely in your browser: nothing is uploaded, and there is no server behind them.

| Demo                                         | Adapter               | Source                                                      |
| -------------------------------------------- | --------------------- | ----------------------------------------------------------- |
| [React](/demo/){target="_self"}              | `docx-react-viewer`   | `demos/demo-vanilla` built with `VITE_DEMO_FRAMEWORK=react` |
| [Vue](/demo-vue/){target="_self"}            | `docx-vue-viewer`     | same app, `vue`                                             |
| [Angular](/demo-angular/){target="_self"}    | `docx-angular-viewer` | same app, `angular`                                         |
| [Svelte](/demo-svelte/){target="_self"}      | `docx-svelte-viewer`  | same app, `svelte`                                          |
| [Solid](/demo-solid/){target="_self"}        | `docx-solid-viewer`   | same app, `solid`                                           |
| [Vanilla JS](/demo-vanilla/){target="_self"} | `docx-vanilla-viewer` | same app, `vanilla`                                         |

The home page embeds them with a framework switcher: [open the live demo section](/#live-demo).

## What you can do in a demo

- Choose **Open the sample document**, or drop in a DOCX or legacy DOC file of your own. Files stay in the tab.
- Add `?sample=1` to a demo URL to open the sample document straight away, and `?locale=de` (also `fr`, `es`, `zh-CN`) to start in another interface language.
- Use the demo's theme button, or the theme toggle on this site: all the demos follow the `vitepress-theme-appearance` preference live (see [theming](/theming)).

## Collaboration demo

The [two-peer coauthoring demo](/demo/collaboration.html){target="_self"} puts two editors on one document through a local, in-memory authority. It is the reference for the host-owned protocol described in [collaboration](/collaboration), and it is not a hosted service. Add `?guest=vue` (or another framework key) to run the second editor in a different adapter.

## Embedded here

<iframe
	src="/docx-viewer/demo/?sample=1"
	title="docx-viewer live demo"
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
