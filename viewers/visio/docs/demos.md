# The live demos

The same playground is built once per framework adapter and published next to these docs. Each one is the real `<visio-viewer>` running entirely in your browser: drawings are parsed and rendered locally, nothing is uploaded, and there is no server behind them. This is a public beta; a demo opening a file is not evidence of Visio fidelity.

| Demo                                      | Adapter                | Route            |
| ----------------------------------------- | ---------------------- | ---------------- |
| [Vanilla JS](/demo/){target="_self"}      | `visio-vanilla-viewer` | `/demo/`         |
| [React](/demo-react/){target="_self"}     | `visio-react-viewer`   | `/demo-react/`   |
| [Vue](/demo-vue/){target="_self"}         | `visio-vue-viewer`     | `/demo-vue/`     |
| [Angular](/demo-angular/){target="_self"} | `visio-angular-viewer` | `/demo-angular/` |
| [Svelte](/demo-svelte/){target="_self"}   | `visio-svelte-viewer`  | `/demo-svelte/`  |
| [Solid](/demo-solid/){target="_self"}     | `visio-solid-viewer`   | `/demo-solid/`   |

The vanilla playground is also served at [/demo-vanilla/](/demo-vanilla/){target="_self"}, the name the other viewers' sites use. The home page embeds the demos with a framework switcher: [open the live demo section](/#live-demo).

## What you can do in a demo

- Open the original sample drawing, or a local `.vsdx` file you choose.
- Add `?sample=1` to a demo URL to open the sample straight away, and `?embed=1` for the compact embedded layout the home page uses.
- Use the demo's theme button or this site's theme toggle: all the demos follow the `vitepress-theme-appearance` preference live (see [theming](/theming)).

## Collaboration demo

Add `?share=<session>` to any demo URL to join a File > Share session on load. Two demos built from different frameworks that use the same session name share one drawing, for example `/demo-vue/?sample=1&share=my-room` and `/demo-angular/?share=my-room` in two tabs of one browser. The home page shows two panes with a framework picker for each. See [collaboration](/collaboration) for what is and is not supported.

## Embedded here

<iframe
	src="/visio-viewer/demo/?embed=1&sample=1"
	title="visio-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Run them locally

```bash
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run dev
```

`npm run docs:build` builds the documentation and all demos exactly as the Pages workflow does.
