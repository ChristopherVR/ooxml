# Office UI package

`ooxml-ui` (`src/ui`) is the second published package of this repository. It holds the web
components every Office product needs, and the product editors' DOM code under `ooxml-ui/<product>`.
The plan this page used to hold (a first publish, wave-by-wave tag migration and viewer adoption
steps written when the viewers lived in other repositories) is complete and was removed; this page
keeps the decisions that still apply. For how the elements are built and the subclass contract, read
`src/ui/README.md`.

## Decisions

- **Layout.** The repository root is a private workspace root. `src/core` is the published
  `ooxml-core` and `src/ui` the published `ooxml-ui`; the UI declares `"ooxml-core": "*"` and
  resolves the core through the workspace link. The release flow rewrites workspace ranges to caret
  ranges when it publishes (`docs/releasing.md`).
- **Direction.** UI may import core types and pure helpers; core never imports the UI (nothing in
  `src/core/` references `src/ui`).
- **Style.** Every shared element is a Lit class on `OfficeElement` (`src/ui/src/base.ts`): a
  `<name>.ts` with a `render()` template next to a real `<name>.css` imported `?raw`, using bare
  `var(--office-x)` tokens, grouped in a folder per kind. The product editors (PowerPoint, Word,
  Excel, Visio, Teams) are imperative custom elements that keep their own `?raw` CSS. Importing a
  module never touches `HTMLElement` at load time where SSR matters. Typed events include
  `office-command`, `office-dialog-close`, `office-presence-select`, `office-status-activate` and
  `office-smartart-render`, plus native `input` and `change` for value controls. Chart and SmartArt
  editing elements report edits as events and never write a model: `office-chart-change`,
  `office-chart-data-change`, `office-chart-format-edit`, `office-smartart-change`,
  `office-smartart-activate` and the `office-smartart-*` text pane requests.
- **Framework bindings.** Every product's framework adapters are generated from one contract per
  product, through the kit in `ooxml-ui/bindings` (`src/ui/src/bindings/`, see its README): React,
  Vue and Solid get generic factories, and Angular and Svelte keep a small adapter per product over
  the same helpers, since they need their own compilers. Each adapter is checked against its
  contract by a parity test, so the adapters cannot drift apart.
- **Product rule.** Viewers list `ooxml-ui` as a normal `dependency` of every published editor
  package, so users never install it by hand.
- **Honesty.** `office-ui-smartart` shows the cached drawing and reports approximated geometry,
  fills and flattened 3D in its render event. It makes no layout-parity claim. Editing exists where
  a product's model supports it (the SmartArt text pane and picker, the chart data sheet and format
  pane); what each product can edit is stated in its own documentation.
- **pptx tags.** The public `pptx-ui-*` tags remain as thin subclasses of the `office-ui-*`
  elements (`src/ui/src/pptx/web-components/office-aliases.ts`), which translate their old events
  and attributes, so existing hosts keep working. The chart and SmartArt insert dialogs are
  PowerPoint's own elements built from the shared pickers.

## Token map

The PowerPoint editor maps its `--pptx-*` variables to the `--office-*` contract on its host
element, so its theming API did not change. The other products follow the same contract with their
own prefixed overrides as fallbacks.

| pptx                            | ooxml-ui                     |
| ------------------------------- | ---------------------------- |
| `--pptx-foreground`             | `--office-foreground`        |
| `--pptx-muted-foreground`       | `--office-muted-foreground`  |
| `--pptx-background`             | `--office-background`        |
| `--pptx-border`, `--pptx-input` | `--office-border`            |
| `--pptx-primary`                | `--office-accent`            |
| `--pptx-primary-foreground`     | `--office-accent-foreground` |
| `--pptx-ring`                   | `--office-ring`              |
