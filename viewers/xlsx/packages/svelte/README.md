# xlsx-svelte-viewer

Available on npm. Install the package using the command below.

The Excel spreadsheet editor for Svelte 5. One self-contained package: the `<xlsx-editor>` web component, the legacy Excel 97-2003 `.xls` reader and the Svelte 5 adapter are bundled in, and the workbook model (`@christophervr/xlsx-core`: `createWorkbook`, `loadWorkbook`, `saveXlsx`, ...) installs with it and is re-exported, so one install and one import path are all an application needs.

```sh
npm install xlsx-svelte-viewer svelte
```

```svelte
<script>
  import XlsxEditor from 'xlsx-svelte-viewer';
  import { createWorkbook } from 'xlsx-svelte-viewer/runtime';
  let workbook = $state.raw(createWorkbook());
</script>
<XlsxEditor {workbook} onworkbookchange={(next) => (workbook = next)} />
```

The package root is the component (a default export). The plain-JavaScript helpers (`mountEditor`, `loadWorkbook`, `createWorkbook`, shared types) are under `xlsx-svelte-viewer/runtime`.

Notes:

- `.xlsx`, `.xlsm`, legacy `.xls` and `.csv` files are opened; the format is sniffed from the bytes. A legacy `.xls` file is saved back as `.xlsx`.
- svelte is a peer dependency. `ooxml-core` and `ooxml-ui` are regular dependencies; never import `ooxml-ui` yourself.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<xlsx-editor>` element.
- An early editor: not Microsoft Excel parity, and saving is not lossless for unsupported features (pivot tables, slicers, sparklines and macros are kept but not editable). See the [features and limitations](https://christophervr.github.io/xlsx-viewer/features).

Documentation: [element API](https://christophervr.github.io/xlsx-viewer/api), [Svelte 5 guide](https://christophervr.github.io/xlsx-viewer/frameworks/svelte). Licensed under Apache-2.0.
