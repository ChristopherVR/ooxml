# xlsx-vue-viewer

Available on npm. Install the package using the command below.

The Excel spreadsheet editor for Vue 3. One self-contained package: the `<xlsx-editor>` web component, the legacy Excel 97-2003 `.xls` reader and the Vue 3 adapter are bundled in, and the workbook model (`@christophervr/xlsx-core`: `createWorkbook`, `loadWorkbook`, `saveXlsx`, ...) installs with it and is re-exported, so one install and one import path are all an application needs.

```sh
npm install xlsx-vue-viewer vue
```

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-vue-viewer';

const workbook = shallowRef(createWorkbook());
</script>
<template>
	<SpreadsheetEditor :workbook="workbook" @workbook-change="workbook = $event" />
</template>
```

The shared option types and the `loadWorkbook` / `detectWorkbookFormat` helpers are exported from the package root.

Notes:

- `.xlsx`, `.xlsm`, legacy `.xls` and `.csv` files are opened; the format is sniffed from the bytes. A legacy `.xls` file is saved back as `.xlsx`.
- vue is a peer dependency. `ooxml-core` and `ooxml-ui` are regular dependencies; never import `ooxml-ui` yourself.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<xlsx-editor>` element.
- An early editor: not Microsoft Excel parity, and saving is not lossless for unsupported features (pivot tables, slicers, sparklines and macros are kept but not editable). See the [features and limitations](https://christophervr.github.io/xlsx-viewer/features).

Documentation: [element API](https://christophervr.github.io/xlsx-viewer/api), [Vue 3 guide](https://christophervr.github.io/xlsx-viewer/frameworks/vue). Licensed under Apache-2.0.
