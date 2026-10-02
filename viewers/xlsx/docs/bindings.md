# Framework bindings

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). Each framework has its own self-contained package: `@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer`, `xlsx-angular-viewer`, `xlsx-svelte-viewer`, `xlsx-solid-viewer` and `xlsx-vanilla-viewer`. Use one editor package per application.
:::

All bindings mount `<xlsx-editor>` through the same `mountEditor` function in `packages/bindings/src/index.ts`, so every framework gets identical semantics.

## Props

| Prop             | Type                          | Default      | Notes                                                                                                  |
| ---------------- | ----------------------------- | ------------ | ------------------------------------------------------------------------------------------------------ |
| `workbook`       | `Workbook`                    | none         | Shown when a **new object** is passed. The workbook the editor emits is never assigned back (no loop). |
| `bytes`          | `Uint8Array \| ArrayBuffer`   | none         | File bytes; loaded again whenever a new array is passed.                                               |
| `src`            | `string`                      | none         | URL fetched and loaded when the string changes; the file name defaults to the URL's last segment.      |
| `fileName`       | `string`                      | `Book1.xlsx` | Forwarded only when the parent changes it, so File > Open in the editor can rename the workbook.       |
| `readOnly`       | `boolean`                     | `false`      |                                                                                                        |
| `locale`         | `string`                      | `en`         | `en`, `fr`, `de`, `es`, `zh-CN` or a tag that maps to one (`de-DE`).                                   |
| `theme`          | `'light' \| 'dark' \| 'auto'` | `auto`       | `auto` follows the operating system.                                                                   |
| `authorName`     | `string`                      | `Author`     | Recorded on new comments.                                                                              |
| `showToolbar`    | `boolean`                     | `true`       | The ribbon.                                                                                            |
| `showFormulaBar` | `boolean`                     | `true`       | The name box and formula bar.                                                                          |
| `hiddenActions`  | `string[]`                    | `[]`         | Ribbon controls to hide, by stable id.                                                                 |
| `themeColors`    | `Partial<Record<...>>`        | `{}`         | Token overrides, see [theming](/theming).                                                              |

## Callbacks and events

| Callback (React, Solid, vanilla) | Vue event          | Angular output    | Svelte prop         | Payload                     |
| -------------------------------- | ------------------ | ----------------- | ------------------- | --------------------------- |
| `onWorkbookChange`               | `workbook-change`  | `workbookChange`  | `onworkbookchange`  | `Workbook`                  |
| `onWorkbookError`                | `workbook-error`   | `workbookError`   | `onworkbookerror`   | `Error`                     |
| `onSelectionChange`              | `selection-change` | `selectionChange` | `onselectionchange` | `{ sheet, ref, active }`    |
| `onDirtyChange`                  | `dirty-change`     | `dirtyChange`     | `ondirtychange`     | `boolean`                   |
| `onReady`                        | `ready`            | `ready`           | `onready`           | the `<xlsx-editor>` element |

Load failures (corrupt, encrypted or unsupported files) reach `onWorkbookError`; a failed `src` fetch does too. The element's other events (`workbook-warning`, `file-command`, `sheet-change`, `ribbon-action`, `readonly-change`) are listened to on the element itself, see the [element API](/api).

## Handle

Every adapter exposes the same handle: React through `ref`, Vue through the template ref, Angular through the component instance, Solid through `editorRef`, Svelte through the component's exports and vanilla as the return value of `mountEditor`.

| Member                          | Effect                                                              |
| ------------------------------- | ------------------------------------------------------------------- |
| `element`                       | The mounted `<xlsx-editor>`.                                        |
| `load(bytes, fileName?)`        | Opens `.xlsx`, `.xlsm`, `.xls` or `.csv` bytes; rejects on failure. |
| `newWorkbook()`                 | Replaces the workbook with a blank one.                             |
| `save()`                        | Resolves to an `.xlsx` Blob. Does not clear `dirty`.                |
| `saveBytes(format?)`            | `'xlsx'` (default) or `'csv'` (active sheet) bytes.                 |
| `download(fileName?)`           | Saves, downloads in the browser and marks the workbook clean.       |
| `markClean()`                   | Clears `dirty` after you persisted a saved Blob yourself.           |
| `select(ref)`, `getSelection()` | Selection as an A1 reference (`'B2:D8'`).                           |
| `setActiveSheet(index)`         | Switches sheet.                                                     |
| `dirty`                         | Unsaved edits (Svelte: `isDirty()`).                                |

Calling a handle method before the editor is mounted throws `Editor is not mounted`.

## Server rendering

Every package imports without a DOM. Registration (`defineXlsxEditor`) and mounting are client operations.
