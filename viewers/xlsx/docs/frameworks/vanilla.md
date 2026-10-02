# Vanilla JavaScript

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `xlsx-vanilla-viewer` package; it has no framework peer.
:::

No framework: `mountEditor` creates the element in a container and returns the [handle](/bindings#handle) plus `update(options)` and `destroy()`.

```ts
import { mountEditor } from 'xlsx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor')!, {
	fileName: 'Budget.xlsx',
	onWorkbookChange: (workbook) => console.log(workbook.sheets.length),
	onWorkbookError: (error) => console.error(error),
});
await editor.load(bytes, 'Budget.xlsx');
editor.update({ readOnly: true });
const blob = await editor.save();
editor.destroy();
```

Or use the element directly:

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const element = document.createElement('xlsx-editor');
element.setAttribute('locale', 'de');
document.body.append(element);
element.newWorkbook();
```

See the [element API](/api) for every attribute, method and event.

See the [complete binding contract](/bindings) or [try the Vanilla JavaScript demo](/demo-vanilla/).
