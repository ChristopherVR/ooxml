# Vanilla JavaScript

Ohne Framework: `mountEditor` erzeugt das Element in einem Container und gibt das [Handle](/de/bindings#handle) sowie `update(options)` und `destroy()` zurück.

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

Oder verwenden Sie das Element direkt:

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const element = document.createElement('xlsx-editor');
element.setAttribute('locale', 'de');
document.body.append(element);
element.newWorkbook();
```

Alle Attribute, Methoden und Ereignisse finden Sie in der [Element-API](/de/api).

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die Vanilla-JavaScript-Demo aus](/demo-vanilla/).
