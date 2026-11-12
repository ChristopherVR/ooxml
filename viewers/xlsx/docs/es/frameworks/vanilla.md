# JavaScript puro

Sin framework: `mountEditor` crea el elemento en un contenedor y devuelve el [handle](/es/bindings#handle), además de `update(options)` y `destroy()`.

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

O use el elemento directamente:

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const element = document.createElement('xlsx-editor');
element.setAttribute('locale', 'de');
document.body.append(element);
element.newWorkbook();
```

Consulte la [API del elemento](/es/api) para ver cada atributo, método y evento.

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de JavaScript puro](/demo-vanilla/).
