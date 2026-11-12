# JavaScript natif

Sans framework : `mountEditor` crée l’élément dans un conteneur et renvoie la [poignée](/fr/bindings#handle), ainsi que `update(options)` et `destroy()`.

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

Ou utilisez directement l’élément :

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const element = document.createElement('xlsx-editor');
element.setAttribute('locale', 'de');
document.body.append(element);
element.newWorkbook();
```

Consultez l’[API de l’élément](/fr/api) pour chaque attribut, méthode et événement.

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo JavaScript natif](/demo-vanilla/).
