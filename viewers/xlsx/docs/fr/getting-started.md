# Prise en main

## Lancer la démo depuis les sources

Vous avez besoin de [Bun](https://bun.sh/) 1.3 et de Node.js 24.

```bash
git clone https://github.com/ChristopherVR/ooxml.git
cd ooxml/viewers/xlsx
bun install
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

La page d’accueil ouvre un fichier `.xlsx`, `.xlsm`, `.xls` ancien format ou `.csv` (déposez-le ou parcourez vos fichiers), démarre un nouveau classeur ou ouvre le classeur d’exemple. Les fichiers sont traités uniquement dans le navigateur.

## Choisir un paquet

Installez un paquet d’éditeur autonome pour votre framework. Chacun embarque le composant web `<xlsx-editor>` et son adaptateur, dépend du modèle de classeur (`@christophervr/xlsx-core`), de `ooxml-core` et de `ooxml-ui`, et réexporte l’API du modèle (`createWorkbook`, `loadWorkbook`, `saveWorkbook`, ...) : une seule installation et un seul chemin d’import suffisent donc à une application.

| Framework | Paquet                             | Composant                      |
| --------- | ---------------------------------- | ------------------------------ |
| React     | `@christophervr/xlsx-react-viewer` | `SpreadsheetEditor`            |
| Vue 3     | `xlsx-vue-viewer`                  | `SpreadsheetEditor`            |
| Angular   | `xlsx-angular-viewer`              | `SpreadsheetEditorComponent`   |
| Svelte 5  | `xlsx-svelte-viewer`               | default export, `/runtime`     |
| Solid     | `xlsx-solid-viewer`                | `SpreadsheetEditor`            |
| Vanilla   | `xlsx-vanilla-viewer`              | `mountEditor`, `<xlsx-editor>` |

N’installez `@christophervr/xlsx-core` seul que pour un usage sans interface (analyser, calculer et enregistrer sans éditeur).

## Monter un éditeur

```tsx
import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	return (
		<div style={{ height: '100vh' }}>
			<SpreadsheetEditor workbook={workbook} onWorkbookChange={setWorkbook} />
		</div>
	);
}
```

Donnez une hauteur au conteneur : la grille le remplit et défile à l’intérieur. Pour ouvrir un fichier, passez ses octets (`bytes={uint8array}` avec `fileName`), une URL (`src="/files/budget.xlsx"`), ou appelez `load(bytes, fileName)` sur le handle du composant.

## Enregistrer

`save()` renvoie un `Blob` `.xlsx`, `saveBytes('csv')` exporte la feuille active au format CSV, et `download(name?)` enregistre puis lance un téléchargement dans le navigateur. `dirty` signale les modifications non enregistrées ; appelez `markClean()` après avoir conservé vous-même un Blob enregistré. Un fichier `.xls` ancien format est enregistré en `.xlsx`.

Ensuite : les [liaisons de frameworks](/bindings), l’[API de l’élément](/fr/api), les [thèmes](/theming) et les [fonctionnalités et limites](/fr/features).
