# Angular

`visio-angular-viewer` vous donne le `<visio-viewer>` partagé comme intégration native Angular.

```bash
npm install visio-angular-viewer @angular/core
```

```ts
import { Component, Input } from '@angular/core';
import { VisioViewerComponent } from 'visio-angular-viewer';
import type { VisioDocument } from 'visio-angular-viewer';

@Component({
	selector: 'app-diagram',
	imports: [VisioViewerComponent],
	template: '<visio-viewer-host [document]="diagram" (documentError)="onError($event)" />',
	styles: ['visio-viewer-host { display: block; height: 600px }'],
})
export class Diagram {
	@Input() diagram: VisioDocument | null = null;
	onError(error: Error) {
		console.error(error);
	}
}
```

Les sorties sont `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` et `fileCommand`. `fileCommand` se contente de signaler une commande du menu Fichier ; pour l'annuler, renvoyez `false` depuis le rappel `events['file-command']`. La table de rappels `events` complète est aussi acceptée comme entrée.

Les propriétés (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) et le handle sont identiques dans chaque framework : voir le [contrat des liaisons](/fr/bindings) et l'[API du visualiseur](/fr/api). Pour composer votre propre interface à partir des éléments du visualiseur, voir les [blocs de construction](/fr/building-blocks) ; pour partager un dessin, voir [collaboration](/fr/collaboration). [Essayer la démo Angular](/demo-angular/){target="_self"}.
