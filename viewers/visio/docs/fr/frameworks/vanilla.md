# JavaScript vanilla

`visio-vanilla-viewer` vous donne le `<visio-viewer>` partagé comme intégration native en JavaScript vanilla.

```bash
npm install visio-vanilla-viewer
```

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	document: diagram,
	zoom: 1,
	events: { 'document-error': console.error },
});

viewer.update({ pageIndex: 1 });
await viewer.load(file);
viewer.destroy(); // idempotent
```

`mountViewer` renvoie la liaison partagée complète, avec `update()` et un `destroy()` idempotent.

Les propriétés (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) et le handle sont identiques dans chaque framework : voir le [contrat des liaisons](/fr/bindings) et l'[API du visualiseur](/fr/api). Pour composer votre propre interface à partir des éléments du visualiseur, voir les [blocs de construction](/fr/building-blocks) ; pour partager un dessin, voir [collaboration](/fr/collaboration). [Essayer la démo JavaScript vanilla](/demo/){target="_self"}.
