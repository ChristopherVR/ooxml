# Solid

`visio-solid-viewer` vous donne le `<visio-viewer>` partagé comme intégration native Solid.

```bash
npm install visio-solid-viewer solid-js
```

```tsx
import { VisioViewer } from 'visio-solid-viewer';
import type { VisioDocument } from 'visio-solid-viewer';

export function Diagram(props: { diagram: VisioDocument }) {
	return (
		<VisioViewer
			document={props.diagram}
			style={{ height: '600px' }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

Passez `viewerRef` pour recevoir le handle au montage (et `undefined` au démontage).

Les propriétés (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) et le handle sont identiques dans chaque framework : voir le [contrat des liaisons](/fr/bindings) et l'[API du visualiseur](/fr/api). Pour composer votre propre interface à partir des éléments du visualiseur, voir les [blocs de construction](/fr/building-blocks) ; pour partager un dessin, voir [collaboration](/fr/collaboration). [Essayer la démo Solid](/demo-solid/){target="_self"}.
