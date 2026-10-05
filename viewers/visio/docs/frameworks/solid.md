# Solid

`visio-solid-viewer` gives you the shared `<visio-viewer>` as a native Solid integration.

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

Pass `viewerRef` to receive the handle on mount (and `undefined` on unmount).

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). [Try the Solid demo](/demo-solid/){target="_self"}.
