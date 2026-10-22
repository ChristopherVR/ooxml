# Vanilla JavaScript

`visio-vanilla-viewer` gives you the shared `<visio-viewer>` as a native Vanilla JavaScript integration.

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

`mountViewer` returns the full shared binding with `update()` and an idempotent `destroy()`.

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `ribbonAddIns`, `collaboration`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). To compose your own UI from the viewer's parts instead, see the [building blocks](/building-blocks); to share a drawing, see [collaboration](/collaboration). [Try the Vanilla JavaScript demo](/demo/){target="_self"}.
