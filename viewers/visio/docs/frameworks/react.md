# React

`visio-react-viewer` gives you the shared `<visio-viewer>` as a native React integration.

```bash
npm install visio-react-viewer react
```

```tsx
import { useRef } from 'react';
import { VisioViewer, type ViewerHandle, type VisioDocument } from 'visio-react-viewer';

export function Diagram({ diagram }: { diagram: VisioDocument }) {
	const ref = useRef<ViewerHandle>(null);
	return (
		<VisioViewer
			ref={ref}
			document={diagram}
			style={{ height: 600 }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

A forwarded `ref` exposes the `ViewerHandle`. React `className` and `style` style the host, which needs a height.

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). [Try the React demo](/demo-react/){target="_self"}.
