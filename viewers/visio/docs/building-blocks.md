# Compose a diagram viewer

`visio-react-viewer` exports `useViewerBuildingBlocks`, `DiagramCanvas`, and optional
`Toolbar`, `PageNavigation` and `Inspector`. These reuse ViewerController and the existing
SVG renderer and selection inspector.

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` accepts the same bytes/Blob source as the viewer. `handle` receives the existing
ViewerController; use `load`, `applyEdits`, `exportVsdx` and its reactive `state.edit.dirty`
for host persistence. `onDocumentChange`, `onDirtyChange` and `onDocumentError` are optional.
Give the canvas container a height and keep it mounted while loading.

The flat `toolbarProps` exposes page, zoom, selection, effective edit availability, undo/redo,
fit, search and the supported duplicate/group/ungroup/plain-text editing callbacks.
`Inspector` consumes `inspectorProps`. The canvas itself mounts no title bar, ribbon, page
controls, stencil/shape panes, inspector, file UI, recovery, settings or sharing dialogs.
Hosts own those optional workflows. Existing VSDX editing/protection limits still apply.

`canEdit={false}` is also available on the full viewer in every framework. It rejects local
mutation methods and editing gestures, including pending edits when permission is revoked.
Page navigation, selection, search, scrolling/pan, zoom, fit and serialization remain available.
Source-format editing support is separate from host permission.

Every framework package exports `mountDiagramCanvas(host, controller)`,
`diagramToolbarProps(controller, fit, locale?)` and `mountDiagramInspector(host, controller)`.
Subscribe through the existing controller, and call returned disposers on unmount. No second
state store is introduced.
