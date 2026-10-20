# Compose a Word editor

`docx-react-viewer` exports `useViewerBuildingBlocks`, `DocumentCanvas` and an optional
`Toolbar`. They use WordEditor's existing editor, layout, history and serialization engine.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		documentModel,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
	});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Render the canvas even while loading. `toolbarProps` is null until it mounts. It exposes
selection, bold/italic/underline (null for mixed values), font family/size, alignment,
list state, undo/redo availability and direct `on*` callbacks. `Toolbar` consumes the same
props when library controls are wanted. Give the canvas container a height.

Use `bytes` to load a file, or the existing handle's `load`, `save`, `download`, `markClean`
and `dirty`. Saving does not mark a document clean until the host confirms persistence.
`canEdit={false}` or `readOnly` blocks editing while selection and viewing remain available.
The canvas mounts no title bar, ribbon, backstage, status bar or application dialogs.
Hosts own optional dialogs and file/settings/sharing flows.

Every framework package also exports `mountEditorSurface(host, options)`. Its binding exposes
`getToolbarProps()` and `subscribe(listener)`, plus the existing editor handle and
`update`/`destroy`. This is the framework-neutral equivalent for native lifecycle adapters:
subscribe to selection/format changes, wire controls directly, and destroy on unmount.

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.
