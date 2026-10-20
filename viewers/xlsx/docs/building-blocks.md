# Compose a spreadsheet editor

`@christophervr/xlsx-react-viewer` exports `useViewerBuildingBlocks`, `SpreadsheetCanvas`
and optional `Toolbar`, `FormulaBar` and `SheetTabs` components. All use the existing workbook
session, grid, formulas, validation, protection and history.

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Render the canvas while loading; `toolbarProps` is null until mount. Give its container a
height. The canvas mounts only the grid. File, settings, sharing, title bar, ribbon, status
bar, formula bar and sheet tabs are absent unless the host supplies separate controls.

`toolbarProps` contains selection/active cell, sheet state, formula text, formatting and
`enabled` states. Connect `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange`,
`onFontSizeChange`, `onNumberFormatChange`, `onAlignmentChange`, `onUndo` and `onRedo`
to custom controls. Formatting fields are null for mixed values. Formatting inspection is
bounded to 4096 cells: larger selections return null and `formattingComplete=false`.

`onSelect`, `onSheetChange`, `onFormulaChange` and `commitEdit` return false when cell validation
blocks the action. Formatting callbacks return a Promise<boolean>. They commit an active
cell edit before acting, and retain its selection when custom controls take focus. Formula
text is committed on Enter by the optional FormulaBar; a rejected entry sets aria-invalid.
The existing handle retains save/serialization and dirty tracking. A save rejects if input
cannot be committed. Permissions and worksheet protection remain authoritative.

For every framework, `mountEditorSurface(host, options)` exposes the same surface, existing
handle, `getToolbarProps()`, `subscribe(listener)`, `update` and `destroy`. Hosts own file
and optional dialog workflows; no hidden library toolbar is required.

`formula`, `formulaCaret`, `editing` and `editMode` read the grid's live edit bridge.
`onFormulaInput(text, caret?)` updates that same draft; `onFormulaChange(text)` commits it.
A rejected draft can be corrected through `onFormulaInput` without first committing the invalid
value. `onCancelEdit()` cancels it. There is no separate formula state store.

The composition API requires the matching `ooxml-ui` runtime, released with these additions.
Existing full-viewer imports remain compatible with the previous published runtime. An older
runtime rejects a composition request explicitly instead of mounting application UI by accident.
