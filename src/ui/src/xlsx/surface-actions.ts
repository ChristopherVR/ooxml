/**
 * Host callbacks for the workflows the bundled application opens a dialog for. A surface-only
 * editor never mounts those dialogs: it calls the host's callback instead, and reports the
 * workflow as not enabled when the host gave none.
 */
import type {
	ContextMenuCustomization,
	ContextMenuRequest,
	FindReplaceRequest,
	FormatCellsRequest,
	InsertFunctionRequest,
} from './surface-workflows';

export interface SurfaceActionContext {
	sheet: number;
	sheetName: string;
	/** The selection as A1 text. */
	selection: string;
	/** The active cell as A1 text. */
	activeCell: string;
}

export interface SurfaceHostActions {
	/** Share / collaboration settings (File > Share in the application). */
	onOpenShareDialog?: ((context: SurfaceActionContext) => void) | undefined;
	/** Application settings and options. */
	onOpenSettings?: ((context: SurfaceActionContext) => void) | undefined;
	/** The Format Cells dialog. */
	onOpenFormatCells?: ((request: FormatCellsRequest) => void) | undefined;
	/** Find and Replace. */
	onOpenFindReplace?: ((request: FindReplaceRequest) => void) | undefined;
	/** Insert Function. */
	onOpenInsertFunction?: ((request: InsertFunctionRequest) => void) | undefined;
	/**
	 * A right click on a cell, a row or column header or a sheet tab (and Shift+F10 or the menu key
	 * on the grid). Return true to show your own menu and suppress the native one; build it from
	 * `request.items` and run a choice with `request.execute(id)`. Any other return value shows
	 * the native menu.
	 */
	onContextMenu?: ((request: ContextMenuRequest) => boolean | void) | undefined;
	/**
	 * Filters and extends the NATIVE menu, which still draws. Runs for the same menus as
	 * `onContextMenu` when that callback did not take the menu over. Return `{ hide, add }`:
	 * `hide` lists native item ids (see `request.items`) to leave out, `add` lists host items with
	 * an id, label, enabled state and `onSelect`. Native items keep their disabled state and
	 * read-only checks; host items are rechecked when they run.
	 */
	onContextMenuItems?:
		| ((request: ContextMenuRequest) => ContextMenuCustomization | void)
		| undefined;
	/**
	 * The data-validation alert of a cell edit that was refused. The host shows it (any UI) and
	 * resolves the request with one of its methods; the grid never draws its own alert here.
	 * Without this callback the edit stays open as if Retry was chosen (see `ValidationAlertRequest`).
	 */
	onOpenValidationAlert?: ((request: ValidationAlertRequest) => void) | undefined;
}
/** The workflows that receive a `SurfaceActionContext` (the validation alert has its own request). */
export type SurfaceActionName = Exclude<
	keyof SurfaceHostActions,
	'onOpenValidationAlert' | 'onContextMenu' | 'onContextMenuItems'
>;

interface ValidationAlertBase {
	/** The cell being edited, as A1 text. */
	ref: string;
	sheet: number;
	sheetName: string;
	/** The text the user entered. */
	value: string;
	/** The rule's error title, or an empty string for the default. */
	title: string;
	/** The (localized) message to show. */
	message: string;
	/** Cancel the edit: the cell keeps its old value. */
	cancel(): void;
}
/** Excel's Stop alert: Retry or Cancel. */
export interface StopValidationAlert extends ValidationAlertBase {
	style: 'stop';
	/** Keep the editor open with the draft. */
	retry(): void;
}
/** Excel's Warning alert: Yes (accept the value), No (retry) or Cancel. */
export interface WarningValidationAlert extends ValidationAlertBase {
	style: 'warning';
	/** Yes: commit the invalid value. */
	accept(): void;
	/** No: keep the editor open with the draft. */
	retry(): void;
}
/** Excel's Information alert: OK (accept the value) or Cancel. */
export interface InformationValidationAlert extends ValidationAlertBase {
	style: 'information';
	/** OK: commit the invalid value. */
	accept(): void;
}
/**
 * Passed to `onOpenValidationAlert`. Resolve it exactly once; later calls are ignored. Until it
 * is resolved the cell editor ignores further commits.
 */
export type ValidationAlertRequest =
	| StopValidationAlert
	| WarningValidationAlert
	| InformationValidationAlert;
