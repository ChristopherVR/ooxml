/**
 * Host callbacks for the workflows the bundled application opens a dialog for. A surface-only
 * editor never mounts those dialogs: it calls the host's callback instead, and reports the
 * workflow as not enabled when the host gave none.
 */
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
	onOpenFormatCells?: ((context: SurfaceActionContext) => void) | undefined;
	/** Find and Replace. */
	onOpenFindReplace?: ((context: SurfaceActionContext) => void) | undefined;
	/** Insert Function. */
	onOpenInsertFunction?: ((context: SurfaceActionContext) => void) | undefined;
}
export type SurfaceActionName = keyof SurfaceHostActions;
