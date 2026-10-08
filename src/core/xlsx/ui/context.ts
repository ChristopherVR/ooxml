// The editor context every UI module (grid, formula bar, sheet tabs, commands, dialogs) receives.
// Shapes are fixed by the UI contract; the shell (component.ts) is the only implementation.
import type { CellAddress, CellRange, EditSession, RemoteRange, Workbook } from '../index';
import type { CommandRegistry } from './commands';
import type { DialogRegistry } from './dialogs';

export interface Selection {
	sheet: number;
	active: CellAddress;
	anchor: CellAddress;
	ranges: CellRange[];
	/**
	 * Index in the sheet's drawings of the selected picture or chart. While set, the drawing (not
	 * the cells) is the selection: the Chart Design tab follows it and Delete removes it. Any
	 * cell selection change clears it.
	 */
	drawing?: number | undefined;
}

export interface SelectionModel {
	get(): Selection;
	set(next: Partial<Selection> & { ranges?: CellRange[] }): void;
	onChange(listener: (selection: Selection) => void): () => void;
}

/** Implemented by the grid module and attached through `ctx.attachGrid`. */
export interface GridController {
	focus(): void;
	scrollTo(address: CellAddress): void;
	/** Repaints after a model change. */
	invalidate(): void;
	/** Starts the in-cell editor (F2 or typing). */
	beginEdit(initialText?: string): void;
	commitEdit(): boolean;
	cancelEdit(): void;
	isEditing(): boolean;
	/** Width in CSS pixels of `text` drawn with a CSS font shorthand. */
	measureText(text: string, font: string): number;
	zoom(): number;
	setZoom(percent: number): void;
}

export interface EditorContext {
	readonly host: HTMLElement;
	readonly root: ShadowRoot;
	/** Undefined until a workbook is loaded. */
	session(): EditSession | undefined;
	workbook(): Workbook | undefined;
	activeSheet(): number;
	setActiveSheet(index: number): void;
	readonly selection: SelectionModel;
	readOnly(): boolean;
	/** English text is the key, like the Word editor. */
	t(key: string, vars?: Record<string, string | number>): string;
	readonly commands: CommandRegistry;
	readonly dialogs: DialogRegistry;
	grid(): GridController | undefined;
	attachGrid(grid: GridController): void;
	/** Fires after every session change, load and sheet switch. */
	onModelChange(listener: (change: unknown) => void): () => void;
	/** Refreshes ribbon and status bar state (debounced). */
	requestRender(): void;
	toast(message: string, kind?: 'info' | 'warning' | 'error'): void;
	/** Dispatches a public CustomEvent on the host; returns `!defaultPrevented`. */
	emit(type: string, detail: unknown): boolean;
	authorName(): string;
	locale(): string;
	/**
	 * The undo history Ctrl+Z, the ribbon and the Quick Access Toolbar drive: the edit session's
	 * own, or a collaboration binding's while the workbook is shared (so undo never publishes a
	 * snapshot restore as a local edit). Absent means the session's; read it through `historyOf`.
	 */
	history?(): EditHistory | undefined;
	/** Collaborators' selections on every sheet; the grid outlines those on the shown sheet. */
	remoteSelections?(): readonly RemoteRange[];
	/** Fires when collaborators' selections (or who is in the room) change. */
	onRemoteSelectionsChange?(listener: () => void): () => void;
}

/** The part of an edit session the undo and redo commands use. */
export type EditHistory = Pick<
	EditSession,
	'undo' | 'redo' | 'canUndo' | 'canRedo' | 'undoLabel' | 'redoLabel'
>;

/** The history undo and redo go through: the context's override, else the edit session. */
export function historyOf(ctx: EditorContext): EditHistory | undefined {
	return ctx.history ? ctx.history() : ctx.session();
}
