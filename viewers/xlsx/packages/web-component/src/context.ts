// The editor context every UI module (grid, formula bar, sheet tabs, commands, dialogs) receives.
// Shapes are fixed by the UI contract; the shell (component.ts) is the only implementation.
import type { CellAddress, CellRange, EditSession, Workbook } from '@christophervr/xlsx-core';
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
}
