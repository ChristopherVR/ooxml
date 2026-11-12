/**
 * What a host receives to build its own context menu, Format Cells, Find and Replace and Insert
 * Function UI on a surface-only editor. Every operation re-reads read-only state, sheet
 * protection and the pending cell edit when it runs, goes through the edit session (so undo,
 * redo, dirty state and change events behave as for the bundled UI), and reports a refusal by
 * its return value.
 */
import type { FindMatch, FindQuery, StylePatch } from 'ooxml-core/xlsx';
import type { FormatCellsState } from 'ooxml-core/xlsx/ui';
import type { SurfaceActionContext } from './surface-actions';

export type ContextMenuKind = 'cell' | 'row' | 'column' | 'tab';

export interface ContextMenuItem {
	/** Stable id, pass it to `execute`. */
	id: string;
	/** Localized label. */
	label: string;
	/** The editor command the item runs, when it runs one. */
	command?: string;
	/** False when the item cannot run now (read-only, protection, command missing in this mode). */
	enabled: boolean;
	checked?: boolean;
	shortcut?: string;
	/** A separator belongs above this item. */
	separatorBefore?: boolean;
}

export interface ContextMenuModel extends SurfaceActionContext {
	menu: ContextMenuKind;
	/** The sheet the menu is for (the clicked tab for `tab`). */
	menuSheet: number;
	readOnly: boolean;
	items: ContextMenuItem[];
	/**
	 * Runs item `id` after re-reading its state; resolves false when the item is unknown, disabled
	 * or refused. A tab menu activates its sheet first.
	 */
	execute(id: string): Promise<boolean>;
}

/** Passed to `onContextMenu`; `x` / `y` are client coordinates of the right click. */
export interface ContextMenuRequest extends ContextMenuModel {
	x: number;
	y: number;
}

export interface FormatCellsOperations {
	/** The formatting of the selection: shared properties, and the names of the mixed ones. */
	getState(): FormatCellsState;
	/** Whether `apply` would run now. */
	canApply(): boolean;
	/** Applies a style patch to the current selection as one undo step. False when refused. */
	apply(patch: StylePatch): boolean;
}
export interface FormatCellsRequest extends SurfaceActionContext, FormatCellsOperations {}

export interface FindOperations {
	/** Replace and Replace All work only when the workbook is editable (finding never needs it). */
	canReplace(): boolean;
	/** Every match; works in read-only workbooks. */
	findAll(query: FindQuery): FindMatch[];
	/** Selects the next (or previous) match after the active cell, wrapping; undefined when none. */
	findNext(query: FindQuery, options?: { backwards?: boolean }): FindMatch | undefined;
	/** Selects a match from `findAll`. */
	select(match: FindMatch): boolean;
	/**
	 * Replaces the match at the active cell (when there is one) and selects the next; without a
	 * match there it only moves on. True when a cell was replaced.
	 */
	replace(query: FindQuery, replacement: string): boolean;
	/** Replaces every match in cells that may be edited; returns how many cells changed. */
	replaceAll(query: FindQuery, replacement: string): number;
}
export interface FindReplaceRequest extends SurfaceActionContext, FindOperations {}

export interface FunctionDoc {
	name: string;
	category: string;
	/** For example `SUM(number1, [number2], ...)`. */
	syntax: string;
	/** Localized description. */
	description: string;
}

export interface InsertFunctionOperations {
	categories(): string[];
	/** Functions matching `text` (name first, then description words) within `category`. */
	search(options?: { text?: string; category?: 'all' | 'recent' | string }): FunctionDoc[];
	/** Whether `insert` would run now. */
	canInsert(): boolean;
	/**
	 * Inserts `NAME(` (or `NAME(args)` when `args` is given) at the caret of the active formula
	 * draft, starting a draft when none is open. Returns the draft text and caret afterwards, or
	 * undefined when refused (read-only, unknown function, or a draft that is plain text).
	 */
	insert(name: string, args?: string): { text: string; caret: number } | undefined;
}
export interface InsertFunctionRequest extends SurfaceActionContext, InsertFunctionOperations {}
