// The editing commands the grid owns (`edit.*`, `grid.pick-from-list`), registered in the shell's
// command registry so the ribbon, menus and shortcuts run them by id.
import {
	currentRegion,
	normalizeRange,
	type CellRange,
	type PasteMode,
	type Worksheet,
} from '@christophervr/xlsx-core';
import type { GridClipboard } from '../clipboard.js';
import type { Command } from '../commands.js';
import type { EditorContext, Selection } from '../context.js';

export interface CommandHost {
	clipboard: GridClipboard;
	cancelEdit(): void;
	isEditing(): boolean;
	selectAll(): void;
	openValidationList(): boolean;
}

const PASTE_MODES: readonly PasteMode[] = ['all', 'values', 'formats', 'formulas', 'transpose'];

/** Ctrl+A: the current region first, everything when the region is already selected. */
export function currentRegionOrAll(sheet: Worksheet, selection: Selection): CellRange | undefined {
	const region = normalizeRange(currentRegion(sheet, selection.active));
	const single = region.start.row === region.end.row && region.start.col === region.end.col;
	const last = selection.ranges[selection.ranges.length - 1];
	const same =
		last &&
		last.start.row === region.start.row &&
		last.start.col === region.start.col &&
		last.end.row === region.end.row &&
		last.end.col === region.end.col;
	return single || same ? undefined : region;
}

/** Fill down / right: the first row (column) of each range, or the one before a single row. */
export function fillPlan(
	range: CellRange,
	axis: 'down' | 'right',
): { source: CellRange; target: CellRange } | undefined {
	const r = normalizeRange(range);
	if (axis === 'down') {
		if (r.start.row === r.end.row) {
			if (r.start.row === 0) return undefined;
			const source = {
				start: { row: r.start.row - 1, col: r.start.col },
				end: { row: r.start.row - 1, col: r.end.col },
			};
			return { source, target: { start: source.start, end: r.end } };
		}
		return { source: { start: r.start, end: { row: r.start.row, col: r.end.col } }, target: r };
	}
	if (r.start.col === r.end.col) {
		if (r.start.col === 0) return undefined;
		const source = {
			start: { row: r.start.row, col: r.start.col - 1 },
			end: { row: r.end.row, col: r.start.col - 1 },
		};
		return { source, target: { start: source.start, end: r.end } };
	}
	return { source: { start: r.start, end: { row: r.end.row, col: r.start.col } }, target: r };
}

const hasSession = (ctx: EditorContext): boolean => ctx.session() !== undefined;

export function gridCommands(host: CommandHost): Command[] {
	const fill = (ctx: EditorContext, axis: 'down' | 'right') => {
		const session = ctx.session();
		if (!session) return;
		const sheet = ctx.activeSheet();
		session.batch(axis === 'down' ? 'Fill Down' : 'Fill Right', () => {
			for (const range of ctx.selection.get().ranges) {
				const plan = fillPlan(range, axis);
				if (plan) session.fill(sheet, plan.source, plan.target);
			}
		});
	};
	return [
		{
			id: 'edit.copy',
			label: 'Copy',
			icon: 'copy',
			shortcut: 'Ctrl+C',
			enabled: hasSession,
			run: () => void host.clipboard.writeSystem(false),
		},
		{
			id: 'edit.cut',
			label: 'Cut',
			icon: 'cut',
			shortcut: 'Ctrl+X',
			editing: true,
			enabled: hasSession,
			run: () => void host.clipboard.writeSystem(true),
		},
		{
			id: 'edit.paste',
			label: 'Paste',
			icon: 'paste',
			shortcut: 'Ctrl+V',
			editing: true,
			enabled: hasSession,
			run: async (_ctx, arg) => {
				const mode =
					typeof arg === 'string' && PASTE_MODES.includes(arg as PasteMode)
						? (arg as PasteMode)
						: 'all';
				await host.clipboard.pasteSystem(mode);
			},
		},
		{
			id: 'edit.paste-values',
			label: 'Paste Values',
			icon: 'paste',
			editing: true,
			enabled: hasSession,
			run: () => host.clipboard.pasteSystem('values').then(() => undefined),
		},
		{
			id: 'edit.paste-special',
			label: 'Paste Special...',
			icon: 'paste',
			shortcut: 'Ctrl+Alt+V',
			editing: true,
			enabled: hasSession,
			run: async (ctx) => {
				const result = await ctx.dialogs.open<unknown>('paste-special');
				const mode =
					typeof result === 'string'
						? result
						: result && typeof result === 'object' && 'mode' in result
							? String((result as { mode: unknown }).mode)
							: undefined;
				if (mode && PASTE_MODES.includes(mode as PasteMode))
					await host.clipboard.pasteSystem(mode as PasteMode);
			},
		},
		{
			id: 'edit.undo',
			label: 'Undo',
			icon: 'undo',
			shortcut: 'Ctrl+Z',
			editing: true,
			enabled: (ctx) => Boolean(ctx.session()?.canUndo()) || host.isEditing(),
			run: (ctx) => {
				if (host.isEditing()) return host.cancelEdit();
				ctx.session()?.undo();
			},
		},
		{
			id: 'edit.redo',
			label: 'Redo',
			icon: 'redo',
			shortcut: 'Ctrl+Y',
			editing: true,
			enabled: (ctx) => Boolean(ctx.session()?.canRedo()),
			run: (ctx) => void ctx.session()?.redo(),
		},
		{
			id: 'edit.select-all',
			label: 'Select All',
			shortcut: 'Ctrl+A',
			enabled: hasSession,
			run: () => host.selectAll(),
		},
		{
			id: 'edit.delete',
			label: 'Clear Contents',
			icon: 'eraser',
			shortcut: 'Delete',
			editing: true,
			enabled: hasSession,
			run: (ctx) => {
				const session = ctx.session();
				if (!session) return;
				const sheet = ctx.activeSheet();
				const ranges = ctx.selection.get().ranges;
				session.batch('Clear Contents', () => {
					for (const range of ranges) session.clearRange(sheet, range, 'contents');
				});
			},
		},
		{
			id: 'edit.fill-down',
			label: 'Fill Down',
			shortcut: 'Ctrl+D',
			editing: true,
			enabled: hasSession,
			run: (ctx) => fill(ctx, 'down'),
		},
		{
			id: 'edit.fill-right',
			label: 'Fill Right',
			shortcut: 'Ctrl+R',
			editing: true,
			enabled: hasSession,
			run: (ctx) => fill(ctx, 'right'),
		},
		{
			id: 'grid.pick-from-list',
			label: 'Pick From Drop-down List...',
			shortcut: 'Alt+Down',
			editing: true,
			enabled: hasSession,
			run: () => void host.openValidationList(),
		},
	];
}
