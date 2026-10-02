/**
 * Commands the shell itself provides: the File commands, sheet switching, keyboard help and
 * fallbacks for undo/redo, row and column selection and Calculate Now. They are registered before
 * `installCommands` and the grid, so a command module that registers the same id replaces them.
 */
import { MAX_COL, MAX_ROW } from '@christophervr/xlsx-core';
import type { Command } from './commands';
import type { EditorContext } from './context';
import type { FileCommand } from './events';

export interface ShellCommandHandlers {
	fileCommand(command: FileCommand): void;
	openBackstage(page: 'info' | 'saveAs' | 'options' | 'new'): void;
	showShortcuts(): void;
	formulaBarShown(): boolean;
	setFormulaBarShown(shown: boolean): void;
}

const hasWorkbook = (ctx: EditorContext) => Boolean(ctx.workbook());

/** The next visible sheet from the active one in `step` direction, or undefined at the end. */
export function adjacentSheet(ctx: EditorContext, step: 1 | -1): number | undefined {
	const sheets = ctx.workbook()?.sheets ?? [];
	for (let at = ctx.activeSheet() + step; at >= 0 && at < sheets.length; at += step)
		if (sheets[at]?.state === 'visible') return at;
	return undefined;
}

export function shellCommands(handlers: ShellCommandHandlers): Command[] {
	const file = (
		id: string,
		label: string,
		command: FileCommand,
		icon: string,
		shortcut?: string,
	): Command => ({
		id,
		label,
		icon,
		...(shortcut ? { shortcut } : {}),
		enabled: (ctx) => command === 'new' || command === 'open' || hasWorkbook(ctx),
		run: () => handlers.fileCommand(command),
	});
	return [
		file('file.new', 'New workbook', 'new', 'newFile', 'Ctrl+N'),
		file('file.open', 'Open', 'open', 'open', 'Ctrl+O'),
		file('file.save', 'Save', 'save', 'save', 'Ctrl+S'),
		file('file.print', 'Print', 'print', 'print', 'Ctrl+P'),
		file('file.export', 'Save a copy as XLSX', 'export', 'export'),
		file('file.export-csv', 'Export as CSV', 'exportCsv', 'export'),
		{
			id: 'file.save-as',
			label: 'Save As',
			icon: 'copy',
			enabled: hasWorkbook,
			run: () => handlers.openBackstage('saveAs'),
		},
		{ id: 'file.info', label: 'Info', icon: 'info', run: () => handlers.openBackstage('info') },
		{
			id: 'file.options',
			label: 'Options',
			icon: 'settings',
			run: () => handlers.openBackstage('options'),
		},
		{
			id: 'edit.undo',
			label: 'Undo',
			icon: 'undo',
			shortcut: 'Ctrl+Z',
			editing: true,
			enabled: (ctx) => ctx.session()?.canUndo() ?? false,
			run: (ctx) => void ctx.session()?.undo(),
		},
		{
			id: 'edit.redo',
			label: 'Redo',
			icon: 'redo',
			shortcut: 'Ctrl+Y',
			editing: true,
			enabled: (ctx) => ctx.session()?.canRedo() ?? false,
			run: (ctx) => void ctx.session()?.redo(),
		},
		{
			id: 'sheet.next',
			label: 'Next sheet',
			shortcut: 'Ctrl+PgDn',
			enabled: (ctx) => adjacentSheet(ctx, 1) !== undefined,
			run: (ctx) => {
				const next = adjacentSheet(ctx, 1);
				if (next !== undefined) ctx.setActiveSheet(next);
			},
		},
		{
			id: 'sheet.previous',
			label: 'Previous sheet',
			shortcut: 'Ctrl+PgUp',
			enabled: (ctx) => adjacentSheet(ctx, -1) !== undefined,
			run: (ctx) => {
				const previous = adjacentSheet(ctx, -1);
				if (previous !== undefined) ctx.setActiveSheet(previous);
			},
		},
		{
			id: 'edit.select-column',
			label: 'Select the entire column',
			shortcut: 'Ctrl+Space',
			enabled: hasWorkbook,
			run: (ctx) => {
				const { ranges, active } = ctx.selection.get();
				const cols = ranges.length ? ranges : [{ start: active, end: active }];
				ctx.selection.set({
					active,
					anchor: active,
					ranges: cols.map((range) => ({
						start: { row: 0, col: range.start.col },
						end: { row: MAX_ROW, col: range.end.col },
					})),
				});
			},
		},
		{
			id: 'edit.select-row',
			label: 'Select the entire row',
			shortcut: 'Shift+Space',
			enabled: hasWorkbook,
			run: (ctx) => {
				const { ranges, active } = ctx.selection.get();
				const rows = ranges.length ? ranges : [{ start: active, end: active }];
				ctx.selection.set({
					active,
					anchor: active,
					ranges: rows.map((range) => ({
						start: { row: range.start.row, col: 0 },
						end: { row: range.end.row, col: MAX_COL },
					})),
				});
			},
		},
		{
			id: 'view.formula-bar',
			label: 'Formula Bar',
			icon: 'formulaBar',
			checked: () => handlers.formulaBarShown(),
			run: () => handlers.setFormulaBarShown(!handlers.formulaBarShown()),
		},
		{
			id: 'help.keyboard-shortcuts',
			label: 'Keyboard shortcuts',
			icon: 'keyboard',
			shortcut: 'Ctrl+/',
			run: () => handlers.showShortcuts(),
		},
	];
}
