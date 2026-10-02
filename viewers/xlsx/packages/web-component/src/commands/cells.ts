// Home > Cells: insert and delete (cells, rows, columns, sheets) and the Format menu (row height,
// column width, autofit, hide and unhide, sheet organisation, protection, Format Cells).
import { nextSheetName, type Color } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import { sheetBaseName } from '../localization.js';
import type { EditorContext } from '../context.js';
import { style } from './font.js';
import { icon } from './icons.js';
import {
	activeStyle,
	colsOf,
	editing,
	measure,
	rowsOf,
	structureLocked,
	target,
	wholeColumns,
	wholeRows,
} from './util.js';

const rowsCount = (ctx: EditorContext) => {
	const t = target(ctx);
	return t
		? { t, at: t.range.start.row, count: t.range.end.row - t.range.start.row + 1 }
		: undefined;
};
const colsCount = (ctx: EditorContext) => {
	const t = target(ctx);
	return t
		? { t, at: t.range.start.col, count: t.range.end.col - t.range.start.col + 1 }
		: undefined;
};

const visibleSheets = (ctx: EditorContext): number =>
	ctx.workbook()?.sheets.filter((s) => s.state === 'visible').length ?? 0;

function sheetCommand(command: Omit<Command, 'editing'>): Command {
	return editing({
		...command,
		lock: false,
		enabled: (ctx) => !structureLocked(ctx) && (command.enabled?.(ctx) ?? true),
	});
}

export function cellCommands(): Command[] {
	return [
		// The Cells group's split buttons show Excel's short captions.
		editing({
			id: 'cells.insert-split',
			label: 'Insert',
			icon: icon('insertCells'),
			lock: 'insertRows',
			run: (ctx) => void ctx.commands.run('cells.insert'),
		}),
		editing({
			id: 'cells.delete-split',
			label: 'Delete',
			icon: icon('deleteCells'),
			lock: 'deleteRows',
			run: (ctx) => void ctx.commands.run('cells.delete'),
		}),
		editing({
			id: 'cells.insert',
			label: 'Insert Cells...',
			icon: icon('insertCells'),
			shortcut: 'Ctrl+Shift+=',
			lock: 'insertRows',
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				if (wholeRows(t.range)) return void ctx.commands.run('cells.insert-rows');
				if (wholeColumns(t.range)) return void ctx.commands.run('cells.insert-columns');
				void ctx.dialogs.open('insert-cells');
			},
		}),
		editing({
			id: 'cells.delete',
			label: 'Delete Cells...',
			icon: icon('deleteCells'),
			shortcut: 'Ctrl+-',
			lock: 'deleteRows',
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				if (wholeRows(t.range)) return void ctx.commands.run('cells.delete-rows');
				if (wholeColumns(t.range)) return void ctx.commands.run('cells.delete-columns');
				void ctx.dialogs.open('delete-cells');
			},
		}),
		editing({
			id: 'cells.insert-rows',
			label: 'Insert Sheet Rows',
			icon: icon('insertCells'),
			lock: 'insertRows',
			run: (ctx) => {
				const r = rowsCount(ctx);
				if (r) r.t.session.insertRows(r.t.sheet, r.at, r.count);
			},
		}),
		editing({
			id: 'cells.delete-rows',
			label: 'Delete Sheet Rows',
			icon: icon('deleteCells'),
			lock: 'deleteRows',
			run: (ctx) => {
				const r = rowsCount(ctx);
				if (r) r.t.session.deleteRows(r.t.sheet, r.at, r.count);
			},
		}),
		editing({
			id: 'cells.insert-columns',
			label: 'Insert Sheet Columns',
			icon: icon('insertCells'),
			lock: 'insertColumns',
			run: (ctx) => {
				const c = colsCount(ctx);
				if (c) c.t.session.insertColumns(c.t.sheet, c.at, c.count);
			},
		}),
		editing({
			id: 'cells.delete-columns',
			label: 'Delete Sheet Columns',
			icon: icon('deleteCells'),
			lock: 'deleteColumns',
			run: (ctx) => {
				const c = colsCount(ctx);
				if (c) c.t.session.deleteColumns(c.t.sheet, c.at, c.count);
			},
		}),
		sheetCommand({
			id: 'sheet.insert',
			label: 'Insert Sheet',
			icon: icon('sheet'),
			shortcut: 'Shift+F11',
			run: (ctx) => {
				const session = ctx.session();
				if (!session) return;
				const name = nextSheetName(session.workbook, sheetBaseName(ctx.t));
				const index = session.addSheet(name, ctx.activeSheet());
				ctx.setActiveSheet(index);
			},
		}),
		sheetCommand({
			id: 'sheet.delete',
			label: 'Delete Sheet',
			icon: icon('deleteCells'),
			enabled: (ctx) => visibleSheets(ctx) > 1,
			run: async (ctx) => {
				const session = ctx.session();
				const ws = session?.workbook.sheets[ctx.activeSheet()];
				if (!session || !ws) return;
				if (ws.rows.size) {
					const ok = await ctx.dialogs.open<boolean>('confirm', {
						heading: 'Delete Sheet',
						message: 'This sheet will be permanently deleted. Do you want to continue?',
						okLabel: 'Delete',
					});
					if (!ok) return;
				}
				const index = ctx.activeSheet();
				session.deleteSheet(index);
				const next = session.workbook.sheets.findIndex(
					(s, i) =>
						i >= Math.min(index, session.workbook.sheets.length - 1) && s.state === 'visible',
				);
				ctx.setActiveSheet(next >= 0 ? next : 0);
			},
		}),
		sheetCommand({
			id: 'sheet.rename',
			label: 'Rename Sheet',
			icon: icon('rename'),
			run: (ctx, arg) => void ctx.dialogs.open('rename-sheet', arg),
		}),
		sheetCommand({
			id: 'sheet.move-copy',
			label: 'Move or Copy Sheet...',
			icon: icon('moveCopy'),
			run: (ctx) => void ctx.dialogs.open('move-copy-sheet'),
		}),
		editing({
			id: 'sheet.tab-color',
			label: 'Tab Color',
			icon: icon('tabColor'),
			lock: false,
			run: (ctx, arg) => {
				const session = ctx.session();
				if (!session) return;
				if (arg === undefined) return void ctx.dialogs.open('tab-color');
				const color =
					arg && typeof arg === 'object' && !(arg as Color).auto ? (arg as Color) : undefined;
				session.setTabColor(ctx.activeSheet(), color);
			},
		}),
		sheetCommand({
			id: 'sheet.hide',
			label: 'Hide Sheet',
			icon: icon('hide'),
			enabled: (ctx) => visibleSheets(ctx) > 1,
			run: (ctx) => {
				const session = ctx.session();
				if (!session) return;
				const index = ctx.activeSheet();
				session.setSheetState(index, 'hidden');
				const next = session.workbook.sheets.findIndex((s) => s.state === 'visible');
				if (next >= 0) ctx.setActiveSheet(next);
			},
		}),
		sheetCommand({
			id: 'sheet.unhide',
			label: 'Unhide Sheet...',
			icon: icon('sheet'),
			enabled: (ctx) => !!ctx.workbook()?.sheets.some((s) => s.state === 'hidden'),
			run: (ctx) => void ctx.dialogs.open('unhide-sheet'),
		}),
		editing({
			id: 'format.row-height',
			label: 'Row Height...',
			icon: icon('rowHeight'),
			lock: 'formatRows',
			run: (ctx) => void ctx.dialogs.open('row-height'),
		}),
		editing({
			id: 'format.autofit-rows',
			label: 'AutoFit Row Height',
			icon: icon('rowHeight'),
			lock: 'formatRows',
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setRowHeight(t.sheet, rowsOf(t.ranges), 'auto');
			},
		}),
		editing({
			id: 'format.column-width',
			label: 'Column Width...',
			icon: icon('columnWidth'),
			lock: 'formatColumns',
			run: (ctx) => void ctx.dialogs.open('column-width'),
		}),
		editing({
			id: 'format.autofit-columns',
			label: 'AutoFit Column Width',
			icon: icon('columnWidth'),
			lock: 'formatColumns',
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setColumnWidth(t.sheet, colsOf(t.ranges), 'auto', measure(ctx));
			},
		}),
		editing({
			id: 'format.default-width',
			label: 'Default Width...',
			icon: icon('columnWidth'),
			lock: 'formatColumns',
			run: (ctx) => void ctx.dialogs.open('default-width'),
		}),
		...(
			[
				['format.hide-rows', 'Hide Rows', 'row', true],
				['format.unhide-rows', 'Unhide Rows', 'row', false],
				['format.hide-columns', 'Hide Columns', 'col', true],
				['format.unhide-columns', 'Unhide Columns', 'col', false],
			] as const
		).map(([id, label, axis, hidden]) =>
			editing({
				id,
				label,
				icon: icon('hide'),
				lock: axis === 'row' ? 'formatRows' : 'formatColumns',
				run: (ctx) => {
					const t = target(ctx);
					if (!t) return;
					const all = axis === 'row' ? rowsOf(t.ranges) : colsOf(t.ranges);
					// Unhiding a single cell's row/column also reveals hidden neighbours like Excel does
					// when the selection spans them; a whole-sheet span is clipped to stored info.
					t.session.setHidden(
						t.sheet,
						axis,
						all.length > 20000 ? all.slice(0, 20000) : all,
						hidden,
					);
				},
			}),
		),
		editing({
			id: 'format.lock-cell',
			label: 'Lock Cell',
			icon: icon('protectSheet'),
			lock: false,
			checked: (ctx) => activeStyle(ctx)?.protection?.locked !== false,
			run: (ctx) => {
				const locked = activeStyle(ctx)?.protection?.locked !== false;
				style(ctx, { protection: { locked: !locked } });
			},
		}),
		editing({
			id: 'format.cells',
			label: 'Format Cells...',
			icon: icon('formatCells'),
			shortcut: 'Ctrl+1',
			lock: 'formatCells',
			run: (ctx, arg) => void ctx.dialogs.open('format-cells', arg),
		}),
	];
}
