// Contextual Table Design and Chart Design commands. A table is "selected" when the active cell is
// inside it; a chart when it is the selected drawing (`selection.drawing`, set by clicking it).
import {
	type ChartObject,
	type ChartPatch,
	type ChartType,
	type Table,
} from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, tableAt, target } from './util.js';

export const activeTable = (ctx: EditorContext): Table | undefined => {
	const t = target(ctx);
	return t ? tableAt(t.ws, t.active) : undefined;
};

/** The selected chart (clicked in the grid, or just inserted), as Excel's Chart Design follows. */
export function activeChart(ctx: EditorContext): { chart: ChartObject; index: number } | undefined {
	const t = target(ctx);
	const index = ctx.selection.get().drawing;
	if (!t || index === undefined) return undefined;
	const picked = t.ws.drawings[index];
	return picked?.kind === 'chart' ? { chart: picked, index } : undefined;
}

/** Edits the active chart (a chart read from a file keeps its part; the core patches it on save). */
export function editChart(ctx: EditorContext, patch: (chart: ChartObject) => ChartPatch): void {
	const t = target(ctx);
	const found = activeChart(ctx);
	if (t && found) t.session.updateChart(t.sheet, found.index, patch(found.chart));
}

type TableFlag = 'showRowStripes' | 'showColumnStripes' | 'showFirstColumn' | 'showLastColumn';

const TABLE_FLAGS: ReadonlyArray<readonly [TableFlag, string, string]> = [
	['showRowStripes', 'table.banded-rows', 'Banded Rows'],
	['showColumnStripes', 'table.banded-columns', 'Banded Columns'],
	['showFirstColumn', 'table.first-column', 'First Column'],
	['showLastColumn', 'table.last-column', 'Last Column'],
];

/** Header Row and Total Row: the core adds or clears the band's cells like Excel. */
function setTableBand(ctx: EditorContext, band: 'headerRow' | 'totalsRow', on: boolean): void {
	const t = target(ctx);
	const table = activeTable(ctx);
	if (!t || !table || table[band] === on) return;
	try {
		t.session.updateTable(t.sheet, table.name, { [band]: on });
	} catch (error) {
		ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
	}
}

export function contextualCommands(): Command[] {
	const tableCmd = (command: Omit<Command, 'editing'>): Command =>
		editing({
			...command,
			lock: false,
			enabled: (ctx) => !!activeTable(ctx) && (command.enabled?.(ctx) ?? true),
		});
	const chartCmd = (command: Omit<Command, 'editing'>): Command =>
		editing({
			...command,
			lock: 'objects',
			enabled: (ctx) => !!activeChart(ctx) && (command.enabled?.(ctx) ?? true),
		});
	return [
		tableCmd({
			id: 'table.name',
			label: 'Table Name',
			icon: icon('tableName'),
			value: (ctx) => activeTable(ctx)?.displayName ?? '',
			run: (ctx) => void ctx.dialogs.open('table-name'),
		}),
		tableCmd({
			id: 'table.resize',
			label: 'Resize Table',
			icon: icon('resizeTable'),
			run: (ctx) => void ctx.dialogs.open('table-resize'),
		}),
		tableCmd({
			id: 'table.convert-to-range',
			label: 'Convert to Range',
			icon: icon('convertRange'),
			run: async (ctx) => {
				const t = target(ctx);
				const table = activeTable(ctx);
				if (!t || !table) return;
				const ok = await ctx.dialogs.open<boolean>('confirm', {
					heading: 'Convert to Range',
					message: 'Do you want to convert the table to a normal range?',
					okLabel: 'Yes',
				});
				if (ok) t.session.convertTableToRange(t.sheet, table.name);
			},
		}),
		tableCmd({
			id: 'table.header-row',
			label: 'Header Row',
			checked: (ctx) => activeTable(ctx)?.headerRow !== false,
			run: (ctx) => setTableBand(ctx, 'headerRow', activeTable(ctx)?.headerRow === false),
		}),
		tableCmd({
			id: 'table.total-row',
			label: 'Total Row',
			shortcut: 'Ctrl+Shift+T',
			checked: (ctx) => !!activeTable(ctx)?.totalsRow,
			run: (ctx) => setTableBand(ctx, 'totalsRow', !activeTable(ctx)?.totalsRow),
		}),
		...TABLE_FLAGS.map(([flag, id, label]) =>
			tableCmd({
				id,
				label,
				checked: (ctx) => !!activeTable(ctx)?.[flag],
				run: (ctx) => {
					const t = target(ctx);
					const table = activeTable(ctx);
					if (t && table) t.session.updateTable(t.sheet, table.name, { [flag]: !table[flag] });
				},
			}),
		),
		tableCmd({
			id: 'table.style',
			label: 'Table Styles',
			icon: icon('formatTable'),
			run: (ctx, arg) => {
				const t = target(ctx);
				const table = activeTable(ctx);
				if (t && table && typeof arg === 'string')
					t.session.updateTable(t.sheet, table.name, { styleName: arg });
			},
		}),
		chartCmd({
			id: 'chart.type',
			label: 'Change Chart Type',
			icon: icon('chartType'),
			run: (ctx, arg) => {
				if (typeof arg === 'string') editChart(ctx, () => ({ chartType: arg as ChartType }));
				else void ctx.dialogs.open('insert-chart', { change: true });
			},
		}),
		chartCmd({
			id: 'chart.title',
			label: 'Chart Title',
			icon: icon('chartTitle'),
			run: (ctx) => void ctx.dialogs.open('chart-title'),
		}),
		chartCmd({
			id: 'chart.legend',
			label: 'Legend',
			icon: icon('legend'),
			checked: (ctx) => !!activeChart(ctx)?.chart.showLegend,
			run: (ctx) => editChart(ctx, (chart) => ({ showLegend: !chart.showLegend })),
		}),
		chartCmd({
			id: 'chart.delete',
			label: 'Delete Chart',
			icon: 'trash',
			run: (ctx) => {
				const t = target(ctx);
				const found = activeChart(ctx);
				if (!t || !found) return;
				ctx.selection.set({ drawing: undefined });
				t.session.deleteDrawing(t.sheet, found.index);
			},
		}),
	];
}
