// Workbook Statistics (Review > Proofing): read-only counts for the current sheet and workbook.
import { type Worksheet, formatAddress, usedRange } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { el, fieldset } from './fields.js';
import { showDialog } from './frame.js';

export interface SheetStats {
	cells: number;
	formulas: number;
	tables: number;
	charts: number;
	images: number;
	comments: number;
}

export function sheetStats(ws: Worksheet): SheetStats {
	let cells = 0;
	let formulas = 0;
	for (const row of ws.rows.values())
		for (const cell of row.values()) {
			if (cell.formula !== undefined) formulas++;
			if (cell.formula !== undefined || (cell.value !== null && cell.value !== '')) cells++;
		}
	return {
		cells,
		formulas,
		tables: ws.tables.length,
		charts: ws.drawings.filter((d) => d.kind === 'chart').length,
		images: ws.drawings.filter((d) => d.kind === 'image').length,
		comments: ws.comments.length,
	};
}

function statsTable(
	ctx: EditorContext,
	rows: ReadonlyArray<readonly [string, string | number]>,
): HTMLTableElement {
	const table = ctx.host.ownerDocument.createElement('table');
	table.className = 'xve-stats';
	const body = table.createTBody();
	for (const [label, value] of rows) {
		const tr = body.insertRow();
		const th = ctx.host.ownerDocument.createElement('th');
		th.scope = 'row';
		th.textContent = ctx.t(label);
		tr.append(th);
		tr.insertCell().textContent = String(value);
	}
	return table;
}

export function openStatistics(ctx: EditorContext): Promise<undefined> {
	const workbook = ctx.workbook();
	if (!workbook) return Promise.resolve(undefined);
	const ws = workbook.sheets[ctx.activeSheet()];
	const current = ws ? sheetStats(ws) : undefined;
	const used = ws ? usedRange(ws) : undefined;
	const all = workbook.sheets.map(sheetStats);
	const sum = (key: keyof SheetStats): number => all.reduce((n, s) => n + s[key], 0);
	const body: HTMLElement[] = [];
	if (current && ws) {
		const title = el(ctx, 'p', 'xve-note');
		title.textContent = ctx.t('Current sheet: {name}', { name: ws.name });
		body.push(
			fieldset(
				ctx,
				'Current sheet',
				title,
				statsTable(ctx, [
					['End of sheet', used ? formatAddress(used.end) : 'A1'],
					['Cells with data', current.cells],
					['Tables', current.tables],
					['Formulas', current.formulas],
					['Charts', current.charts],
					['Images', current.images],
					['Notes', current.comments],
				]),
			),
		);
	}
	body.push(
		fieldset(
			ctx,
			'Workbook',
			statsTable(ctx, [
				['Sheets', workbook.sheets.length],
				['Cells with data', sum('cells')],
				['Tables', sum('tables')],
				['Formulas', sum('formulas')],
				['Charts', sum('charts')],
				['Images', sum('images')],
				['Notes', sum('comments')],
			]),
		),
	);
	return showDialog<undefined>(ctx, {
		name: 'workbook-statistics',
		heading: 'Workbook Statistics',
		okLabel: null,
		body,
	});
}
