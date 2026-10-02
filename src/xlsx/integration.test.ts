// Cross-module checks over every fixture: load, view, edit, save, reload.
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { formatAddress } from './address.js';
import { forEachCell } from './cells.js';
import { createEditSession } from './edit/session.js';
import { createCalcEngine } from './formula/engine.js';
import { isSpilledCell } from './formula/spill.js';
import { cellView } from './layout/cell-view.js';
import { createConditionalFormatEvaluator } from './layout/cf-evaluator.js';
import { chartView, renderChartSvg } from './layout/index.js';
import { createRefEvaluator } from './layout/ref-evaluator.js';
import { loadWorkbook } from './load/index.js';
import type { CellValue, Workbook } from './model.js';
import { loadXlsx } from './read/index.js';
import { saveXlsx } from './write/index.js';

const DIR = path.join(import.meta.dirname, '__fixtures__');
const XLSX = readdirSync(DIR).filter((f) => f.endsWith('.xlsx'));
const XLS = readdirSync(path.join(DIR, 'xls')).filter(
	(f) => f.endsWith('.xls') && !f.includes('encrypted'),
);
const bytes = (...parts: string[]) => new Uint8Array(readFileSync(path.join(DIR, ...parts)));

/** Every stored value and formula, keyed `Sheet!A1`, spilled results excluded. */
function snapshot(workbook: Workbook): Map<string, { value: CellValue; formula?: string }> {
	const out = new Map<string, { value: CellValue; formula?: string }>();
	for (const sheet of workbook.sheets)
		forEachCell(sheet, (cell, row, col) => {
			if (isSpilledCell(cell) || (cell.value === null && cell.formula === undefined)) return;
			const entry: { value: CellValue; formula?: string } = { value: cell.value };
			if (cell.formula !== undefined) entry.formula = cell.formula;
			out.set(`${sheet.name}!${formatAddress({ row, col })}`, entry);
		});
	return out;
}

function viewEverything(workbook: Workbook): number {
	const calc = createCalcEngine(workbook);
	calc.recalculateAll();
	let views = 0;
	workbook.sheets.forEach((sheet, s) => {
		const cf = createConditionalFormatEvaluator(workbook, s, (formula, at) =>
			calc.evaluate(formula, at),
		);
		forEachCell(sheet, (_cell, row, col) => {
			const view = cellView(workbook, s, row, col, cf);
			expect(typeof view.text).toBe('string');
			views++;
		});
		const evaluateRef = createRefEvaluator(workbook, calc, { sheet: s });
		for (const drawing of sheet.drawings) {
			if (drawing.kind !== 'chart') continue;
			const model = chartView(workbook, s, drawing, evaluateRef);
			expect(renderChartSvg(model, 400, 300)).toContain('<svg');
		}
	});
	return views;
}

/** A sequence of edits touching structure, sheets, formulas and clipboard. */
function edit(workbook: Workbook): void {
	const s = createEditSession(workbook);
	const first = workbook.sheets.findIndex((sheet) => sheet.state === 'visible');
	s.insertRows(first, 1, 2);
	s.insertColumns(first, 0, 1);
	s.setCellInput(first, 0, 0, '=SEQUENCE(2,2)');
	s.setCellInput(first, 3, 0, '=SUM(A1:B2)');
	s.deleteRows(first, 5, 1);
	s.paste(
		first,
		{ row: 40, col: 10 },
		s.cut(first, { start: { row: 3, col: 0 }, end: { row: 3, col: 0 } }),
	);
	s.renameSheet(first, `${workbook.sheets[first]!.name.slice(0, 20)} R`);
	const copy = s.duplicateSheet(first);
	s.addSheet();
	s.setCellInput(workbook.sheets.length - 1, 0, 0, `='${workbook.sheets[copy]!.name}'!A1*2`);
	if (workbook.sheets.length > 3) s.deleteSheet(workbook.sheets.length - 2);
	s.undo();
	s.redo();
}

async function roundTrip(workbook: Workbook): Promise<Workbook> {
	const once = await loadXlsx(await saveXlsx(workbook));
	createCalcEngine(once).recalculateAll();
	return once;
}

describe.each(XLSX)('fixture %s', (name) => {
	it('views every cell, conditional format and chart without throwing', async () => {
		expect(viewEverything(await loadXlsx(bytes(name)))).toBeGreaterThan(0);
	});
	it('edits, saves and reloads to the same cells (and is stable on a second save)', async () => {
		const workbook = await loadXlsx(bytes(name));
		edit(workbook);
		const before = snapshot(workbook);
		const once = await roundTrip(workbook);
		const after = snapshot(once);
		expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
		for (const [key, entry] of before) expect(after.get(key), key).toEqual(entry);
		expect(once.sheets.map((s) => s.name)).toEqual(workbook.sheets.map((s) => s.name));
		const twice = await roundTrip(once);
		expect(snapshot(twice)).toEqual(after);
		viewEverything(twice);
	});
});

describe.each(XLS)('legacy fixture %s', (name) => {
	it('loads, views, edits and saves as xlsx', async () => {
		const workbook = await loadWorkbook(bytes('xls', name), { fileName: name });
		viewEverything(workbook);
		edit(workbook);
		const before = snapshot(workbook);
		const once = await roundTrip(workbook);
		for (const [key, entry] of before) expect(snapshot(once).get(key), key).toEqual(entry);
	});
});
