// Row and column dimension commands record only the dimensions they change for undo, and undo and
// redo still restore the whole sheet exactly as a whole-sheet snapshot would.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { Workbook } from '../model';
import { internStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { setColumnWidth, setHidden, setRowHeight } from './dimensions';
import { groupColumns, groupRows, setDefaultColumnWidth, ungroupRows } from './outline';
import { autoGrowRows } from './row-autofit';
import { createEditSession } from './session';
import { testContext } from './test-context';
import type { WorkbookChange } from './types';

const CELL_SENTINEL = 'cell held by the sheet';
/** A height no command under test writes, stored on a row none of them touches. */
const OTHER_ROW_HEIGHT = 77.25;

function workbook(): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const base = defaultCellStyle();
	const big = internStyle(wb, { ...base, font: { ...base.font, size: 20 } });
	for (let row = 0; row < 40; row++) {
		putCell(sheet, row, 0, { value: `${CELL_SENTINEL} ${row}` });
		putCell(sheet, row, 1, { value: row, ...(row % 7 === 3 ? { styleId: big } : {}) });
	}
	sheet.rowInfo.set(35, { height: OTHER_ROW_HEIGHT, customHeight: true });
	sheet.rowInfo.set(6, { height: 0, hidden: true });
	sheet.columns.push(
		{ min: 0, max: 2, width: 12, customWidth: true },
		{ min: 5, max: 9, hidden: true },
	);
	sheet.autoFilter = {
		range: { start: { row: 0, col: 0 }, end: { row: 20, col: 1 } },
		columns: [{ offset: 1, values: ['3'] }],
	};
	return wb;
}

function childrenOf(value: object): unknown[] {
	if (value instanceof Map) return [...value.keys(), ...value.values()];
	if (value instanceof Set) return [...value];
	return Object.values(value);
}

/** True when `value` holds `needle` anywhere inside (objects, arrays, maps, sets). */
function holds(value: unknown, needle: unknown, seen = new Set<unknown>()): boolean {
	if (value === needle) return true;
	if (typeof value === 'string' && typeof needle === 'string') return value.includes(needle);
	if (!value || typeof value !== 'object' || seen.has(value)) return false;
	seen.add(value);
	return childrenOf(value).some((child) => holds(child, needle, seen));
}

const DIMENSION_COMMANDS: [string, (ctx: ReturnType<typeof testContext>) => unknown][] = [
	['auto-grow', (ctx) => autoGrowRows(ctx, 0, [3, 10, 17])],
	['setRowHeight', (ctx) => setRowHeight(ctx, 0, [2, 4], 30)],
	['setRowHeight auto', (ctx) => setRowHeight(ctx, 0, [3, 6], 'auto')],
	['setColumnWidth', (ctx) => setColumnWidth(ctx, 0, [1, 3], 20)],
	['setColumnWidth auto', (ctx) => setColumnWidth(ctx, 0, [0, 1], 'auto')],
	['hide rows', (ctx) => setHidden(ctx, 0, 'row', [4, 8], true)],
	['unhide rows', (ctx) => setHidden(ctx, 0, 'row', [6], false)],
	['hide columns', (ctx) => setHidden(ctx, 0, 'col', [1], true)],
	['unhide columns', (ctx) => setHidden(ctx, 0, 'col', [7], false)],
	['group rows', (ctx) => groupRows(ctx, 0, 2, 5)],
	['group columns', (ctx) => groupColumns(ctx, 0, 1, 3)],
	['standard width', (ctx) => setDefaultColumnWidth(ctx, 0, 11)],
];

describe('dimension undo scope', () => {
	it.each(DIMENSION_COMMANDS)('%s records neither cells nor untouched rows', (_, command) => {
		const ctx = testContext(workbook());
		command(ctx);
		expect(ctx.steps).toHaveLength(1);
		const entries = ctx.steps[0]?.entries;
		expect(holds(entries, CELL_SENTINEL)).toBe(false);
		expect(holds(entries, OTHER_ROW_HEIGHT)).toBe(false);
	});
});

describe('dimension undo and redo', () => {
	it('restore the whole workbook state before and after every step', () => {
		const wb = workbook();
		const session = createEditSession(wb, {
			recalc: false,
			measureText: (text) => text.length * 7,
		});
		const changes: WorkbookChange[] = [];
		session.onChange((change) => changes.push(change));
		const one = (row: number, col: number) => ({ start: { row, col }, end: { row, col } });
		// Each step and whether it changes dimensions (a structural change for collaboration).
		const steps: [() => void, boolean][] = [
			// Edits that grow rows join their auto-grow in one step.
			[() => session.setCellValue(0, 10, 1, 'larger text'), true],
			[() => session.applyStyle(0, [one(14, 1)], { font: { size: 26 } }), true],
			[() => session.applyStyle(0, [one(13, 0)], { alignment: { wrapText: true } }), true],
			[() => session.setCellValue(0, 15, 1, 'no growth'), false],
			[() => session.setRowHeight(0, [2, 4, 35], 30), true],
			[() => session.setRowHeight(0, [2, 3], 'auto'), true],
			[() => session.setColumnWidth(0, [1, 3], 20), true],
			[() => session.setColumnWidth(0, [0], 'auto'), true],
			[() => session.setHidden(0, 'row', [5, 9], true), true],
			[() => session.setHidden(0, 'row', [6, 9], false), true],
			[() => session.setHidden(0, 'col', [2, 6], true), true],
			[() => session.setHidden(0, 'col', [7], false), true],
			[() => session.groupRows(0, 2, 8), true],
			[() => session.ungroupRows(0, 2, 4), true],
			[() => session.groupColumns(0, 0, 3), true],
			[() => session.ungroupColumns(0, 3, 3), true],
			[() => session.setDefaultColumnWidth(0, 11), true],
			[() => session.setDefaultColumnWidth(0, undefined), true],
			[
				() =>
					session.batch('Several', () => {
						session.setRowHeight(0, [20], 44);
						session.setCellValue(0, 21, 1, 'x '.repeat(40));
						session.setColumnWidth(0, [4], 3);
						session.setHidden(0, 'row', [22], true);
					}),
				true,
			],
		];
		const states = [structuredClone(wb.sheets)];
		for (const [step] of steps) {
			step();
			states.push(structuredClone(wb.sheets));
		}
		expect(changes.map((change) => change.structural)).toEqual(steps.map(([, dims]) => dims));
		for (let i = steps.length - 1; i >= 0; i--) {
			expect(session.undo()).toBe(true);
			expect(wb.sheets).toStrictEqual(states[i]);
		}
		for (let i = 1; i <= steps.length; i++) {
			expect(session.redo()).toBe(true);
			expect(wb.sheets).toStrictEqual(states[i]);
		}
	});

	it('undoes a dimension step in the same sheet objects', () => {
		const wb = workbook();
		const session = createEditSession(wb, { recalc: false });
		const sheet = wb.sheets[0];
		const rowInfo = sheet?.rowInfo;
		session.setRowHeight(0, [2], 30);
		session.undo();
		expect(wb.sheets[0]).toBe(sheet);
		expect(wb.sheets[0]?.rowInfo).toBe(rowInfo);
		expect(rowInfo?.has(2)).toBe(false);
	});

	it('keeps an undo that ungroups rows exact', () => {
		const ctx = testContext(workbook());
		groupRows(ctx, 0, 1, 3);
		const grouped = structuredClone(ctx.workbook.sheets);
		ungroupRows(ctx, 0, 1, 3);
		ctx.undo();
		expect(ctx.workbook.sheets).toStrictEqual(grouped);
	});
});
