import { describe, expect, it } from 'vitest';
import { deleteCell, getCell, putCell } from '../cells.js';
import { CIRCULAR_REFERENCE_WARNING } from './engine.js';
import { isSpilledCell, spillAnchorOf } from './spill.js';
import { book, E, engine, get, grid, locate, set } from './test-helpers.js';

const at = (wb: Parameters<typeof locate>[0], key: string) => locate(wb, key);

describe('recalculateAll', () => {
	it('evaluates formulas in dependency order regardless of position', () => {
		const wb = book({ A1: '=A2*2', A2: '=A3+1', A3: 5, B1: '=SUM(A1:A3)' });
		engine(wb).recalculateAll();
		expect(grid(wb, 'A1:B3')).toEqual([
			[12, 23],
			[6, null],
			[5, null],
		]);
	});

	it('handles long chains without deep recursion', () => {
		const wb = book({ A1: 1 });
		for (let r = 2; r <= 5000; r++) set(wb, `A${r}`, `=A${r - 1}+1`);
		engine(wb).recalculateAll();
		expect(get(wb, 'A5000')).toBe(5000);
	});

	it('follows dynamic references (INDIRECT, OFFSET) to formulas computed later', () => {
		const wb = book({
			A1: '=INDIRECT("C1")*2',
			B1: '=SUM(OFFSET(C1,0,0,2,1))',
			C1: '=C2+1',
			C2: 4,
		});
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(10);
		expect(get(wb, 'B1')).toBe(9);
	});

	it('keeps cached values of formulas using unsupported functions when loading', () => {
		const wb = book({ A1: 1 });
		const sheet = wb.sheets[0];
		putCell(sheet!, 0, 1, { value: 42, formula: 'CUBEVALUE("x",A1)' });
		putCell(sheet!, 0, 2, { value: 7, formula: 'B1+1' });
		engine(wb).recalculateAll();
		expect(get(wb, 'B1')).toBe(42);
		expect(get(wb, 'C1')).toBe(43);
	});

	it('shows #NAME? for an unsupported function typed after loading', () => {
		const wb = book({ A1: 1 });
		const calc = engine(wb);
		calc.recalculateAll();
		set(wb, 'B1', '=CUBEVALUE("x")');
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(get(wb, 'B1')).toEqual(E.NAME);
	});
});

describe('recalculateFrom', () => {
	it('recalculates only dependents of the changed cells', () => {
		const wb = book({ A1: 1, B1: '=A1*10', C1: '=B1+1', D1: '=5', E1: 2 });
		const calc = engine(wb);
		calc.recalculateAll();
		const d1 = getCell(wb.sheets[0]!, 0, 3);
		if (d1) d1.value = 'untouched';
		set(wb, 'A1', 3);
		calc.recalculateFrom([at(wb, 'A1')]);
		expect(get(wb, 'C1')).toBe(31);
		expect(get(wb, 'D1')).toBe('untouched');
	});

	it('propagates through ranges, other sheets and defined names', () => {
		const wb = book({ A1: 1, A2: 2, 'Sheet2!A1': '=SUM(Sheet1!A:A)', B1: '=Total*2' });
		wb.definedNames.push({ name: 'Total', formula: 'Sheet2!$A$1' });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'B1')).toBe(6);
		set(wb, 'A3', 10);
		calc.recalculateFrom([at(wb, 'A3')]);
		expect(get(wb, 'Sheet2!A1')).toBe(13);
		expect(get(wb, 'B1')).toBe(26);
	});

	it('picks up new, changed and removed formulas', () => {
		const wb = book({ A1: 2, B1: '=A1+1' });
		const calc = engine(wb);
		calc.recalculateAll();
		set(wb, 'C1', '=B1*2');
		calc.recalculateFrom([at(wb, 'C1')]);
		expect(get(wb, 'C1')).toBe(6);
		set(wb, 'B1', '=A1+100');
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(get(wb, 'C1')).toBe(204);
		set(wb, 'B1', 1);
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(get(wb, 'C1')).toBe(2);
		set(wb, 'A1', 50);
		calc.recalculateFrom([at(wb, 'A1')]);
		expect(get(wb, 'B1')).toBe(1);
	});

	it('updates table and structured-reference dependents', () => {
		const wb = book({ A1: 'Qty', A2: 1, A3: 2, C1: '=SUM(T[Qty])' });
		wb.sheets[0]?.tables.push({
			id: 1,
			name: 'T',
			displayName: 'T',
			range: { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } },
			headerRow: true,
			totalsRow: false,
			columns: [{ name: 'Qty' }],
		});
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'C1')).toBe(3);
		set(wb, 'A3', 10);
		calc.recalculateFrom([at(wb, 'A3')]);
		expect(get(wb, 'C1')).toBe(11);
	});
});

describe('circular references', () => {
	it('sets cycle members to 0 and adds one warning', () => {
		const wb = book({ A1: '=B1+1', B1: '=A1+1', C1: '=A1+5', D1: '=D1' });
		const calc = engine(wb);
		calc.recalculateAll();
		calc.recalculateAll();
		expect(get(wb, 'A1')).toBe(0);
		expect(get(wb, 'B1')).toBe(0);
		expect(get(wb, 'C1')).toBe(5);
		expect(get(wb, 'D1')).toBe(0);
		expect(wb.warnings.filter((w) => w === CIRCULAR_REFERENCE_WARNING)).toHaveLength(1);
		expect(calc.circularCells()).toHaveLength(3);
	});

	it('treats a range that contains its own cell as circular', () => {
		const wb = book({ A1: 1, A2: 2, A3: '=SUM(A1:A3)' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A3')).toBe(0);
	});

	it('detects cycles only reachable through INDIRECT', () => {
		const wb = book({ A1: '=INDIRECT("B1")+1', B1: '=A1+1' });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(calc.circularCells().length).toBeGreaterThan(0);
		expect(wb.warnings).toContain(CIRCULAR_REFERENCE_WARNING);
	});
});

describe('dynamic arrays', () => {
	it('spills into empty neighbours and marks the spilled cells', () => {
		const wb = book({ A1: '=SEQUENCE(3,2)' });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(grid(wb, 'A1:B3')).toEqual([
			[1, 2],
			[3, 4],
			[5, 6],
		]);
		const b3 = getCell(wb.sheets[0]!, 2, 1);
		expect(isSpilledCell(b3)).toBe(true);
		expect(spillAnchorOf(b3)).toEqual({ row: 0, col: 0 });
		expect(isSpilledCell(getCell(wb.sheets[0]!, 0, 0))).toBe(false);
		expect(calc.spillRange(0, 0, 0)).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		});
	});

	it('shows #SPILL! when blocked and spills again once the blocker is cleared', () => {
		const wb = book({ A1: '=SEQUENCE(3)', A3: 'x' });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(get(wb, 'A2')).toBe(null);
		deleteCell(wb.sheets[0]!, 2, 0);
		calc.recalculateFrom([at(wb, 'A3')]);
		expect(grid(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
	});

	it('becomes #SPILL! when the user types into the spill range', () => {
		const wb = book({ A1: '=SEQUENCE(3)' });
		const calc = engine(wb);
		calc.recalculateAll();
		set(wb, 'A2', 'mine');
		calc.recalculateFrom([at(wb, 'A2')]);
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(get(wb, 'A2')).toBe('mine');
		expect(get(wb, 'A3')).toBe(null);
	});

	it('shrinks and grows the spill range with its inputs', () => {
		const wb = book({ B1: 3, A1: '=SEQUENCE(B1)' });
		const calc = engine(wb);
		calc.recalculateAll();
		set(wb, 'B1', 1);
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(grid(wb, 'A1:A3')).toEqual([[1], [null], [null]]);
		set(wb, 'B1', 4);
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(grid(wb, 'A1:A4')).toEqual([[1], [2], [3], [4]]);
	});

	it('lets other formulas read spilled values and spill references', () => {
		const wb = book({
			D1: '=SUM(A1#)',
			E1: '=A3*10',
			A1: '=SEQUENCE(B1)',
			B1: 3,
			F1: '=ROWS(A1#)',
			G1: '=B5#',
		});
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'D1')).toBe(6);
		expect(get(wb, 'E1')).toBe(30);
		expect(get(wb, 'F1')).toBe(3);
		expect(get(wb, 'G1')).toEqual(E.REF);
		set(wb, 'B1', 5);
		calc.recalculateFrom([at(wb, 'B1')]);
		expect(get(wb, 'D1')).toBe(15);
		expect(get(wb, 'F1')).toBe(5);
	});

	it('blocks when two spills would overlap', () => {
		const wb = book({ A1: '=SEQUENCE(3)', B2: 'x', A3: '=1' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(get(wb, 'A3')).toBe(1);
	});

	it('blocks on merged cells and at the grid edge', () => {
		const wb = book({ A1: '=SEQUENCE(3)', A1048576: '=SEQUENCE(2)' });
		wb.sheets[0]?.merges.push({ start: { row: 1, col: 0 }, end: { row: 1, col: 1 } });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(get(wb, 'A1048576')).toEqual(E.SPILL);
	});

	it('fills a CSE array formula range exactly', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, {
			value: null,
			formula: '{1;2}*10',
			arrayRange: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
		});
		engine(wb).recalculateAll();
		expect(grid(wb, 'A1:B3')).toEqual([
			[10, 10],
			[20, 20],
			[E.NA, E.NA],
		]);
		expect(isSpilledCell(getCell(sheet, 1, 1))).toBe(false);
	});

	it('clears spilled values on invalidate and recomputes them afterwards', () => {
		const wb = book({ A1: '=SEQUENCE(2)' });
		const calc = engine(wb);
		calc.recalculateAll();
		calc.invalidate();
		expect(get(wb, 'A2')).toBe(null);
		calc.recalculateFrom([]);
		expect(get(wb, 'A2')).toBe(2);
	});
});

describe('ad hoc evaluation', () => {
	it('evaluates a formula at a position without storing it', () => {
		const wb = book({ A1: 1, A2: 2, A3: 3 });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(calc.evaluate('SUM(A1:A3)', { sheet: 0, row: 0, col: 5 })).toBe(6);
		expect(calc.evaluate('A1:A3', { sheet: 0, row: 1, col: 5 })).toBe(2);
		expect(calc.evaluate('A1>0', { sheet: 0, row: 0, col: 0 })).toBe(true);
		expect(calc.evaluate('1+', { sheet: 0, row: 0, col: 0 })).toEqual(E.NAME);
		expect(calc.evaluateArray('A1:A3*2', { sheet: 0, row: 0, col: 0 })).toEqual([[2], [4], [6]]);
		expect(calc.evaluateArray('$A$1:$A$2', { sheet: 0, row: 0, col: 0 })).toEqual([[1], [2]]);
	});
});

describe('performance', () => {
	it('recalculates 10k formula cells quickly, fully and incrementally', () => {
		const wb = book();
		for (let r = 1; r <= 5000; r++) {
			set(wb, `A${r}`, r);
			set(wb, `B${r}`, `=A${r}*2+IF(MOD(A${r},2)=0,1,0)`);
			set(wb, `C${r}`, r === 1 ? '=B1' : `=C${r - 1}+B${r}`);
		}
		set(wb, 'D1', '=SUM(B:B)');
		set(wb, 'D2', '=VLOOKUP(2500,A1:B5000,2,FALSE)');
		const calc = engine(wb);
		const t0 = performance.now();
		calc.recalculateAll();
		const full = performance.now() - t0;
		expect(get(wb, 'D2')).toBe(5001);
		expect(get(wb, 'C5000')).toBe(get(wb, 'D1'));
		set(wb, 'A1', 100);
		const t1 = performance.now();
		calc.recalculateFrom([at(wb, 'A1')]);
		const incremental = performance.now() - t1;
		expect(get(wb, 'B1')).toBe(201);
		expect(get(wb, 'C5000')).toBe(get(wb, 'D1'));
		expect(full).toBeLessThan(3000);
		expect(incremental).toBeLessThan(3000);
	});
});
