// The exact-lookup index must be invisible: every result equals what a linear scan returns, over
// mixed types, near-equal numbers, wildcards and reverse searches, and it reads no cell a scan
// would not have reached.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { createCalcEngine } from './engine';
import { findExact } from './functions/lookup-core';
import { ExactSearch } from './pass-cache';
import { book, E, engine, get, locate, set } from './test-helpers';
import type { Scalar } from './values';

/** A small seeded generator, so a failure reproduces from its trial number. */
function random(seed: number) {
	let s = seed >>> 0;
	const next = (): number => {
		s = (s + 0x6d2b79f5) >>> 0;
		let t = s;
		t = Math.imul(t ^ (t >>> 15), t | 1);
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
	return {
		int: (n: number) => Math.floor(next() * n),
		pick: <T>(items: readonly T[]) => items[Math.floor(next() * items.length)] as T,
	};
}

const NUMBERS = [
	0,
	-0,
	1,
	-1,
	1.5,
	0.1 + 0.2,
	0.3,
	1.0000000000000002,
	1.00000000000001,
	1.00000000000002,
	9.999999999999996,
	10,
	10.00000000000004,
	123456789012345.6,
	123456789012345.7,
	999999999999999,
	1e15,
	1e15 + 2,
	1e21,
	1e308,
	5e-324,
	1e-320,
];
const TEXTS = ['a', 'A', 'ab', 'AB', 'é', 'É', 'ß', 'SS', 'İ', 'i̇', '', ' ', '1', 'TRUE', 'x~y'];
const PATTERNS = ['a*', '*', '?', 'a?', '~*', 'x~~y', '*b', '?B'];
const POOL: Scalar[] = [...NUMBERS, ...TEXTS, ...PATTERNS, true, false, null, E.NA, E.VALUE];
/** Values only an API caller can store; no formula produces NaN. */
const RAW_POOL: Scalar[] = [...POOL, Number.NaN];

describe('ExactSearch', () => {
	it.each([false, true])('agrees with a scan and reads nothing more (reverse %s)', (reverse) => {
		for (let trial = 0; trial < 300; trial++) {
			const rng = random(trial);
			const values = Array.from({ length: 1 + rng.int(120) }, () => rng.pick(RAW_POOL));
			const scanned = new Set<number>();
			const read = new Set<number>();
			const search = new ExactSearch(values.length, reverse);
			for (let i = 0; i < 40; i++) {
				const lookup = rng.pick(RAW_POOL);
				const expected = findExact(
					{
						length: values.length,
						get: (at) => {
							scanned.add(at);
							return values[at] as Scalar;
						},
					},
					lookup,
					false,
					reverse,
				);
				const found = search.find(lookup, (at) => {
					read.add(at);
					return values[at] as Scalar;
				});
				expect(found, `trial ${trial}, lookup ${String(lookup)}`).toBe(expected);
			}
			for (const at of read) expect(scanned.has(at)).toBe(true);
		}
	});

	it('repeats a read that unwound, so no position is skipped', () => {
		const search = new ExactSearch(3, false);
		let unwind = true;
		const read = (at: number): Scalar => {
			if (at === 1 && unwind) {
				unwind = false;
				throw new Error('computed first');
			}
			return at * 10;
		};
		expect(() => search.find(10, read)).toThrow('computed first');
		expect(search.find(10, read)).toBe(1);
		expect(search.find(20, read)).toBe(2);
	});
});

const KEYS = 48;

/** Literal text for a lookup value stored in a helper cell. */
function store(wb: Workbook, key: string, value: Scalar): void {
	const { sheet, row, col } = locate(wb, key);
	const ws = wb.sheets[sheet];
	if (ws && value !== null) putCell(ws, row, col, { value: value as CellValue });
}

const FORMS: readonly ((v: string) => string)[] = [
	(v) => `VLOOKUP(${v},$A$1:$B$${KEYS},2,FALSE)`,
	(v) => `HLOOKUP(${v},Sheet2!$A$1:$AV$2,2,FALSE)`,
	(v) => `MATCH(${v},$A$1:$A$${KEYS},0)`,
	(v) => `MATCH(${v},A:A,0)`,
	(v) => `INDEX($B$1:$B$${KEYS},MATCH(${v},$A$1:$A$${KEYS},0))`,
	(v) => `XMATCH(${v},$A$1:$A$${KEYS})`,
	(v) => `XMATCH(${v},$A$1:$A$${KEYS},0,-1)`,
	(v) => `XMATCH(${v},$A$1:$A$60,0,-1)`,
	(v) => `XMATCH(${v},$A$1:$A$60,2)`,
	(v) => `XLOOKUP(${v},$A$1:$A$${KEYS},$B$1:$B$${KEYS},"none")`,
	(v) => `XLOOKUP(${v},$A$1:$A$${KEYS},$B$1:$B$${KEYS},"none",0,-1)`,
	(v) => `XLOOKUP(${v},$A$1:$A$${KEYS},$B$1:$B$${KEYS},"none",2)`,
	(v) => `XLOOKUP(${v},$A$1:$A$${KEYS},$B$1:$B$${KEYS},"none",2,-1)`,
	(v) => `XLOOKUP(${v},Sheet2!$A$1:$AV$1,Sheet2!$A$2:$AV$2,"none",0,-1)`,
];

describe('lookups in a recalculation', () => {
	/**
	 * Every lookup first runs in one pass, sharing the index. Recalculating one at a time then
	 * gives each a pass of its own, where its only search is a plain scan: the reference result.
	 */
	it.each([0, 1, 2, 3, 4, 5, 6, 7])('match a scan after each edit (seed %i)', (seed) => {
		const rng = random(1000 + seed);
		const wb = book({});
		const sheet2 = wb.sheets[1];
		if (!sheet2) throw new Error('no sheet');
		for (let i = 0; i < KEYS; i++) {
			const row = i + 1;
			const value = rng.pick(POOL);
			// Some keys are formulas, computed while the search reads them.
			if (i % 5 === 0) {
				store(wb, `K${row}`, value);
				set(wb, `A${row}`, `=IF(ISBLANK(K${row}),"",K${row})`);
			} else store(wb, `A${row}`, value);
			set(wb, `B${row}`, row * 2);
			const across = rng.pick(POOL);
			if (across !== null) putCell(sheet2, 0, i, { value: across as CellValue });
			putCell(sheet2, 1, i, { value: i });
		}
		const lookups: string[] = [];
		for (let i = 0; i < 120; i++) {
			const helper = `H${i + 1}`;
			store(wb, helper, rng.pick(POOL));
			const key = `D${i + 1}`;
			set(wb, key, `=${rng.pick(FORMS)(`$${helper[0]}$${i + 1}`)}`);
			lookups.push(key);
		}
		const e = engine(wb);
		const matchesScan = (label: string) => {
			const shared = lookups.map((key) => get(wb, key));
			lookups.forEach((key, i) => {
				e.recalculateFrom([locate(wb, key)]);
				expect(
					get(wb, key),
					`${label}, ${key}: ${wb.sheets[0]?.rows.get(i)?.get(3)?.formula}`,
				).toEqual(shared[i]);
			});
		};
		e.recalculateAll();
		matchesScan('full');
		for (let edit = 0; edit < 6; edit++) {
			const key = `A${1 + rng.int(KEYS)}`;
			if (wb.sheets[0]?.rows.get(locate(wb, key).row)?.get(0)?.formula) continue;
			const ws = wb.sheets[0];
			if (!ws) throw new Error('no sheet');
			const { row, col } = locate(wb, key);
			putCell(ws, row, col, { value: rng.pick(POOL) as CellValue });
			e.recalculateFrom([locate(wb, key)]);
			matchesScan(`edit ${edit}`);
		}
	});

	it('stops at the first match, so a later cell that reads the lookup is no cycle', () => {
		const cells: Parameters<typeof book>[0] = { A1: 7, A2: 8 };
		for (let row = 3; row <= 40; row++) cells[`A${row}`] = row === 30 ? '=C1' : row;
		const wb = book({
			...cells,
			C1: '=MATCH(7,INDIRECT("A1:A40"),0)',
			C2: '=MATCH(8,INDIRECT("A1:A40"),0)',
			C3: '=MATCH(7,INDIRECT("A1:A40"),0)',
		});
		const e = engine(wb);
		e.recalculateAll();
		expect([get(wb, 'C1'), get(wb, 'C2'), get(wb, 'C3'), get(wb, 'A30')]).toEqual([1, 2, 1, 1]);
		expect(e.circularCells()).toEqual([]);
	});

	it('sees a spill that arrives after earlier searches of the same cells', () => {
		const wb = book({
			A1: 1,
			E1: '=SEQUENCE(A1*40)',
			C1: '=MATCH(7,INDIRECT("E1:E40"),0)',
			C2: '=MATCH(9,INDIRECT("E1:E40"),0)',
			C3: '=MATCH(40,INDIRECT("E1:E40"),0)',
		});
		const e = engine(wb);
		e.recalculateAll();
		expect([get(wb, 'C1'), get(wb, 'C2'), get(wb, 'C3')]).toEqual([7, 9, 40]);
		set(wb, 'A1', 0.5);
		e.recalculateFrom([locate(wb, 'A1')]);
		expect([get(wb, 'C1'), get(wb, 'C2'), get(wb, 'C3')]).toEqual([7, 9, E.NA]);
	});

	it('reads fresh values on every step of an iterative calculation', () => {
		const cells: Parameters<typeof book>[0] = { A1: '=C1+1' };
		for (let row = 2; row <= 40; row++) cells[`A${row}`] = 1000 + row;
		const wb = book({
			...cells,
			C1: '=IFERROR(MATCH(A1,A1:A40,0),0)+IF(ISNUMBER(C2),0,0)',
			C2: '=MATCH(A1,A1:A40,0)',
			C3: '=SUM(A1:A100)',
			C4: '=SUM(A1:A100)',
		});
		wb.iterate = { count: 5, delta: 0 };
		engine(wb).recalculateAll();
		expect(get(wb, 'C2')).toBe(1);
		expect(get(wb, 'C3')).toBe(get(wb, 'C4'));
		expect(get(wb, 'C3')).toBe((get(wb, 'A1') as number) + 39_819);
	});
});

describe('what a recalculation must not reuse', () => {
	it('drops what a spilling formula read of its own old spill', () => {
		// E1 searches the cells it spills into while it runs, so it reads their old values.
		const own =
			'IFERROR(MATCH(-1,INDIRECT("E2:E60"),0),0)+IFERROR(MATCH(-2,INDIRECT("E2:E60"),0),0)';
		const wb = book({
			A1: 50,
			B1: 1,
			E1: `=SEQUENCE(A1,1,B1)+0*(${own})`,
			C1: '=IFERROR(MATCH(7,INDIRECT("E2:E60"),0)+0*E1,-1)',
		});
		const e = engine(wb);
		e.recalculateAll();
		expect(get(wb, 'C1')).toBe(6);
		set(wb, 'B1', 11);
		e.recalculateFrom([locate(wb, 'B1')]);
		expect(get(wb, 'E2')).toBe(12);
		expect(get(wb, 'C1')).toBe(-1);
	});

	it('lets a lookup computed during a search of the same cells scan on its own', () => {
		// C2 reaches A5 while extending the index, which computes C3 searching the same cells.
		const wb = book({
			C1: '=MATCH(1,INDIRECT("A1:A40"),0)',
			C2: '=MATCH(TRUE,INDIRECT("A1:A40"),0)',
			C3: '=MATCH(7,INDIRECT("A1:A40"),0)',
			A1: 1,
			A2: 2,
			A5: '=C3',
			A10: true,
			A40: 0,
		});
		engine(wb).recalculateAll();
		expect([get(wb, 'C1'), get(wb, 'C2'), get(wb, 'C3'), get(wb, 'A5')]).toEqual([1, 10, E.NA, 0]);
	});

	it('drops a total read before a spill arrived in the range', () => {
		// A1 reads E1:E80 row by row; E10 computes D3 on the way, which spills into E3:E7.
		const wb = book({
			A1: '=SUM(INDIRECT("E1:E80"))',
			B1: '=SUM(INDIRECT("E1:E80"))',
			E10: '=INDIRECT("D3")*0+1',
			D3: '=SEQUENCE(5,2)',
		});
		engine(wb).recalculateAll();
		expect(get(wb, 'B1')).toBe(31);
	});

	it('drops what was read of a circular formula before it was marked', () => {
		// A1 never reads A2 while it runs, so computing it on demand sees no loop; the loop
		// sets it to 0 afterwards, and C20 runs after that. A40 makes the range long enough to index.
		const wb = book({
			C1: '=MATCH(5,INDIRECT("A1:A40"),0)',
			C2: '=MATCH(5,INDIRECT("A1:A40"),0)',
			A10: '=IF(FALSE,A11,5)',
			A11: '=A10',
			A40: 1,
			C20: '=IFERROR(MATCH(5,INDIRECT("A1:A40"),0)+0*A11,-1)',
		});
		engine(wb).recalculateAll();
		expect(get(wb, 'A10')).toBe(0);
		expect(get(wb, 'C20')).toBe(-1);
	});
});

describe('aggregates shared within a recalculation', () => {
	const column = (extra: Parameters<typeof book>[0]) => {
		const cells: Parameters<typeof book>[0] = { ...extra };
		for (let row = 1; row <= 100; row++) cells[`A${row}`] ??= row;
		return book(cells);
	};

	it('computes the formulas inside the range before using the result', () => {
		const wb = column({ A50: '=B1*2', B1: 500, C1: '=SUM(A1:A100)', C2: '=SUM(A1:A100)+1' });
		const e = engine(wb);
		e.recalculateAll();
		expect([get(wb, 'C1'), get(wb, 'C2')]).toEqual([5050 + 950, 5051 + 950]);
		set(wb, 'B1', 0);
		e.recalculateFrom([locate(wb, 'B1')]);
		expect([get(wb, 'C1'), get(wb, 'C2')]).toEqual([5000, 5001]);
	});

	it('repeats errors and keeps functions and ranges apart', () => {
		const wb = column({
			A9: '=NA()',
			C1: '=SUM(A1:A100)',
			C2: '=SUM(A1:A100)',
			C3: '=COUNT(A1:A100)',
			C4: '=COUNTA(A1:A100)',
			C5: '=COUNT(A1:A99)',
			C6: '=MAX(A10:A100)',
			C7: '=MIN(A10:A100)',
		});
		engine(wb).recalculateAll();
		expect(['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7'].map((key) => get(wb, key))).toEqual([
			E.NA,
			E.NA,
			99,
			100,
			98,
			100,
			10,
		]);
	});

	it('does not reuse a total read before a spill changed the range', () => {
		const wb = book({
			A1: 2,
			E1: '=SEQUENCE(A1*40)',
			C1: '=SUM(INDIRECT("E1:E100"))',
			C2: '=SUM(INDIRECT("E1:E100"))',
		});
		const e = engine(wb);
		e.recalculateAll();
		expect([get(wb, 'C1'), get(wb, 'C2')]).toEqual([3240, 3240]);
		set(wb, 'A1', 1);
		e.recalculateFrom([locate(wb, 'A1')]);
		expect([get(wb, 'C1'), get(wb, 'C2')]).toEqual([820, 820]);
	});
});

describe('volatile formulas', () => {
	it('follow formulas added and removed between recalculations', () => {
		let draws = 0;
		const wb = book({ A1: 1, B1: '=RAND()', C1: '=A1' });
		const e = createCalcEngine(wb, {
			random: () => {
				draws++;
				return 0.5;
			},
		});
		e.recalculateAll();
		expect(draws).toBe(1);
		e.recalculateFrom([locate(wb, 'A1')]);
		expect(draws).toBe(2);
		set(wb, 'B2', '=RAND()');
		e.recalculateFrom([locate(wb, 'B2')]);
		expect(draws).toBe(4);
		set(wb, 'B1', 1);
		e.recalculateFrom([locate(wb, 'B1')]);
		expect(draws).toBe(5);
		e.recalculateFrom([locate(wb, 'A1')]);
		expect(draws).toBe(6);
	});
});
