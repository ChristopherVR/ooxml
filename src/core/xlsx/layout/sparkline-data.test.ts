import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { putCell } from '../cells';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { cellView } from './cell-view';
import {
	hasSparklineAt,
	readSparklineValues,
	resolveSparklineRange,
	sparklineViewAt,
} from './sparkline-data';

const round = (n: number) => Math.round(n * 1e6) / 1e6;

const fixture = () =>
	loadXlsx(
		new Uint8Array(
			readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-sparklines.xlsx')),
		),
	);

describe('sparklines in the cell view', () => {
	it('lays out each Excel sparkline from its data range', async () => {
		const wb = await fixture();
		const sheet = wb.sheets[0]!;
		expect(hasSparklineAt(sheet, 0, 5)).toBe(true);
		expect(hasSparklineAt(sheet, 0, 4)).toBe(false);
		const line = sparklineViewAt(wb, 0, 0, 5)!;
		expect(line.type).toBe('line');
		expect(line.lineWeightPt).toBe(1.5);
		expect(line.lines[0]!.map((p) => p.y)).toEqual([4 / 7, 2 / 7, 1, 0, 1 / 7]);
		expect(line.markers).toHaveLength(5);
		// The column sparkline leaves a gap for the empty C2 and colours first and last.
		const column = sparklineViewAt(wb, 0, 1, 5)!;
		expect(column.columns.map((c) => c.color)).toEqual([
			'#D00000',
			'#376092',
			'#376092',
			'#D00000',
		]);
		const stacked = sparklineViewAt(wb, 0, 2, 5)!;
		expect(stacked.columns.map((c) => c.y)).toEqual([0, 0.5, 0, 0, 0.5]);
		const manual = sparklineViewAt(wb, 0, 3, 5)!;
		expect(manual.lines[0]!.map((p) => round(p.y))).toEqual([0.8, 0.2, 0.4, 0.7, 0.3]);
		expect(manual.axis).toEqual({ y: 1, color: '#000000' });
		expect(cellView(wb, 0, 3, 5).sparkline).toEqual(manual);
		expect(cellView(wb, 0, 3, 4).sparkline).toBeUndefined();
	});

	it('reflects a changed data cell on the next view', async () => {
		const wb = await fixture();
		putCell(wb.sheets[0]!, 0, 3, { value: 50 });
		expect(sparklineViewAt(wb, 0, 0, 5)!.lines[0]![3]).toEqual({ x: 0.75, y: 0 });
	});

	it('resolves qualified, quoted and unqualified ranges', () => {
		const wb = createWorkbook({ sheets: ['Data', "Bob's sheet"] });
		const [data, other] = wb.sheets as [(typeof wb.sheets)[0], (typeof wb.sheets)[0]];
		expect(resolveSparklineRange(wb, other, 'Data!$A$1:$C$1')).toEqual({
			sheet: data,
			range: parseRange('A1:C1'),
		});
		expect(resolveSparklineRange(wb, data, "'Bob''s sheet'!B2:B4")?.sheet).toBe(other);
		expect(resolveSparklineRange(wb, data, 'A1:A3')?.sheet).toBe(data);
		expect(resolveSparklineRange(wb, data, 'Missing!A1:A3')).toBeUndefined();
	});

	it('reads numbers along the range, skipping hidden rows unless asked', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 1 });
		putCell(sheet, 1, 0, { value: 'text' });
		putCell(sheet, 2, 0, { value: 3 });
		sheet.rowInfo.set(2, { hidden: true });
		const range = parseRange('A1:A4')!;
		expect(readSparklineValues(sheet, range, false)).toEqual([1, null, null]);
		expect(readSparklineValues(sheet, range, true)).toEqual([1, null, 3, null]);
		// A whole-column range stops at the last stored row.
		expect(readSparklineValues(sheet, parseRange('A:A')!, true)).toEqual([1, null, 3]);
	});
});
