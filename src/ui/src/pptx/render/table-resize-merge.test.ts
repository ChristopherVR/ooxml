import type { PptxTableCell, PptxTableRow } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	columnHandleSegments,
	columnIndexAtFraction,
	computeTableMergeCrossings,
	isColumnBoundaryMergedAt,
	isRowBoundaryMergedAt,
	openBoundaryRuns,
	rowHandleSegments,
	rowIndexAtOffset,
} from './table-resize-merge';

const cell = (extra: Partial<PptxTableCell> = {}): PptxTableCell => ({ text: '', ...extra });
const row = (cells: PptxTableCell[]): PptxTableRow => ({ cells }) as PptxTableRow;

/** 3x3: R1C1 spans two rows, R3C2 spans two columns. */
const merged = [
	row([cell({ rowSpan: 2 }), cell(), cell()]),
	row([cell({ vMerge: true }), cell(), cell()]),
	row([cell(), cell({ gridSpan: 2 }), cell({ hMerge: true })]),
];

describe('computeTableMergeCrossings', () => {
	it('marks the boundaries a merged cell spans', () => {
		const crossings = computeTableMergeCrossings(merged, 3);
		expect(crossings.rows).toStrictEqual([
			[true, false, false],
			[false, false, false],
		]);
		expect(crossings.columns).toStrictEqual([
			[false, false, false],
			[false, false, true],
		]);
	});

	it('crosses nothing in an unmerged table', () => {
		const plain = [row([cell(), cell()]), row([cell(), cell()])];
		const crossings = computeTableMergeCrossings(plain, 2);
		expect(crossings.rows).toStrictEqual([[false, false]]);
		expect(crossings.columns).toStrictEqual([[false, false]]);
	});

	it('clamps a span that runs past the grid', () => {
		const crossings = computeTableMergeCrossings([row([cell({ rowSpan: 4, gridSpan: 5 })])], 1);
		expect(crossings).toStrictEqual({ rows: [], columns: [] });
	});
});

describe('boundary segments', () => {
	it('splits a boundary into its real edges', () => {
		expect(openBoundaryRuns([true, false, false])).toStrictEqual([{ start: 1, end: 2 }]);
		expect(openBoundaryRuns([false, true, false])).toStrictEqual([
			{ start: 0, end: 0 },
			{ start: 2, end: 2 },
		]);
		expect(openBoundaryRuns([true, true])).toStrictEqual([]);
	});

	it('positions row handles by column width', () => {
		expect(rowHandleSegments([true, false, false], [0.25, 0.25, 0.5])).toStrictEqual([
			{ leftPct: 25, widthPct: 75 },
		]);
	});

	it('positions column handles by measured row boundaries', () => {
		expect(columnHandleSegments([false, false, true], [30, 60], 100)).toStrictEqual([
			{ top: 0, height: 60 },
		]);
		expect(columnHandleSegments([true, false, false], [30, 60], 100)).toStrictEqual([
			{ top: 30, height: 70 },
		]);
	});
});

describe('boundaries without merge information', () => {
	it('treats every grid line as an edge', () => {
		expect(rowHandleSegments([], [0.5, 0.5])).toStrictEqual([{ leftPct: 0, widthPct: 100 }]);
		expect(columnHandleSegments([], [40], 80)).toStrictEqual([{ top: 0, height: 80 }]);
	});
});

describe('boundary hit tests', () => {
	const crossings = computeTableMergeCrossings(merged, 3);
	const widths = [0.25, 0.25, 0.5];

	it('finds the column and row under a point', () => {
		expect(columnIndexAtFraction(widths, 0.1)).toBe(0);
		expect(columnIndexAtFraction(widths, 0.3)).toBe(1);
		expect(columnIndexAtFraction(widths, 1)).toBe(2);
		expect(rowIndexAtOffset([30, 60], 10)).toBe(0);
		expect(rowIndexAtOffset([30, 60], 45)).toBe(1);
		expect(rowIndexAtOffset([30, 60], 90)).toBe(2);
	});

	it('ignores a press on a boundary inside a merged cell', () => {
		expect(isRowBoundaryMergedAt(crossings, 0, widths, 0.1)).toBe(true);
		expect(isRowBoundaryMergedAt(crossings, 0, widths, 0.6)).toBe(false);
		expect(isColumnBoundaryMergedAt(crossings, 1, [30, 60], 80)).toBe(true);
		expect(isColumnBoundaryMergedAt(crossings, 1, [30, 60], 10)).toBe(false);
	});
});
