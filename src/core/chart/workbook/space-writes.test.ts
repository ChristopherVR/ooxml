import { describe, expect, it } from 'vitest';

import type { ChartSpace } from '../model';
import { applyChartSpaceDataSheetEdits } from '../edit/data-sheet';
import { newChartSpace } from '../edit/new-chart';
import { chartSeriesEntries } from '../edit/draft';
import { loadXlsx } from '../../xlsx/read/index';
import { createChartWorkbookFromRows } from './create';
import { chartSpaceRangeConflicts, chartSpaceWorkbookWrites } from './space-writes';
import { writeChartWorkbookUpdates } from './writer';

/** A new chart with `edit` applied to a copy of its first series. */
function withFirstSeries(
	edit: (series: ChartSpace['plotArea']['groups'][number]['series'][number]) => void,
) {
	const space = structuredClone(newChartSpace());
	edit(chartSeriesEntries(space)[0]!.series);
	return space;
}

describe('chartSpaceRangeConflicts', () => {
	it('finds none where every range holds its cache', () => {
		expect(chartSpaceRangeConflicts(newChartSpace())).toEqual([]);
	});

	it('names a range of another size than its cache', () => {
		const space = withFirstSeries((series) => {
			series.values!.cache!.points.pop();
			series.values!.cache!.pointCount = 3;
		});
		expect(chartSpaceRangeConflicts(space)).toEqual(['Sheet1!$B$2:$B$5']);
	});

	it('names a range two series cache differently, but not one they share alike', () => {
		const space = structuredClone(newChartSpace());
		const [first, second] = chartSeriesEntries(space).map(({ series }) => series);
		second!.values!.formula = first!.values!.formula!;
		expect(chartSpaceRangeConflicts(space)).toEqual(['Sheet1!$B$2:$B$5']);
		second!.values!.cache = structuredClone(first!.values!.cache!);
		expect(chartSpaceRangeConflicts(space)).toEqual([]);
	});

	it('leaves out a conflict the part it replaces already had', () => {
		// A chart over a hidden row caches one point fewer than its range names, edit after edit.
		const hidden = (value: string) =>
			withFirstSeries((series) => {
				series.values!.cache!.points.pop();
				series.values!.cache!.pointCount = 3;
				series.values!.cache!.points[0]!.value = value;
			});
		expect(chartSpaceRangeConflicts(hidden('9'), hidden('1'))).toEqual([]);
		const shorter = withFirstSeries((series) => {
			series.values!.cache!.points.splice(2);
			series.values!.cache!.pointCount = 2;
		});
		expect(chartSpaceRangeConflicts(shorter, hidden('1'))).toEqual(['Sheet1!$B$2:$B$5']);
	});

	it('leaves a reference that names no plain range alone', () => {
		const space = withFirstSeries((series) => {
			series.values!.formula = 'SalesTotals';
			series.values!.cache!.points.pop();
		});
		expect(chartSpaceRangeConflicts(space)).toEqual([]);
	});
});

describe('chartSpaceWorkbookWrites for a range the chart did not read before', () => {
	it('leaves the range holding only what the chart caches', async () => {
		// Column B was a series removed earlier; a later save brings it back with blank points.
		const before = withFirstSeries((series) => {
			series.values!.formula = 'Sheet1!$E$2:$E$5';
		});
		const after = withFirstSeries((series) => {
			series.values!.cache!.points.splice(1, 2);
		});
		const writes = chartSpaceWorkbookWrites(after, before);
		expect(writes.find((write) => write.formula === 'Sheet1!$B$2:$B$5')?.only).toBe(true);
		const bytes = await createChartWorkbookFromRows([
			[null, 'Series 1'],
			['a', 1],
			['b', 99],
			['c', 'note'],
			['d', 4],
		]);
		const book = await loadXlsx((await writeChartWorkbookUpdates(bytes, writes)).bytes!);
		const column = [1, 2, 3, 4].map((row) => book.sheets[0]!.rows.get(row)?.get(1)?.value ?? null);
		// The cached points are written; under the blanks a stale number is emptied, text stays.
		expect(column).toEqual([4.3, null, 'note', 4.5]);
	});
});

describe('chartSpaceWorkbookWrites for row and series removal', () => {
	it('detects removed category rows and omits cell writes for shifted unchanged cells', () => {
		const space = newChartSpace();
		const edited = applyChartSpaceDataSheetEdits(space, [{ kind: 'remove-category', category: 1 }]);
		const writes = chartSpaceWorkbookWrites(edited, space);
		expect(writes.length).toBeGreaterThan(0);
		for (const write of writes) {
			expect(write.removedRows).toEqual([1]);
			expect(write.totalRowsBefore).toBe(4);
			expect(write.cells).toEqual([]);
		}
	});

	it('emits cell writes only for surviving rows that were explicitly edited', () => {
		const space = newChartSpace();
		const edited = applyChartSpaceDataSheetEdits(space, [
			{ kind: 'remove-category', category: 1 },
			{ kind: 'value', series: 0, category: 1, value: 9.9 },
		]);
		const writes = chartSpaceWorkbookWrites(edited, space);
		const series1Write = writes.find((w) => w.formula.includes('$B$2'));
		expect(series1Write).toBeDefined();
		expect(series1Write?.removedRows).toEqual([1]);
		expect(series1Write?.cells).toEqual([{ index: 1, value: '9.9' }]);
	});

	it('detects removed series and attaches seriesCount', () => {
		const space = newChartSpace();
		const edited = applyChartSpaceDataSheetEdits(space, [{ kind: 'remove-series', series: 2 }]);
		const writes = chartSpaceWorkbookWrites(edited, space);
		expect(writes.length).toBeGreaterThan(0);
		for (const write of writes) {
			expect(write.seriesCount).toBe(2);
		}
	});
});
