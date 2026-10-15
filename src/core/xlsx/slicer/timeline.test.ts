import { readFileSync, writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import {
	periodsInSelection,
	selectionLabel,
	selectionOfPeriods,
	timelinePeriods,
} from './timeline-periods';
import { timelineGeometry, timelineView } from './timeline-view';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-slicer-${name}.xlsx`, import.meta.url));
const text = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';

it('builds periods at every level and converts ranges both ways', () => {
	const bounds = ['2025-01-01T00:00:00', '2026-01-01T00:00:00'] as const;
	expect(timelinePeriods(0, ...bounds).map((p) => p.label)).toEqual(['2025']);
	const quarters = timelinePeriods(1, ...bounds);
	expect(quarters.map((p) => p.label)).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
	const months = timelinePeriods(2, ...bounds);
	expect(months).toHaveLength(12);
	expect(months[1]).toMatchObject({
		label: 'Feb',
		group: '2025',
		first: '2025-02-01',
		last: '2025-02-28',
	});
	expect(timelinePeriods(3, ...bounds)).toHaveLength(365);
	const selection = selectionOfPeriods(months, 3, 1)!;
	expect(selection).toEqual({ start: '2025-02-01T00:00:00', end: '2025-04-30T00:00:00' });
	expect(periodsInSelection(months, selection)).toEqual([1, 3]);
	expect(selectionLabel(months, selection, 2)).toBe('Feb 2025 - Apr 2025');
	expect(selectionLabel(quarters, selectionOfPeriods(quarters, 0, 0), 1)).toBe('Q1 2025');
});

it('lays out a timeline: groups over period buttons with the range selected', async () => {
	const wb = await loadXlsx(fixture('timeline-range'));
	const cache = wb.slicers!.timelineCaches[0]!;
	const timeline = wb.slicers!.timelines[0]!;
	const view = timelineView(cache, timeline);
	expect(view.selected).toEqual([1, 3]);
	const geometry = timelineGeometry(timeline, view, 480, 130);
	expect(geometry.groups).toEqual([{ label: '2025', x: 0, w: 12 * 52 }]);
	expect(geometry.cells.filter((cell) => cell.selected).map((cell) => cell.label)).toEqual([
		'Feb',
		'Mar',
		'Apr',
	]);
	expect(geometry.contentWidth).toBe(12 * 52);
	expect(geometry.body.h).toBeGreaterThan(0);
});

it('filters the PivotTable with a date range, saves the filter and clears it again', async () => {
	const wb = await loadXlsx(fixture('timeline'));
	const edit = createEditSession(wb);
	const sheet = wb.sheets[0]!;
	const grid = () =>
		[2, 3, 4, 5].map((row) => [getCell(sheet, row, 0)?.value, getCell(sheet, row, 1)?.value]);
	expect(grid()).toEqual([
		['Row Labels', 'Sum of Amount'],
		['Fruit', 37],
		['Nuts', 9],
		['Veg', 20],
	]);
	edit.setTimelineSelection('Day', { start: '2025-02-01T00:00:00', end: '2025-04-30T00:00:00' });
	expect(grid()).toEqual([
		['Row Labels', 'Sum of Amount'],
		['Fruit', 20],
		['Veg', 20],
		['Grand Total', 40],
	]);
	const saved = await saveXlsx(wb);
	const xml = await text(saved, 'xl/pivotTables/pivotTable1.xml');
	expect(xml).toContain('type="dateBetween"');
	expect(xml).toContain('<customFilter operator="greaterThanOrEqual" val="45689"/>');
	expect(xml).toContain('<customFilter operator="lessThanOrEqual" val="45777"/>');
	const cache = await text(saved, 'xl/timelineCaches/timelineCache1.xml');
	expect(cache).toContain('filterType="dateBetween"');
	expect(cache).toContain(
		'<selection startDate="2025-02-01T00:00:00" endDate="2025-04-30T00:00:00"/>',
	);
	const again = await loadXlsx(saved);
	expect(again.warnings).toEqual([]);
	expect(again.sheets[0]!.pivotTables![0]!.dateFilter).toEqual({
		field: 'Day',
		start: 45689,
		end: 45777,
	});
	expect(again.slicers!.timelineCaches[0]!.selection).toEqual({
		start: '2025-02-01T00:00:00',
		end: '2025-04-30T00:00:00',
	});
	edit.undo();
	expect(grid()[1]).toEqual(['Fruit', 37]);
	edit.setTimelineSelection('Day', { start: '2025-02-01T00:00:00', end: '2025-04-30T00:00:00' });
	edit.setTimelineSelection('Day', null);
	expect(sheet.pivotTables![0]!.dateFilter).toBeUndefined();
	const cleared = await saveXlsx(wb);
	expect(await text(cleared, 'xl/pivotTables/pivotTable1.xml')).not.toContain('<filters');
	expect(await text(cleared, 'xl/timelineCaches/timelineCache1.xml')).not.toContain('<selection');
});

it('inserts, formats and deletes a timeline on a date field', async () => {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(wb);
	const rows = [
		['Category', 'Amount', 'Day'],
		['Fruit', '10', '2025-01-15'],
		['Fruit', '20', '2025-02-15'],
		['Veg', '5', '2025-03-15'],
	];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{ source: 'Data!A1:C4', rowField: 'Category', valueField: 'Amount' },
	);
	expect(wb.sheets[1]!.pivotTables![0]!.dateFields).toEqual(['Day']);
	const name = edit.insertTimeline(1, { pivot: 'PivotTable1', field: 'Day' });
	expect(name).toBe('Day');
	edit.updateTimeline('Day', { caption: 'When', level: 1 });
	edit.setTimelineSelection('NativeTimeline_Day', {
		start: '2025-02-01T00:00:00',
		end: '2025-03-31T00:00:00',
	});
	const saved = await saveXlsx(wb);
	if (process.env.OOXML_DUMP) writeFileSync(process.env.OOXML_DUMP, saved);
	const again = await loadXlsx(saved);
	expect(again.warnings).toEqual([]);
	expect(again.slicers!.timelines[0]).toMatchObject({ name: 'Day', caption: 'When', level: 1 });
	expect(again.slicers!.timelineCaches[0]!.bounds).toEqual({
		start: '2025-01-01T00:00:00',
		end: '2026-01-01T00:00:00',
	});
	const pivot = again.sheets[1]!.pivotTables![0]!;
	expect(pivot.dateFilter?.field).toBe('Day');
	expect(getCell(again.sheets[1]!, 1, 1)?.value).toBe(20);
	const second = createEditSession(again);
	second.deleteTimeline('Day');
	expect(again.slicers!.timelines).toEqual([]);
	expect(again.sheets[1]!.pivotTables![0]!.dateFilter).toBeUndefined();
	const final = await saveXlsx(again);
	const parts = Object.keys((await JSZip.loadAsync(final)).files).filter((n) =>
		/timeline/i.test(n),
	);
	expect(parts).toEqual([]);
});
