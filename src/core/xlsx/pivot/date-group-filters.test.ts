import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url));
const parts = async (bytes: Uint8Array) => {
	const zip = await JSZip.loadAsync(bytes);
	const part = Object.keys(zip.files).find((n) => /pivotTables\/pivotTable\d+\.xml$/.test(n))!;
	const xml = await zip.file(part)!.async('string');
	const plain = (text: string) => text.replaceAll('<x v="0"/>', '<x/>');
	return {
		items: (xml.match(/<items .*?<\/items>/g) ?? []).join(''),
		// Excel writes the page field's attributes in another order.
		axes: plain(xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<dataFields'))).replace(
			/ hier="-1"/g,
			'',
		),
		filters: xml.match(/<filters .*?<\/filters>/)?.[0],
	};
};

const CASES = [
	['group-page-all', ['Years (Date)', 'Months (Date)'], 'Region'],
	['group-page-west', ['Years (Date)', 'Months (Date)'], 'Region'],
	['group-page-tabular-quarters', ['Years (Date)', 'Quarters (Date)', 'Months (Date)'], 'Category'],
	['group-page-columns', ['Years (Date)', 'Months (Date)'], 'Category'],
	['group-hidden-items', ['Years (Date)', 'Months (Date)'], undefined],
	['group-slicer', ['Years (Date)', 'Months (Date)'], undefined],
	['group-timeline', ['Years (Date)', 'Months (Date)'], undefined],
	['group-timeline-days', ['Days (Date)'], undefined],
	['group-page-timeline', ['Years (Date)', 'Months (Date)'], 'Region'],
] as const;

describe('date groups with report filters, hidden items, slicers and timelines', () => {
	it.each(CASES)('imports %s and reproduces its parts', async (name, rows, page) => {
		const bytes = fixture(name);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(
			[pivot.rowField, pivot.secondRowField, ...(pivot.moreRowFields ?? [])].filter(Boolean),
		).toEqual(rows);
		expect(pivot.pageFilter?.field).toBe(page);
		expect(pivot.dateGroups).toHaveLength(1);
		const original = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(original.files).filter((n) =>
			/^xl\/(pivot|slicer|timeline)/.test(n),
		))
			expect(await saved.file(part)?.async('string'), part).toBe(
				await original.file(part)?.async('string'),
			);
		const before = await parts(bytes);
		createEditSession(wb).refreshPivotTable(0, pivot.name);
		const out = await saveXlsx(wb);
		expect(await parts(out)).toEqual(before);
		const reloaded = await loadXlsx(out);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('keeps the group range of the whole cache under a report filter', async () => {
		const wb = await loadXlsx(fixture('group-page-west'));
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		// Only the West records remain, yet the automatic range still starts at the earliest date.
		expect(pivotGrid(pivot).values.map((line) => line[0])).toEqual([
			'Region',
			null,
			'Row Labels',
			'2023',
			'Dec',
			'2024',
			'Feb',
			'Apr',
			'Nov',
			'Grand Total',
		]);
	});

	it('applies a timeline range to the grouped dates', async () => {
		const wb = await loadXlsx(fixture('group-timeline'));
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.dateFilter).toBeDefined();
		expect(pivotGrid(pivot).values).toEqual([
			['Row Labels', 'Sum of Amount'],
			['2024', 36],
			['Jan', 5],
			['Feb', 22],
			['Apr', 9],
			['Grand Total', 36],
		]);
	});

	it('adds a timeline to the grouped date field and filters the groups', async () => {
		const wb = await loadXlsx(fixture('group-page-all'));
		const session = createEditSession(wb);
		const name = session.insertTimeline(0, { pivot: 'DateFilterPivot', field: 'Date' });
		session.setTimelineSelection(`NativeTimeline_${name}`, {
			start: '2024-01-01T00:00:00',
			end: '2024-06-30T00:00:00',
		});
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.dateFilter).toMatchObject({ field: 'Date' });
		expect(pivotGrid(pivot).values.map((line) => line[0])).toEqual([
			'Region',
			null,
			'Row Labels',
			'2024',
			'Jan',
			'Feb',
			'Apr',
			'Grand Total',
		]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(reloaded.sheets[0]!.pivotTables![0]!.dateFilter).toEqual(pivot.dateFilter);
	});

	it('creates a grouped report with a report filter in an edit session', async () => {
		const wb = await loadXlsx(fixture('group-page-west'));
		const session = createEditSession(wb);
		session.refreshPivotTable(0, 'DateFilterPivot', {
			pageFilter: { field: 'Category', item: 'Fruit' },
		});
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.pageFilter).toEqual({ field: 'Category', item: 'Fruit' });
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});
});
