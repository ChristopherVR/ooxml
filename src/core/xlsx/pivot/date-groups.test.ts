import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { ymdToSerial } from '../formula/date-serial';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';
import { dateGroupRange, dateUnitIndex, dateUnitLabels } from './date-groups';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-group-date-${name}.xlsx`, import.meta.url));
const serial = (y: number, m: number, d: number) => ymdToSerial(y, m, d, false)!;
const part = async (bytes: Uint8Array, pattern: RegExp) => {
	const zip = await JSZip.loadAsync(bytes);
	return zip.file(Object.keys(zip.files).find((n) => pattern.test(n))!)!.async('string');
};
/** The derived cache fields: what Excel wrote for each unit, without the volatile parts. */
const groupFields = (cache: string) =>
	cache.match(/<cacheField [^>]*databaseField="0".*?<\/cacheField>/g);

describe('date group items', () => {
	const range = { start: serial(2023, 11, 15), end: serial(2024, 11, 3) };
	it('list the units the way Excel names them', () => {
		const labels = (unit: Parameters<typeof dateUnitLabels>[0], days = 1) =>
			dateUnitLabels(unit, range, days, false);
		expect(labels('years')).toEqual(['<2023-11-15', '2023', '2024', '>2024-11-03']);
		expect(labels('quarters')).toEqual([
			'<2023-11-15',
			'Qtr1',
			'Qtr2',
			'Qtr3',
			'Qtr4',
			'>2024-11-03',
		]);
		expect(labels('months').slice(0, 3)).toEqual(['<2023-11-15', 'Jan', 'Feb']);
		expect(labels('hours').slice(1, 3)).toEqual(['12 AM', '1 AM']);
		expect(labels('hours').at(-2)).toBe('11 PM');
		expect(labels('minutes')).toHaveLength(62);
		expect(labels('seconds')[3]).toBe(':02');
		const days = labels('days');
		expect(days).toHaveLength(368);
		expect(days.slice(1, 3)).toEqual(['01-Jan', '02-Jan']);
		expect(days[60]).toBe('29-Feb');
		expect(days.at(-2)).toBe('31-Dec');
		const weeks = labels('days', 7);
		expect(weeks).toHaveLength(53);
		expect(weeks[1]).toBe('2023-11-15 - 2023-11-21');
		expect(weeks.at(-2)).toBe('2024-10-30 - 2024-11-03');
	});
	it('place a date before, inside and after the range', () => {
		const index = (unit: Parameters<typeof dateUnitLabels>[0], value: number, days = 1) =>
			dateUnitIndex(
				unit,
				range,
				days,
				false,
				value,
				dateUnitLabels(unit, range, days, false).length,
			);
		expect(index('months', serial(2023, 11, 14))).toBe(0);
		expect(index('months', serial(2024, 2, 14))).toBe(2);
		expect(index('months', range.end)).toBe(11);
		expect(index('months', range.end + 1)).toBe(13);
		expect(index('years', serial(2024, 1, 1))).toBe(2);
		expect(index('quarters', serial(2024, 8, 1))).toBe(3);
		expect(index('days', serial(2024, 2, 29))).toBe(60);
		expect(index('days', serial(2023, 11, 28), 7)).toBe(2);
		expect(index('hours', serial(2024, 1, 1) + 13 / 24 + 0.0001)).toBe(14);
		expect(index('minutes', serial(2024, 1, 1) + 90 / 1440)).toBe(31);
		expect(index('seconds', serial(2024, 1, 1) + 59 / 86_400)).toBe(60);
	});
	it('end one day after the latest whole-day date unless chosen', () => {
		const dates = [serial(2024, 1, 5), serial(2024, 3, 1)];
		expect(dateGroupRange({}, dates)).toEqual({ start: dates[0], end: dates[1]! + 1 });
		expect(dateGroupRange({}, [1.5, 3.25])).toEqual({ start: 1.5, end: 3.25 });
		expect(dateGroupRange({ start: 4, end: 9 }, dates)).toEqual({ start: 4, end: 9 });
	});
});

describe('native date groups', () => {
	const NAMES = [
		['months-years', ['months', 'years']],
		['months-quarters-years-tabular', ['months', 'quarters', 'years']],
		['quarters-years', ['quarters', 'years']],
		['years', ['years']],
		['days-1', ['days']],
		['days-7', ['days']],
		['time', ['seconds', 'minutes', 'hours']],
		['months-custom', ['months']],
		['year-end', ['months', 'years']],
	] as const;
	it.each(NAMES)('imports %s and rewrites the same group fields', async (name, units) => {
		const bytes = fixture(name);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.dateGroups).toHaveLength(1);
		expect(pivot.dateGroups![0]).toMatchObject({ field: 'Date', by: units });
		// Untouched parts are saved verbatim.
		const original = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const p of Object.keys(original.files).filter((n) => n.startsWith('xl/pivot')))
			expect(await saved.file(p)?.async('string'), p).toBe(await original.file(p)?.async('string'));
		// A refresh regenerates the derived fields and the rows Excel wrote.
		createEditSession(wb).refreshPivotTable(0, pivot.name);
		const out = await saveXlsx(wb);
		const cache = await part(out, /pivotCacheDefinition\d+\.xml$/);
		const native = await part(bytes, /pivotCacheDefinition1\.xml$/);
		expect(groupFields(cache)).toEqual(groupFields(native));
		expect(cache).toContain(`<fieldGroup par="${units.length + 2}"/>`);
		const rows = (xml: string) => xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<colItems'));
		expect(rows(await part(out, /pivotTable\d+\.xml$/)).replaceAll('<x v="0"/>', '<x/>')).toBe(
			rows(await part(bytes, /pivotTable1\.xml$/)).replaceAll('<x v="0"/>', '<x/>'),
		);
		const again = await loadXlsx(out);
		expect(again.warnings).toEqual([]);
		expect(pivotGrid(again.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('groups by month and year with a total per year', async () => {
		const wb = await loadXlsx(fixture('months-years'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.map((line) => line[0])).toEqual([
			'Row Labels',
			'2023',
			'Nov',
			'Dec',
			'2024',
			'Jan',
			'Feb',
			'Apr',
			'Jul',
			'Nov',
			'Grand Total',
		]);
		expect(grid.map((line) => line[1])).toEqual([
			'Sum of Amount',
			30,
			10,
			20,
			51,
			5,
			22,
			9,
			12,
			3,
			81,
		]);
	});

	it('puts dates outside a chosen range in the before and after items', async () => {
		const wb = await loadXlsx(fixture('months-custom'));
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.dateGroups![0]).toMatchObject({
			start: serial(2024, 1, 9),
			end: serial(2024, 7, 27),
		});
		expect(pivotGrid(pivot).values.map((line) => [line[0], line[1]])).toEqual([
			['Months (Date)', 'Sum of Amount'],
			['<2024-01-09', 35],
			['Feb', 22],
			['Apr', 9],
			['>2024-07-27', 15],
			['Grand Total', 81],
		]);
	});

	it('keeps every year the range touches, even a year with no dates', async () => {
		const wb = await loadXlsx(fixture('year-end'));
		const cache = await part(await saveXlsx(wb), /pivotCacheDefinition1\.xml$/);
		expect(cache).toContain('<s v="2025"/>');
	});
});

describe('grouping dates through an edit session', () => {
	const DATES: [number, number, number][] = [
		[2023, 11, 15],
		[2023, 12, 20],
		[2024, 1, 5],
		[2024, 2, 14],
		[2024, 4, 9],
	];
	const setup = () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		edit.setCellInput(0, 0, 0, 'Date');
		edit.setCellInput(0, 0, 1, 'Amount');
		DATES.forEach(([y, m, d], i) => {
			edit.setCellInput(0, i + 1, 0, String(serial(y, m, d)));
			edit.setCellInput(0, i + 1, 1, String((i + 1) * 10));
		});
		edit.applyStyle(0, [{ start: { row: 1, col: 0 }, end: { row: 5, col: 0 } }], {
			numFmt: 'yyyy-mm-dd',
		});
		return { wb, edit };
	};
	const base = { source: 'Data!A1:B6', rowField: 'Date', valueField: 'Amount' };
	const pivotOf = (wb: ReturnType<typeof setup>['wb']) => wb.sheets[1]!.pivotTables![0]!;

	it('replaces the date field on the rows with its group fields', async () => {
		const { wb, edit } = setup();
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				...base,
				layout: 'compact',
				dateGroups: [{ field: 'Date', by: ['months', 'years'] }],
			},
		);
		const pivot = pivotOf(wb);
		expect(pivot).toMatchObject({ rowField: 'Years (Date)', secondRowField: 'Months (Date)' });
		expect(pivotGrid(pivot).values.map((line) => line[0])).toEqual([
			'Row Labels',
			'2023',
			'Nov',
			'Dec',
			'2024',
			'Jan',
			'Feb',
			'Apr',
			'Grand Total',
		]);
		const out = await saveXlsx(wb);
		const cache = await part(out, /pivotCacheDefinition\d+\.xml$/);
		expect(cache).toContain('<cacheField name="Months (Date)" numFmtId="0" databaseField="0">');
		expect(cache).toContain(
			'<rangePr groupBy="years" startDate="2023-11-15T00:00:00" endDate="2024-04-10T00:00:00"/>',
		);
		expect(cache).toContain('<fieldGroup par="3"/>');
		const reloaded = await loadXlsx(out);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[1]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('regroups, ungroups and undoes like Excel', () => {
		const { wb, edit } = setup();
		const name = edit.createPivotTable(1, { row: 0, col: 0 }, base);
		const rows = () => pivotGrid(pivotOf(wb)).values.map((line) => line[0]);
		edit.refreshPivotTable(1, name, { dateGroups: [{ field: 'Date', by: ['years'] }] });
		expect(pivotOf(wb).rowField).toBe('Years (Date)');
		expect(rows()).toEqual(['Years (Date)', '2023', '2024', 'Grand Total']);
		edit.refreshPivotTable(1, name, { dateGroups: [{ field: 'Date', by: ['years', 'quarters'] }] });
		expect(pivotOf(wb)).toMatchObject({
			rowField: 'Years (Date)',
			secondRowField: 'Quarters (Date)',
		});
		edit.refreshPivotTable(1, name, {
			dateGroups: [{ field: 'Date', by: ['days'], days: 30, start: serial(2023, 11, 1) }],
		});
		expect(pivotOf(wb).rowField).toBe('Days (Date)');
		expect(rows()[1]).toBe('2023-11-01 - 2023-11-30');
		edit.refreshPivotTable(1, name, { dateGroups: [] });
		expect(pivotOf(wb)).toMatchObject({ rowField: 'Date' });
		expect(pivotOf(wb).dateGroups).toBeUndefined();
		edit.undo();
		expect(pivotOf(wb).rowField).toBe('Days (Date)');
	});

	it('refuses groups that cannot work', () => {
		const { edit } = setup();
		edit.setCellInput(0, 1, 1, 'text');
		const group =
			(extra: object, options: object = {}) =>
			() =>
				edit.createPivotTable(1, { row: 0, col: 0 }, {
					...base,
					...options,
					dateGroups: [{ field: 'Date', by: ['months'], ...extra }],
				} as never);
		expect(group({ by: [] })).toThrow('Choose a valid date group');
		expect(group({ by: ['weeks'] })).toThrow('Choose a valid date group');
		expect(group({ days: 7 })).toThrow('Choose a valid date group');
		expect(group({ start: 10, end: 5 })).toThrow('must end after it starts');
		expect(group({ field: 'Amount' })).toThrow('Only fields that contain dates');
	});
});
