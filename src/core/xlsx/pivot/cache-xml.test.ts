// A date field with more distinct items than a call's argument limit (about 125,000 in Node):
// the cached date range must not pass one argument per item.
import { expect, it } from 'vitest';
import type { PivotTable } from '../model';
import { createWorkbook } from '../workbook';
import { pivotCacheData, pivotCacheDefinitionXml } from './cache-xml';

const ROWS = 200_000;

/** The `sharedItems` attributes of a date field of `ROWS` date-times whose latest is `latest`. */
function dateItems(latest: number): string | undefined {
	// Date-times a minute apart in 2025, with the earliest and latest mid-column.
	const day = (row: number) =>
		row === 70_000 ? 43_626.25 : row === 130_000 ? latest : 45_658 + row / 1440;
	const pivot: PivotTable = {
		name: 'PivotTable1',
		source: `Data!A1:B${ROWS + 1}`,
		rowField: 'Category',
		valueField: 'Category',
		aggregate: 'count',
		valueCaption: 'Count of Category',
		fields: ['Category', 'Day'],
		records: Array.from({ length: ROWS }, (_, row) => ['Fruit', day(row)]),
		dateFields: ['Day'],
		location: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
	};
	const xml = pivotCacheDefinitionXml(
		pivot,
		createWorkbook({ sheets: ['Data'] }),
		pivotCacheData(pivot),
		'rId1',
	);
	return /<cacheField name="Day"[^>]*><sharedItems([^>]*)>/.exec(xml)?.[1];
}

it('writes the date range of a field with many distinct dates', () => {
	const items = dateItems(47_912);
	// 43626.25 is 2019-06-10 06:00; the range ends a day after the latest whole day, 2031-03-05.
	expect(items).toContain(' minDate="2019-06-10T06:00:00"');
	expect(items).toContain(' maxDate="2031-03-06T00:00:00"');
	expect(items).toContain(` count="${ROWS}"`);
});

it('ends the range at a latest date-time itself, not a day after it', () => {
	// 47912.75 is 2031-03-05 18:00: only a whole-day latest date moves the end on a day.
	expect(dateItems(47_912.75)).toContain(' maxDate="2031-03-05T18:00:00"');
});
