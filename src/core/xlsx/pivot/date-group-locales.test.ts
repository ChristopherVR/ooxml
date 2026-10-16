import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

/**
 * Only the English Excel is at hand, so another language is imitated: the real-Excel fixture has
 * its group field names, month names and quarter names (in the cache and in the cached cells)
 * rewritten the way a German Excel names them. The structure stays Excel's own.
 */
const GERMAN: [string, string][] = [
	['Years (Date)', 'Jahre (Date)'],
	['Quarters (Date)', 'Quartale (Date)'],
	['Months (Date)', 'Monate (Date)'],
	['Mar', 'Mrz'],
	['May', 'Mai'],
	['Oct', 'Okt'],
	['Dec', 'Dez'],
	['Qtr1', 'Quartal 1'],
	['Qtr2', 'Quartal 2'],
	['Qtr3', 'Quartal 3'],
	['Qtr4', 'Quartal 4'],
];
async function german(name: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url)),
	);
	for (const part of Object.keys(zip.files)) {
		if (!/sharedStrings\.xml$|pivotCacheDefinition\d+\.xml$/.test(part)) continue;
		let text = await zip.file(part)!.async('string');
		for (const [english, local] of GERMAN)
			text = text
				.replaceAll(`>${english}<`, `>${local}<`)
				.replaceAll(`>${english} Total<`, `>${local} Total<`)
				.replaceAll(`"${english}"`, `"${local}"`);
		zip.file(part, text);
	}
	return zip.generateAsync({ type: 'uint8array' });
}

describe('date groups in another language', () => {
	it('reads the group field names and item labels the file uses', async () => {
		const wb = await loadXlsx(await german('group-date-months-quarters-years-tabular'));
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect([pivot.rowField, pivot.secondRowField, ...(pivot.moreRowFields ?? [])]).toEqual([
			'Jahre (Date)',
			'Quartale (Date)',
			'Monate (Date)',
		]);
		const group = pivot.dateGroups![0]!;
		expect(group.by).toEqual(['months', 'quarters', 'years']);
		expect(group.localized!.months!.labels.slice(1, 13)).toEqual([
			'Jan',
			'Feb',
			'Mrz',
			'Apr',
			'Mai',
			'Jun',
			'Jul',
			'Aug',
			'Sep',
			'Okt',
			'Nov',
			'Dez',
		]);
		expect(group.localized!.quarters!.name).toBe('Quartale (Date)');
		const labels = pivotGrid(pivot).values.flat();
		expect(labels).toContain('Quartal 4');
		expect(labels).toContain('Dez');
		expect(labels).not.toContain('Dec');
	});

	it('writes the same labels back and keeps them on refresh', async () => {
		const bytes = await german('group-date-months-years');
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const original = await JSZip.loadAsync(bytes);
		const cache = (zip: JSZip) =>
			zip.file(Object.keys(zip.files).find((n) => /pivotCacheDefinition\d+\.xml$/.test(n))!)!;
		expect(await cache(await JSZip.loadAsync(await saveXlsx(wb))).async('string')).toBe(
			await cache(original).async('string'),
		);
		createEditSession(wb).refreshPivotTable(0, wb.sheets[0]!.pivotTables![0]!.name);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		const groupItems = (xml: string) => (xml.match(/<groupItems.*?<\/groupItems>/g) ?? []).join('');
		const before = groupItems(await cache(original).async('string'));
		expect(groupItems(await cache(saved).async('string'))).toBe(before);
		expect(before).toContain('v="Dez"');
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values.flat()).toContain('Dez');
	});

	it('falls back to the English items when the range changes', async () => {
		const wb = await loadXlsx(await german('group-date-months-years'));
		const session = createEditSession(wb);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		const [group] = pivot.dateGroups!;
		// A new range has another number of year items, so the saved German years cannot apply.
		session.refreshPivotTable(0, pivot.name, {
			dateGroups: [{ field: group!.field, by: group!.by, start: 45000, end: 46500 }],
		});
		const updated = wb.sheets[0]!.pivotTables![0]!;
		expect(updated.dateGroups![0]!.localized).toBeUndefined();
		expect(updated.rowField).toBe('Years (Date)');
	});

	it('refuses a group whose items cannot be told apart', async () => {
		const zip = await JSZip.loadAsync(
			readFileSync(
				new URL('../__fixtures__/excel-pivot-group-date-months-years.xlsx', import.meta.url),
			),
		);
		const part = Object.keys(zip.files).find((n) => /pivotCacheDefinition\d+\.xml$/.test(n))!;
		const text = (await zip.file(part)!.async('string')).replace('<s v="Feb"/>', '<s v="Jan"/>');
		zip.file(part, text);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.warnings).toHaveLength(1);
	});
});
