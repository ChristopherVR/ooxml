import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { systemDatePicture } from './text-field-system-date';

it('matches inherited native system date/time profiles without rewriting their source', async () => {
	const evidence = JSON.parse(
		readFileSync(new URL('./__fixtures__/field-system-date-native.json', import.meta.url), 'utf8'),
	) as { cases: { id: number; code: number; formula: string; text: string }[] };
	const bytes = readFileSync(
		new URL('./__fixtures__/field-system-date-native.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const saved = await editVsdx(bytes, [
		{ type: 'rename-page', pageId: source.id, name: 'System dates' },
	]);
	const originalZip = await JSZip.loadAsync(bytes),
		savedZip = await JSZip.loadAsync(saved.bytes);
	for (const part of Object.keys(originalZip.files).filter(
		(name) => name === 'visio/pages/page1.xml' || /^visio\/masters\/master\d+\.xml$/.test(name),
	))
		expect(await savedZip.file(part)?.async('string')).toBe(
			await originalZip.file(part)!.async('string'),
		);
	expect(evidence.cases).toHaveLength(30);
	for (const page of [source, (await parseVsdx(saved.bytes)).pages[0]!])
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				`${item.code}: ${item.formula}`,
			).toBe(
				item.formula === 'DATE(1899,12,29)'
					? item.code === 200
						? '2026-01-01'
						: item.code === 211
							? '2026-01-01 1:02 AM'
							: '2026-01-01 1:02:03 AM'
					: item.text,
			);
});

it.each([
	['200', '01/01/2026'],
	['200', '2026-02-30'],
	['200', ''],
	['211', '2026-01-01 01:02 AM'],
	['211', '2026-01-01 1:02:03 AM'],
	['211', '2026-01-01 13:02 AM'],
	['211', '2026-01-01 1:60 AM'],
	['212', '2026-02-30 1:02:03 AM'],
	['212', '2026-01-01 1:02 AM'],
	['212', '2026-01-01 1:02:03 am'],
	['212', '01/01/2026 1:02:03 AM'],
	['201', 'Thursday, January 1, 2026'],
])('retains an unmeasured or invalid system profile %s: %s', (code, cached) => {
	expect(systemDatePicture(new Date('2026-10-13T14:15:16Z'), code!, cached!)).toBeUndefined();
});

it.each(['200', '211', '212'])('refuses out-of-range source dates for picture %s', (code) => {
	const cached =
		code === '200' ? '2026-01-01' : code === '211' ? '2026-01-01 1:02 AM' : '2026-01-01 1:02:03 AM';
	for (const date of [
		new Date(NaN),
		new Date('1899-12-29T00:00:00Z'),
		new Date('+010000-01-01T00:00:00Z'),
	])
		expect(systemDatePicture(date, code, cached)).toBeUndefined();
});
