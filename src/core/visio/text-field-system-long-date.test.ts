import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { systemDatePicture } from './text-field-system-date';
import { fixture, shape } from './test-fixtures';

it('matches inherited native system date/time profiles without rewriting their source', async () => {
	const evidence = JSON.parse(
		readFileSync(
			new URL('./__fixtures__/field-system-long-date-native.json', import.meta.url),
			'utf8',
		),
	) as { cases: { id: number; code: number; formula: string; text: string }[] };
	const bytes = readFileSync(
		new URL('./__fixtures__/field-system-long-date-native.vsdx', import.meta.url),
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
	expect(evidence.cases).toHaveLength(44);
	for (const page of [source, (await parseVsdx(saved.bytes)).pages[0]!])
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				`${item.code}: ${item.formula}`,
			).toBe(item.formula === 'DATE(1899,12,29)' ? 'Thursday, January 1, 2026' : item.text);
});

const refused = [
	'Friday, January 1, 2026',
	'Thursday, February 29, 2026',
	'Saturday, April 31, 2026',
	'Thursday, Jan 1, 2026',
	'Thursday, January 01, 2026',
	'thursday, January 1, 2026',
	'Donnerstag, 1. Januar 2026',
	'jeudi 1 janvier 2026',
	'Thursday 1 January 2026',
	'January 1, 2026',
	'2026-01-01',
	'',
	'Wednesday, February 29, 1900',
	'Friday, December 29, 1899',
	'Saturday, January 1, 0000',
];
for (const code of ['21', '201']) {
	it.each(refused)(
		'retains invalid or unmeasured long-date ' + code + ' cache: %s',
		async (cached) => {
			expect(systemDatePicture(new Date('2026-10-13T14:15:16Z'), code, cached)).toBeUndefined();
			const bytes = await fixture({
				pages: [
					{
						id: '0',
						contents:
							'<Shapes>' +
							shape(
								'1',
								`<Section N="Field"><Row IX="0"><Cell N="Value" V="2026-10-13T14:15:16" U="DATE"/><Cell N="Format" V="dddddd" F="FIELDPICTURE(${code})"/></Row></Section><Text><fld IX="0">${cached}</fld></Text>`,
							) +
							'</Shapes>',
					},
				],
			});
			expect((await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.plainText).toBe(cached);
		},
	);
	it('validates every native stored English weekday/month profile for code ' + code, () => {
		const evidence = JSON.parse(
			readFileSync(
				new URL('./__fixtures__/field-system-long-date-native.json', import.meta.url),
				'utf8',
			),
		) as { cases: { formula: string; text: string }[] };
		for (const item of evidence.cases.filter((item) => item.formula !== 'DATE(1899,12,29)'))
			expect(systemDatePicture(new Date('2026-10-13T14:15:16Z'), code, item.text)).toBe(
				'Tuesday, October 13, 2026',
			);
		for (const date of [
			new Date(NaN),
			new Date('1899-12-29T00:00:00Z'),
			new Date('+010000-01-01T00:00:00Z'),
		])
			expect(systemDatePicture(date, code, 'Thursday, January 1, 2026')).toBeUndefined();
	});
}
