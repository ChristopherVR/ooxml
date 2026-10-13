import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import JSZip from 'jszip';

it('matches eleven explicit ISO picture fields through source-preserving edit/reopen', async () => {
	const evidence = JSON.parse(
		readFileSync(
			new URL('./__fixtures__/field-iso-date204-native.json', import.meta.url),
			'utf8',
		).replace(/^\uFEFF/, ''),
	) as { cases: { id: number; formula: string; text: string }[] };
	expect(evidence.cases).toHaveLength(12);
	const bytes = readFileSync(
		new URL('./__fixtures__/field-iso-date204-native.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const saved = await editVsdx(bytes, [
		{ type: 'rename-page', pageId: source.id, name: 'ISO dates' },
	]);
	const originalZip = await JSZip.loadAsync(bytes),
		savedZip = await JSZip.loadAsync(saved.bytes);
	for (const part of Object.keys(originalZip.files).filter(
		(name) => name === 'visio/pages/page1.xml' || /^visio\/masters\/master\d+\.xml$/.test(name),
	))
		expect(await savedZip.file(part)?.async('string')).toBe(
			await originalZip.file(part)!.async('string'),
		);
	for (const page of [source, (await parseVsdx(saved.bytes)).pages[0]!])
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				item.formula,
			).toBe(item.formula === 'DATE(1899,12,29)' ? '2026-01-01' : item.text);
});

it('formats the explicit ISO picture independently of stored locale placeholders', async () => {
	const { fixture, shape } = await import('./test-fixtures');
	for (const cached of ['01/01/2026', '1.1.2026', '2026-02-30', '']) {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						`<Section N="Field"><Row IX="0"><Cell N="Value" V="2026-10-13T00:00:00" U="DATE"/><Cell N="Format" V="ddddd" F="FIELDPICTURE(204)"/></Row></Section><Text><fld IX="0">${cached}</fld></Text>`,
					)}</Shapes>`,
				},
			],
		});
		expect((await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.plainText).toBe('2026-10-13');
	}
});
