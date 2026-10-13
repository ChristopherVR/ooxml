import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { numericFieldPicture } from './text-field-numeric';

it('matches 46 native general numeric fields after inherited cache projection and edit/reopen', async () => {
	const evidence = JSON.parse(
		readFileSync(
			new URL('./__fixtures__/numeric-field-general-native.json', import.meta.url),
			'utf8',
		).replace(/^\uFEFF/, ''),
	) as { cases: { id: number; code: number; formula: string; text: string }[] };
	expect(evidence.cases).toHaveLength(46);
	const bytes = readFileSync(
		new URL('./__fixtures__/numeric-field-general-native.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const edited = await editVsdx(bytes, [
		{ type: 'rename-page', pageId: source.id, name: 'General' },
	]);
	for (const page of [source, (await parseVsdx(edited.bytes)).pages[0]!])
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				`${item.code}: ${item.formula}`,
			).toBe(item.text);
});

it('keeps unverified large general values and page-dependent units at stored text', () => {
	expect(
		numericFieldPicture({ kind: 'number', value: 2e15, unit: 'scalar' }, 0, 'NUM'),
	).toBeUndefined();
	expect(
		numericFieldPicture({ kind: 'number', value: 1, unit: 'length' }, 1, 'DL'),
	).toBeUndefined();
	for (const value of [NaN, Infinity, -Infinity])
		expect(
			numericFieldPicture({ kind: 'number', value, unit: 'scalar' }, 0, 'NUM'),
		).toBeUndefined();
});
