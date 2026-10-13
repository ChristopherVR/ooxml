import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { fractionalFieldNumber } from './text-field-fraction';

it('matches 120 native fractional inherited fields before and after edit/reopen', async () => {
	const evidence = JSON.parse(
		readFileSync(
			new URL('./__fixtures__/numeric-field-fraction-native.json', import.meta.url),
			'utf8',
		).replace(/^\uFEFF/, ''),
	) as { cases: { id: number; code: number; formula: string; text: string }[] };
	expect(evidence.cases).toHaveLength(120);
	const bytes = readFileSync(
		new URL('./__fixtures__/numeric-field-fraction-native.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const edited = await editVsdx(bytes, [
		{ type: 'rename-page', pageId: source.id, name: 'Fractions' },
	]);
	for (const page of [source, (await parseVsdx(edited.bytes)).pages[0]!])
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				`${item.code}: ${item.formula}`,
			).toBe(item.text);
});

it('preserves unverified large and nonfinite values', () => {
	for (const value of [2e15, NaN, Infinity, -Infinity])
		expect(fractionalFieldNumber(value, 99)).toBeUndefined();
});
