import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { formatVisioFieldValue } from './text-fields';

it('matches native decimal ties, zeros, scientific small values and large decimals', async () => {
	const evidence = JSON.parse(
		readFileSync(
			new URL('./__fixtures__/numeric-field-boundary-native.json', import.meta.url),
			'utf8',
		).replace(/^\uFEFF/, ''),
	) as { cases: { id: number; code: number; formula: string; text: string }[] };
	const bytes = readFileSync(
		new URL('./__fixtures__/numeric-field-boundary-native.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const edited = await editVsdx(bytes, [
		{ type: 'rename-page', pageId: source.id, name: 'Boundary' },
	]);
	const page = (await parseVsdx(edited.bytes)).pages[0]!;
	for (const item of evidence.cases)
		expect(
			page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
			`${item.code}: ${item.formula}`,
		).toBe(item.text);
});

it('rounds explicit decimal pictures at native half-decimal boundaries', () => {
	for (const [value, text] of [
		[1.005, '1.01'],
		[-1.005, '-1.01'],
		[2.675, '2.68'],
	] as const)
		expect(formatVisioFieldValue({ kind: 'number', value, unit: 'scalar' }, '0.00')).toBe(text);
});
