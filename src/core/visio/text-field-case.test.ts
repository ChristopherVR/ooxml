import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { visioFieldCase } from './text-field-case';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';

const corpus = JSON.parse(
	readFileSync(new URL('./__fixtures__/field-case-native.json', import.meta.url), 'utf8').replace(
		/^\uFEFF/,
		'',
	),
) as {
	cases: { input: string; '38': string; '39': string }[];
};

describe('native field case formats', () => {
	for (const [index, sample] of corpus.cases.entries()) {
		it(`matches native lower and upper mappings in scalar chunk ${index + 1}`, () => {
			expect(visioFieldCase(sample.input, false)).toBe(sample['38']);
			expect(visioFieldCase(sample.input, true)).toBe(sample['39']);
		});
	}
	it('uses instance text with native casing and retains it through an unrelated edit', async () => {
		const bytes = readFileSync(new URL('./__fixtures__/field-case-native.vsdx', import.meta.url));
		const source = await parseVsdx(bytes);
		const expected = ['Straße café İΣ', 'straße café iσ', 'STRAßE CAFÉ İΣ'];
		expect(source.pages[0]!.shapes.map((shape) => shape.text.plainText)).toEqual(expected);
		const result = await editVsdx(bytes, [
			{ type: 'rename-page', pageId: source.pages[0]!.id, name: 'Cased fields' },
		]);
		expect(
			(await parseVsdx(result.bytes)).pages[0]!.shapes.map((shape) => shape.text.plainText),
		).toEqual(expected);
	});
	it('preserves punctuation, uncased scalars and string length', () => {
		expect(visioFieldCase('123 🙂 ß', true)).toBe('123 🙂 ß');
		expect(visioFieldCase('ΑΣ ΟΣ', false)).toBe('ασ οσ');
	});
});
