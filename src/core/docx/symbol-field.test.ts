import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { computedFields } from './computed-fields';
import { fieldResultSpans } from './field-result-spans';
import { loadDocx } from './parse';
import { SYMBOL_ERROR, symbolFieldResult } from './symbol-field';
import { drawn } from './test-support/computed-fields';

// symbol-1.docx: SYMBOL fields Word 16 inserted and updated
// (`__fixtures__/computed-fields/generate-word-computed-fields.ps1`). Word saves them without a
// result; what it drew, read from its PDF export (each character's font and size), is below. A
// symbol font's character is its U+F000 code, as the document model holds such characters.
const fixture = await readFile(
	new URL('./__fixtures__/computed-fields/symbol-1.docx', import.meta.url),
);

type Drawn = [text: string, font?: string | undefined, size?: number | undefined];
const WORD: Drawn[][] = [
	[['A']],
	[['A']],
	[['A']],
	[['', 'Symbol']],
	[['', 'Symbol']],
	[['', 'Symbol']],
	[['€']],
	[['€']],
	[['€']],
	[['€']],
	[['€']],
	[['A', undefined, 20]],
	[['', 'Wingdings', 16]],
	[['a']],
	[[SYMBOL_ERROR]],
	[[SYMBOL_ERROR]],
	[['A']],
	[[SYMBOL_ERROR]],
	[[SYMBOL_ERROR]],
	[['ᅰ']],
	[[SYMBOL_ERROR]],
	[['é', 'Times New Roman', 9]],
	[['A', 'Arial']],
	[['“']],
	[[SYMBOL_ERROR]],
	[['A', undefined, 11], ['A'], ['A']],
	[
		['', 'Wingdings'],
		['', 'Wingdings'],
		['', 'Wingdings'],
		['a', 'Courier New'],
	],
	[['A'], ['A'], ['', 'Symbol'], [SYMBOL_ERROR]],
];

describe('SYMBOL', () => {
	it('draws the characters Word draws, in their fonts and sizes', async () => {
		const { model } = await loadDocx(fixture);
		const fields = computedFields(model);
		const shown = model.blocks.flatMap((block) => {
			if (block.type !== 'paragraph') return [];
			const results = fieldResultSpans(block.runs, { resultless: true }).map((field): Drawn => {
				const result = fields.result(field, block);
				return [result?.text ?? '', result?.fontFamily, result?.fontSizePt];
			});
			return results.length ? [results] : [];
		});
		const trim = (items: Drawn[][]) =>
			items.map((line) => line.map((item) => [item[0], item[1], item[2]]));
		expect(trim(shown)).toEqual(trim(WORD));
		expect(drawn(model)[0]?.text).toBe('xAx');
	});

	it('keeps the result of a field with a switch it does not read', () => {
		expect(symbolFieldResult('SYMBOL 65 \\z')).toBeUndefined();
		expect(symbolFieldResult('SYMBOL 65 \\h')).toEqual({ text: 'A', keepLineSpacing: true });
		expect(symbolFieldResult('SYMBOL 65 \\s 2000')).toEqual({ text: 'A', fontSizePt: 1638 });
	});
});
