import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { numericFieldPicture } from './text-field-numeric';

const native = () =>
	readFileSync(new URL('./__fixtures__/numeric-field-native.vsdx', import.meta.url));
const evidence = JSON.parse(
	readFileSync(
		new URL('./__fixtures__/numeric-field-native.json', import.meta.url),
		'utf8',
	).replace(/^\uFEFF/, ''),
) as { cases: { id: number; text: string }[] };

describe('native decimal text field pictures', () => {
	it('uses the inherited instance numeric cache and original unit for all 64 native cases', async () => {
		const page = (await parseVsdx(native())).pages[0]!;
		for (const item of evidence.cases)
			expect(
				page.shapes.find((shape) => shape.id === String(item.id))?.text.plainText,
				String(item.id),
			).toBe(item.text);
	});
	it('preserves the native display through an unrelated edit and reopen', async () => {
		const source = native();
		const page = (await parseVsdx(source)).pages[0]!;
		const result = await editVsdx(source, [
			{ type: 'rename-page', pageId: page.id, name: 'Numbers' },
		]);
		const reopened = (await parseVsdx(result.bytes)).pages[0]!;
		expect(reopened.shapes.map((shape) => shape.text.plainText)).toEqual(
			evidence.cases.map((item) => item.text),
		);
	});
	it('refuses unknown pictures, page-dependent units and mismatched dimensions', () => {
		const length = { kind: 'number', value: 2, unit: 'length' } as const;
		expect(numericFieldPicture(length, 3, 'DL')).toBeUndefined();
		expect(numericFieldPicture(length, 3, 'DEG')).toBeUndefined();
		expect(numericFieldPicture(length, 15, 'IN')).toBeUndefined();
		expect(numericFieldPicture(length, 3.5, 'IN')).toBeUndefined();
	});
});
