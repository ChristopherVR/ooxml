import { expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout.js';
import type { LayoutParagraph } from './input.js';

const measure = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const layout = (runs: LayoutParagraph['runs'], width = 300) =>
	layoutParagraph({ kind: 'paragraph', id: 'p', runs }, width, measure, () => {});

it('keeps source ranges across formatting, tabs and wrapped trailing spaces', () => {
	const result = layout([{ text: 'one ', bold: true }, { text: 'two\tthree' }], 80);
	const fragments = result.lines.flatMap((line) => line.fragments);
	expect(fragments.filter((fragment) => fragment.text)).toMatchObject([
		{ text: 'one', sourceStart: 0, sourceEnd: 3 },
		{ text: ' ', sourceStart: 3, sourceEnd: 4 },
		{ text: 'two', sourceStart: 4, sourceEnd: 7 },
		{ text: 'three', sourceStart: 8, sourceEnd: 13 },
	]);
	expect(fragments.find((fragment) => !fragment.text)).toMatchObject({
		sourceStart: 7,
		sourceEnd: 8,
	});
});

it('maps generated list markers to the paragraph start without shifting body text', () => {
	const result = layout([
		{ text: '12.\t', synthetic: true, marker: { length: 3, alignment: 'right' } },
		{ text: 'Body' },
	]);
	expect(result.lines[0]!.fragments).toMatchObject([
		{ text: '12.', sourceStart: 0, sourceEnd: 0 },
		{ text: '', sourceStart: 0, sourceEnd: 0 },
		{ text: 'Body', sourceStart: 0, sourceEnd: 4 },
	]);
});

it('retains UTF-16 ranges and source positions on continuation lines', () => {
	const result = layout([{ text: 'A😀 B C' }], 40);
	expect(result.lines.flatMap((line) => line.fragments)).toMatchObject([
		{ text: 'A😀', sourceStart: 0, sourceEnd: 3 },
		{ text: 'B', sourceStart: 4, sourceEnd: 5 },
		{ text: ' ', sourceStart: 5, sourceEnd: 6 },
		{ text: 'C', sourceStart: 6, sourceEnd: 7 },
	]);
});
