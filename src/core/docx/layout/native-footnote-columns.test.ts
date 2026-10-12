import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
const cases = [
	{ name: 'default-two-twenty', counts: [[10, 10]], x: 226.6666667, widths: [200, 200] },
	{
		name: 'two-two-twenty',
		counts: [[10, 10]],
		x: 237.3333333,
		widths: [189.3333333, 189.3333333],
	},
	{ name: 'unequal-two-twenty', counts: [[10, 10]], x: 160, widths: [133.3333333, 266.6666667] },
	{
		name: 'two-two-custom-eighty',
		counts: [
			[33, 33],
			[7, 7],
		],
		x: 237.3333333,
		widths: [189.3333333, 189.3333333],
	},
];
describe('native Word footnote area columns', () => {
	it.each(cases)(
		'$name matches native distribution without losing note lines',
		async ({ name, counts, x, widths }) => {
			const source = new Uint8Array(
				readFileSync(
					new URL('../__fixtures__/footnote-columns/' + name + '.docx', import.meta.url),
				),
			);
			const loaded = await loadDocx(source);
			const result = layoutDocumentModel(loaded.model, measurer);
			expect(loaded.model.compatibilityMode).toBe(15);
			expect(result.pages).toHaveLength(counts.length);
			const sourceIds: string[] = [];
			for (const [pageIndex, page] of result.pages.entries()) {
				const notes = page.footnotes ?? [];
				expect(notes.map((note) => note.paragraphs.flatMap((p) => p.lines).length)).toEqual(
					counts[pageIndex],
				);
				expect(notes.map((note) => note.columnIndex)).toEqual([0, 1]);
				expect(notes[0]?.xPx).toBe(0);
				expect(notes[1]?.xPx).toBeCloseTo(x, 5);
				for (const [index, note] of notes.entries()) {
					expect(note.widthPx).toBeCloseTo(widths[index]!, 5);
					for (const paragraph of note.paragraphs)
						for (const line of paragraph.lines) {
							const text = line.fragments.map((fragment) => fragment.text).join('');
							sourceIds.push(...(text.match(/N0L\d{2}/g) ?? []));
						}
				}
				if (name.includes('custom')) {
					const separatorText = page.footnoteSeparator?.paragraphs
						.flatMap((p) =>
							p.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)),
						)
						.join('');
					expect(separatorText).toContain('Sources');
					expect(separatorText).not.toContain('Continued');
				}
			}
			const total = counts.flat().reduce((a, b) => a + b, 0);
			expect(sourceIds).toEqual(
				Array.from({ length: total }, (_, index) => 'N0L' + String(index).padStart(2, '0')),
			);
		},
	);
	it('reflows a continued paragraph into unequal columns without repeating or dropping tokens', async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL('../__fixtures__/footnote-columns/unequal-two-twenty.docx', import.meta.url),
				),
			),
		);
		const model = structuredClone(loaded.model);
		const words = Array.from(
			{ length: 160 },
			(_, index) => 'word' + String(index).padStart(3, '0'),
		);
		const note = model.footnotes![0]!;
		const paragraph = note.blocks.find((block) => block.type === 'paragraph')!;
		if (paragraph.type !== 'paragraph') throw new Error('Missing source paragraph');
		paragraph.runs = [{ text: words.join(' ') }];
		note.blocks = [paragraph];
		const result = layoutDocumentModel(model, measurer);
		const rendered = result.pages
			.flatMap((page) => page.footnotes ?? [])
			.flatMap((note) =>
				note.paragraphs.flatMap((p) =>
					p.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)),
				),
			)
			.join(' ');
		expect(rendered.match(/word\d{3}/g)).toEqual(words);
		expect(
			result.pages.flatMap((page) => page.footnotes ?? []).some((note) => note.columnIndex === 1),
		).toBe(true);
	});
});
