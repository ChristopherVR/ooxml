import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
const variants = [
	{
		name: 'explicit-two-eighty',
		counts: [[33], [34, 13]],
		indexes: [[1], [0, 1]],
		widths: [200, 200],
		x: 226.6666667,
	},
	{
		name: 'two-custom-eighty',
		counts: [[33], [34, 13]],
		indexes: [[1], [0, 1]],
		widths: [200, 200],
		x: 226.6666667,
	},
	{ name: 'two-custom-twenty', counts: [[20]], indexes: [[1]], widths: [200, 200], x: 226.6666667 },
	{
		name: 'left-custom-sixty',
		counts: [[33, 27]],
		indexes: [[0, 1]],
		widths: [200, 200],
		x: 226.6666667,
	},
	{
		name: 'left-custom-eighty',
		counts: [[33, 34], [13]],
		indexes: [[0, 1], [0]],
		widths: [200, 200],
		x: 226.6666667,
	},
	{
		name: 'unequal-custom-eighty',
		counts: [[33], [34, 13]],
		indexes: [[1], [0, 1]],
		widths: [133.3333333, 266.6666667],
		x: 160,
	},
	{
		name: 'two-standard-eighty',
		counts: [[33], [34, 13]],
		indexes: [[1], [0, 1]],
		widths: [200, 200],
		x: 226.6666667,
	},
];
describe('native legacy Word footnote columns', () => {
	for (const mode of [12, 14])
		it.each(variants)(
			`mode ${mode}: $name preserves native column flow and separators`,
			async ({ name, counts, indexes, widths, x }) => {
				const loaded = await loadDocx(
					new Uint8Array(
						readFileSync(
							new URL(
								'../__fixtures__/footnote-columns/mode' + mode + '-' + name + '.docx',
								import.meta.url,
							),
						),
					),
				);
				const result = layoutDocumentModel(loaded.model, measurer);
				expect(result.pages).toHaveLength(counts.length);
				expect(
					result.pages.map((page) =>
						page.footnotes?.map((note) => note.paragraphs.flatMap((p) => p.lines).length),
					),
				).toEqual(counts);
				expect(result.pages.map((page) => page.footnotes?.map((note) => note.columnIndex))).toEqual(
					indexes,
				);
				if (name.startsWith('left')) {
					const positions = result.pages.flatMap((page, pageIndex) =>
						page.columns.flatMap((column, columnIndex) =>
							column.blocks
								.filter(
									(block) =>
										block.kind === 'paragraph' &&
										block.lines
											.flatMap((line) => line.fragments.map((fragment) => fragment.text))
											.join('')
											.includes('Right body'),
								)
								.map((block) => ({ pageIndex, columnIndex, y: block.yPx })),
						),
					);
					expect(positions).toEqual([
						{
							pageIndex: name.endsWith('sixty') ? 0 : 1,
							columnIndex: name.endsWith('sixty') ? 1 : 0,
							y: 0,
						},
					]);
				}
				const sourceIds: string[] = [];
				for (const [pageIndex, page] of result.pages.entries()) {
					expect(page.footnoteSeparator).toBeUndefined();
					expect(page.footnoteColumnAreas).toHaveLength(indexes[pageIndex]!.length);
					for (const area of page.footnoteColumnAreas!) {
						expect(area.xPx).toBeCloseTo(area.columnIndex === 0 ? 0 : x, 5);
						expect(area.widthPx).toBeCloseTo(widths[area.columnIndex]!, 5);
						expect(area.yPx + area.heightPx).toBeCloseTo(page.footnoteHeightPx!, 5);
						const fragments = area.separator!.paragraphs.flatMap((p) =>
							p.lines.flatMap((line) => line.fragments),
						);
						const text = fragments.map((fragment) => fragment.text).join('');
						if (name.includes('custom') || name.startsWith('explicit'))
							expect(text).toContain(area.continued ? 'Continued' : 'Sources');
						const rule = fragments.find((fragment) => fragment.object?.noteSeparator)!;
						expect(rule.xPx + rule.widthPx).toBeLessThanOrEqual(area.widthPx + 1e-6);
					}
					for (const note of page.footnotes ?? [])
						for (const paragraph of note.paragraphs)
							for (const line of paragraph.lines)
								sourceIds.push(
									...(line.fragments
										.map((fragment) => fragment.text)
										.join('')
										.match(/N0L\d{2}/g) ?? []),
								);
				}
				const total = counts.flat().reduce((sum, count) => sum + count, 0);
				expect(sourceIds).toEqual(
					Array.from({ length: total }, (_, index) => 'N0L' + String(index).padStart(2, '0')),
				);
			},
		);
	it('reflows remaining source tokens into a narrower continuation column without loss', async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL(
						'../__fixtures__/footnote-columns/mode14-unequal-custom-eighty.docx',
						import.meta.url,
					),
				),
			),
		);
		const model = structuredClone(loaded.model),
			note = model.footnotes![0]!;
		const paragraph = note.blocks.find((block) => block.type === 'paragraph')!;
		if (paragraph.type !== 'paragraph') throw new Error('Missing note paragraph');
		const words = Array.from(
			{ length: 400 },
			(_, index) => 'word' + String(index).padStart(3, '0'),
		);
		paragraph.runs = [{ text: words.join(' ') }];
		note.blocks = [paragraph];
		const result = layoutDocumentModel(model, measurer);
		const notes = result.pages.flatMap((page) => page.footnotes ?? []);
		const text = notes
			.flatMap((note) =>
				note.paragraphs.flatMap((p) =>
					p.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)),
				),
			)
			.join(' ');
		expect(text.match(/word\d{3}/g)).toEqual(words);
		expect(notes.some((note) => note.widthPx! < 150 && note.continued)).toBe(true);
	});

	it.each([undefined, 11, 13, 16])(
		'preserves independent settings with an explicit fallback for unmeasured mode %s',
		async (mode) => {
			const loaded = await loadDocx(
				new Uint8Array(
					readFileSync(
						new URL('../__fixtures__/footnote-columns/two-two-twenty.docx', import.meta.url),
					),
				),
			);
			const model = structuredClone(loaded.model);
			if (mode === undefined) delete model.compatibilityMode;
			else model.compatibilityMode = mode;
			const result = layoutDocumentModel(model, measurer);
			expect(model.sections![0]!.footnoteColumns).toBe(2);
			expect(result.pages.every((page) => page.footnoteColumnAreas === undefined)).toBe(true);
			expect(
				result.pages
					.flatMap((page) => page.footnotes ?? [])
					.every((note) => note.columnIndex === undefined),
			).toBe(true);
			expect(result.approximations).toContain(
				'Footnote column settings are preserved, but this compatibility mode has not been verified against native Word; Print Layout uses the existing page-wide fallback.',
			);
		},
	);
});
