import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow';
import type { LayoutParagraph, LayoutTable } from './input';
import type { LayoutResult } from './result';

const page = {
	widthPx: 200,
	heightPx: 100,
	marginTopPx: 0,
	marginBottomPx: 0,
	marginLeftPx: 0,
	marginRightPx: 0,
};
const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const paragraph = (id: string, count = 1): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	runs: [{ text: Array.from({ length: count }, (_, index) => `${id}${index}`).join('\n') }],
});
const reference = (count: number): LayoutParagraph => ({
	...paragraph('body'),
	footnotes: [{ id: '1', paragraphs: [paragraph('note', count)] }],
});
const texts = (result: LayoutResult) =>
	result.pages.flatMap((sheet) =>
		(sheet.footnotes ?? []).flatMap((note) =>
			note.paragraphs.flatMap((box) =>
				box.lines.map((line) => line.fragments.map((fragment) => fragment.text).join('')),
			),
		),
	);

describe('long footnote continuation', () => {
	it('keeps the reference with the start of its note and drains remaining lines after the body ends', () => {
		const result = layoutSections({ sections: [{ page, blocks: [reference(10)] }] }, measurer);
		expect(result.pages).toHaveLength(3);
		expect(result.pages.map((sheet) => sheet.footnotes?.[0]?.paragraphs[0]?.lines.length)).toEqual([
			3, 4, 3,
		]);
		expect(result.pages.map((sheet) => sheet.footnotes?.[0]?.continued ?? false)).toEqual([
			false,
			true,
			true,
		]);
		expect(result.pages.flatMap((sheet) => sheet.footnotes ?? []).map((note) => note.id)).toEqual([
			'1',
			'1',
			'1',
		]);
		expect(texts(result)).toEqual(Array.from({ length: 10 }, (_, index) => `note${index}`));
		expect(result.approximations).toEqual([]);
		for (const sheet of result.pages) {
			const height = (sheet.footnotes ?? []).reduce((sum, note) => sum + note.heightPx, 0) + 12;
			for (const column of sheet.columns)
				for (const block of column.blocks)
					expect(block.yPx + block.heightPx).toBeLessThanOrEqual(100 - height);
		}
	});

	it('reserves continued notes before placing later body paragraphs', () => {
		const result = layoutSections(
			{ sections: [{ page, blocks: [reference(10), paragraph('next', 3)] }] },
			measurer,
		);
		expect(texts(result)).toEqual(Array.from({ length: 10 }, (_, index) => `note${index}`));
		const body = result.pages.flatMap((sheet) => sheet.columns.flatMap((column) => column.blocks));
		expect(
			body
				.filter((block) => block.blockId === 'next')
				.flatMap((block) =>
					block.kind === 'paragraph' ? block.lines.map((line) => line.sourceStart) : [],
				),
		).toHaveLength(3);
	});

	it('continues paragraphs without doubling their offsets or losing formatting/source positions', () => {
		const ref = reference(1);
		ref.footnotes![0]!.paragraphs = [
			paragraph('a', 2),
			{ ...paragraph('b', 4), runs: [{ text: 'b0\nb1\nb2\nb3', bold: true }] },
		];
		const result = layoutSections({ sections: [{ page, blocks: [ref] }] }, measurer);
		expect(texts(result)).toEqual(['a0', 'a1', 'b0', 'b1', 'b2', 'b3']);
		const first = result.pages[0]!.footnotes![0]!;
		expect(first.paragraphs[1]!.yPx).toBe(40);
		expect(first.paragraphs[1]!.lines[0]!.yPx).toBe(0);
		const second = result.pages[1]!.footnotes![0]!;
		expect(second.paragraphs[0]!.lines[0]!.sourceStart).toBe(3);
		expect(second.paragraphs[0]!.lines[0]!.fragments[0]!.bold).toBe(true);
	});

	it('retains reference order for multiple notes when the first continues', () => {
		const ref = reference(6);
		ref.footnotes!.push({ id: '2', paragraphs: [paragraph('other', 2)] });
		const result = layoutSections({ sections: [{ page, blocks: [ref] }] }, measurer);
		expect(result.pages.flatMap((sheet) => sheet.footnotes ?? []).map((note) => note.id)).toEqual([
			'1',
			'1',
			'2',
			'2',
		]);
		expect(texts(result)).toEqual([
			'note0',
			'note1',
			'note2',
			'note3',
			'note4',
			'note5',
			'other0',
			'other1',
		]);
	});

	it('reports indivisible oversized note lines and still makes finite progress', () => {
		const ref = reference(1);
		ref.footnotes![0]!.paragraphs[0]!.runs = [
			{
				text: '',
				object: { partName: 'image', contentType: 'image/png', widthPx: 20, heightPx: 150 },
			},
		];
		const result = layoutSections({ sections: [{ page, blocks: [ref] }] }, measurer);
		expect(result.pages).toHaveLength(1);
		expect(result.approximations).toContain(
			'A footnote line was taller than its available page area and was placed without splitting.',
		);
	});

	it.each([false, true])(
		'moves later table content past full continuation pages (cantSplit=%s)',
		(cantSplit) => {
			const table: LayoutTable = {
				kind: 'table',
				id: 'table',
				rows: [{ cantSplit, cells: [{ paragraphs: [paragraph('cell', 2)] }] }],
			};
			const result = layoutSections(
				{ sections: [{ page, blocks: [reference(10), table] }] },
				measurer,
			);
			for (const sheet of result.pages) {
				const reserved = sheet.footnotes?.length
					? 12 + sheet.footnotes.reduce((sum, note) => sum + note.heightPx, 0)
					: 0;
				for (const column of sheet.columns)
					for (const block of column.blocks)
						expect(block.yPx + block.heightPx).toBeLessThanOrEqual(100 - reserved);
			}
			expect(texts(result)).toHaveLength(10);
		},
	);

	it('moves a later-column reference before its footnote area could overlap an earlier column', () => {
		const ref = reference(1);
		const filler = paragraph('filler', 4);
		filler.runs.push({ text: '', breakAfter: 'column' });
		const result = layoutSections(
			{ sections: [{ page, columns: { count: 2, gapPx: 0 }, blocks: [filler, ref] }] },
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(result.pages[0]!.footnotes).toBeUndefined();
		expect(result.pages[1]!.footnotes![0]!.id).toBe('1');
	});
});
