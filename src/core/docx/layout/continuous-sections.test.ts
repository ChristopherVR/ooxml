import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, type TextMeasurer } from './measure';
import { layoutDocumentModel } from './layout';
import { layoutSections } from './page-flow';
import type { LayoutPageGeometry, LayoutParagraph } from './input';

const fixture = (name: string) =>
	new URL(`./fixtures/continuous-sections/${name}`, import.meta.url);
const page: LayoutPageGeometry = {
	widthPx: 200,
	heightPx: 100,
	marginTopPx: 0,
	marginBottomPx: 0,
	marginLeftPx: 0,
	marginRightPx: 0,
};
const measurer: TextMeasurer = { widthOf: (text) => text.length * 5, lineHeightOf: () => 20 };
const paragraph = (id: string, lines = 1): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	runs: [{ text: Array.from({ length: lines }, () => 'line').join('\n') }],
});

describe('native Word continuous section references', () => {
	it.each([
		'same',
		'left-margin',
		'top-margin',
		'columns',
		'page-size',
		'orientation',
		'balanced-columns',
		'balanced-odd',
		'balanced-overflow',
		'balanced-keep',
		'balanced-unequal',
		'balanced-unequal-odd',
		'balanced-unequal-overflow',
		'balanced-unequal-reverse',
	])('matches paragraph page and origin in the %s reference', async (name) => {
		const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
			cases: {
				name: string;
				pages: number;
				positions: { text: string; page: number; xPt: number; yPt: number }[];
			}[];
		};
		const reference = evidence.cases.find((entry) => entry.name === name)!;
		const loaded = await loadDocx(new Uint8Array(await readFile(fixture(`${name}.docx`))));
		const result = layoutDocumentModel(loaded.model, createFakeMeasurer());
		expect(result.pages).toHaveLength(reference.pages);
		for (const position of reference.positions.filter((entry) => entry.text)) {
			const source = loaded.model.blocks.find(
				(block) =>
					block.type === 'paragraph' &&
					block.runs.map((run) => run.text).join('') === position.text,
			)!;
			const placements = result.pages.flatMap((sheet) =>
				sheet.columns.flatMap((column) =>
					column.blocks
						.filter((block) => block.blockId === source.id)
						.map((block) => ({
							page: sheet.index + 1,
							xPt: (sheet.marginLeftPx + column.xPx) * 0.75,
							yPt: (sheet.marginTopPx + block.yPx) * 0.75,
						})),
				),
			);
			expect(placements).toEqual([{ page: position.page, xPt: position.xPt, yPt: position.yPt }]);
		}
		expect(result.approximations).not.toContain(
			'Continuous section breaks are rendered as page breaks; changing page size, margins or column count without starting a new page is not modeled.',
		);
	});
});

describe('continuous section flow', () => {
	it('keeps a whole paragraph together while balancing columns', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page,
						columns: { count: 2, gapPx: 10 },
						blocks: [{ ...paragraph('whole', 3), keepLines: true }, paragraph('tail')],
					},
					{ page, break: 'continuous', blocks: [paragraph('after')] },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(1);
		const boxes = result.pages[0]!.columns.flatMap((column) => column.blocks);
		expect(boxes.filter((box) => box.blockId === 'whole')).toHaveLength(1);
		expect(boxes.find((box) => box.blockId === 'after')!.yPx).toBe(60);
		expect(result.approximations).toEqual([]);
	});
	it('continues into new columns at the section band top and uses new margins on overflow', () => {
		const result = layoutSections(
			{
				sections: [
					{ page, blocks: [paragraph('before')] },
					{
						page: { ...page, marginTopPx: 10, marginLeftPx: 10 },
						break: 'continuous',
						columns: { count: 2, gapPx: 10 },
						blocks: [paragraph('after', 10)],
					},
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(result.pages[0]).toMatchObject({ sectionIndex: 0, sectionIndices: [0, 1] });
		const columns = result.pages[0]!.columns;
		expect(columns[1]).toMatchObject({ sectionIndex: 1, xPx: 10, startYPx: 20 });
		expect(columns[2]!.blocks[0]!.yPx).toBe(20);
		expect(result.pages[1]).toMatchObject({
			sectionIndex: 1,
			pageInSection: 1,
			marginTopPx: 10,
			marginLeftPx: 10,
		});
		expect(result.pages[1]!.columns[0]!.blocks[0]!.yPx).toBe(0);
	});

	it('honors page-break-before at the first line of a shared band', () => {
		const result = layoutSections(
			{
				sections: [
					{ page, blocks: [paragraph('before')] },
					{ page, break: 'continuous', blocks: [{ ...paragraph('after'), pageBreakBefore: true }] },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(result.pages[1]).toMatchObject({ sectionIndex: 1, pageInSection: 0 });
	});

	it('keeps consecutive continuous sections in physical flow order', () => {
		const result = layoutSections(
			{
				sections: [
					{ page, blocks: [paragraph('a')] },
					{ page, break: 'continuous', blocks: [paragraph('b')] },
					{ page, break: 'continuous', blocks: [paragraph('c')] },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(1);
		expect(result.pages[0]!.sectionIndices).toEqual([0, 1, 2]);
		expect(
			result.pages[0]!.columns.flatMap((column) => column.blocks.map((block) => block.yPx)),
		).toEqual([0, 20, 40]);
	});

	it('reports the remaining table-column balancing approximation', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page,
						columns: { count: 2, gapPx: 10 },
						blocks: [{ kind: 'table', id: 'table', rows: [] }],
					},
					{ page, break: 'continuous', blocks: [paragraph('after')] },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(result.approximations.some((note) => note.includes('balancing those layouts'))).toBe(
			true,
		);
	});
});
