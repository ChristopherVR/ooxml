import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow';
import { layoutDocumentModel } from './layout';
import { adaptDocumentModel } from './adapter';
import type { DocumentModel } from '../model';
import type { LayoutParagraph } from './input';
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
const note = (id: string, lines = 1) => ({
	id,
	paragraphs: [
		{
			kind: 'paragraph' as const,
			id: `note-${id}`,
			runs: [{ text: Array.from({ length: lines }, (_, index) => `N${index}`).join('\n') }],
		},
	],
});
const reference = (before: string, after = ''): LayoutParagraph => ({
	kind: 'paragraph',
	id: 'body',
	widowControl: false,
	runs: [{ text: before }, { text: '1', footnoteReferenceId: '1' }, { text: after }],
	footnotes: [note('1')],
});
const referencePage = (result: LayoutResult, id = '1') =>
	result.pages.findIndex((sheet) =>
		sheet.columns.some((column) =>
			column.blocks.some(
				(block) =>
					block.kind === 'paragraph' &&
					block.lines.some((line) => line.fragments.some((fragment) => fragment.text === id)),
			),
		),
	);

describe('footnotes attached to their actual reference line', () => {
	it('keeps a short kept paragraph together with a later reference and note on a fresh page', () => {
		const block = reference('A\nClaim');
		block.keepLines = true;
		const filler: LayoutParagraph = { kind: 'paragraph', id: 'filler', runs: [{ text: 'A\nB' }] };
		const result = layoutSections({ sections: [{ page, blocks: [filler, block] }] }, measurer);
		expect(result.pages[0]!.columns[0]!.blocks.map((item) => item.blockId)).toEqual(['filler']);
		expect(result.pages[1]!.columns[0]!.blocks[0]!.blockId).toBe('body');
		expect(result.pages[1]!.footnotes![0]!.id).toBe('1');
	});
	it('places a late-paragraph reference and its note on the second page, leaving earlier lines alone', () => {
		const result = layoutSections(
			{ sections: [{ page, blocks: [reference('A\nB\nC\nD\nE\nClaim')] }] },
			measurer,
		);
		expect(referencePage(result)).toBe(1);
		expect(result.pages[0]!.footnotes).toBeUndefined();
		expect(result.pages[0]!.columns[0]!.blocks[0]!.heightPx).toBe(100);
		expect(result.pages[1]!.footnotes![0]!.id).toBe('1');
	});

	it('stops before a reference line when that line and its short note cannot fit', () => {
		const result = layoutSections(
			{ sections: [{ page, blocks: [reference('A\nB\nC\nClaim')] }] },
			measurer,
		);
		expect(referencePage(result)).toBe(1);
		expect(result.pages[0]!.footnotes).toBeUndefined();
		expect(result.pages[0]!.columns[0]!.blocks[0]!.heightPx).toBe(60);
	});

	it('assigns multiple notes on different pages of one paragraph without duplication', () => {
		const block = reference('Claim');
		block.runs.push({ text: '\nA\nB\nC\nClaim' }, { text: '2', footnoteReferenceId: '2' });
		block.footnotes!.push(note('2'));
		const result = layoutSections({ sections: [{ page, blocks: [block] }] }, measurer);
		expect(result.pages[referencePage(result)]!.footnotes!.map((item) => item.id)).toEqual(['1']);
		expect(result.pages[referencePage(result, '2')]!.footnotes!.map((item) => item.id)).toEqual([
			'2',
		]);
	});

	it('keeps the note with a reference line moved by widow control', () => {
		const block = reference('A\nB\nClaim', '\nD');
		block.widowControl = true;
		const result = layoutSections({ sections: [{ page, blocks: [block] }] }, measurer);
		expect(referencePage(result)).toBe(1);
		expect(result.pages[0]!.footnotes).toBeUndefined();
	});

	it('starts a long note beside a late reference and continues it on following physical pages', () => {
		const block = reference('A\nB\nC\nD\nE\nClaim');
		block.footnotes = [note('1', 8)];
		const result = layoutSections({ sections: [{ page, blocks: [block] }] }, measurer);
		expect(result.pages[0]!.footnotes).toBeUndefined();
		expect(result.pages[referencePage(result)]!.footnotes![0]!.continued).toBeUndefined();
		expect(result.pages.slice(2).every((sheet) => sheet.footnotes![0]!.continued)).toBe(true);
	});

	it('uses the adapted reference run even with generated list labels, instead of its source text offset', () => {
		const model: DocumentModel = {
			page: {
				width: 200,
				height: 100,
				marginTop: 0,
				marginBottom: 0,
				marginLeft: 0,
				marginRight: 0,
			},
			warnings: [],
			blocks: [
				{
					type: 'paragraph',
					id: 'p',
					widowControl: false,
					runs: [
						{ text: 'A\nB\nC\nD\nE\nClaim' },
						{ text: '', noteReference: { kind: 'footnote', id: '1' } },
					],
				},
			],
			footnotes: [
				{ id: '1', blocks: [{ type: 'paragraph', id: 'n', runs: [{ text: 'Source' }] }] },
			],
		};
		const input = adaptDocumentModel(model);
		const block = input.sections[0]!.blocks[0] as LayoutParagraph;
		expect(block.runs[1]!.footnoteReferenceId).toBe('1');
		block.runs.unshift({ text: '1. ', synthetic: true });
		const result = layoutSections(input, measurer);
		expect(result.pages[0]!.footnotes).toBeUndefined();
		expect(result.pages[referencePage(result)]!.footnotes![0]!.id).toBe('1');
		expect(layoutDocumentModel(model, measurer).pages[0]!.footnotes).toBeUndefined();
	});
});
