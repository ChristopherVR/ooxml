import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow';
import { FOOTNOTE_SEPARATOR_PX } from './page-cursor';
import type { LayoutParagraph, LayoutSection } from './input';
import { at } from './test-helpers';

const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const page = {
	widthPx: 200,
	heightPx: 140,
	marginTopPx: 20,
	marginRightPx: 0,
	marginBottomPx: 20,
	marginLeftPx: 0,
};
const paragraph = (id: string): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	runs: [{ text: 'Claim' }],
});
function noted(id: string, noteLines = 1): LayoutParagraph {
	return {
		...paragraph(id),
		footnotes: [
			{
				id: `note-${id}`,
				paragraphs: [
					{
						...paragraph(`note-body-${id}`),
						runs: [{ text: Array.from({ length: noteLines }, () => 'Source').join('\n') }],
					},
				],
			},
		],
	};
}

describe('section pagination with footnotes', () => {
	it.each(['center', 'bottom'] as const)(
		'keeps %s-aligned body clear of notes and separator',
		(verticalAlign) => {
			const result = layoutSections(
				{ sections: [{ page, blocks: [noted('p')], verticalAlign }] },
				measurer,
			);
			const sheet = at(result.pages, 0);
			const body = at(at(sheet.columns, 0).blocks, 0);
			const note = at(sheet.footnotes, 0);
			const bodyArea = 100 - FOOTNOTE_SEPARATOR_PX - note.heightPx;
			expect(body.yPx).toBe((bodyArea - body.heightPx) / (verticalAlign === 'center' ? 2 : 1));
			expect(body.yPx + body.heightPx).toBeLessThanOrEqual(bodyArea);
			expect(note.yPx).toBe(0);
		},
	);

	it('reserves all notes but only one separator when bottom-aligning', () => {
		const result = layoutSections(
			{ sections: [{ page, blocks: [noted('a'), noted('b')], verticalAlign: 'bottom' }] },
			measurer,
		);
		const sheet = at(result.pages, 0);
		const blocks = at(sheet.columns, 0).blocks;
		expect(sheet.footnotes?.map((note) => note.yPx)).toEqual([0, 20]);
		expect(blocks.map((block) => block.yPx)).toEqual([8, 28]);
		expect(at(blocks, 1).yPx + at(blocks, 1).heightPx).toBe(48);
	});

	it('reserves notes only on the page containing their reference', () => {
		const next = noted('next', 2);
		next.pageBreakBefore = true;
		const result = layoutSections(
			{ sections: [{ page, blocks: [paragraph('first'), next], verticalAlign: 'bottom' }] },
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(at(at(at(result.pages, 0).columns, 0).blocks, 0).yPx).toBe(80);
		expect(at(result.pages, 0).footnotes).toBeUndefined();
		expect(at(at(at(result.pages, 1).columns, 0).blocks, 0).yPx).toBe(28);
	});

	it('keeps the existing top alignment when notes are present', () => {
		const result = layoutSections(
			{ sections: [{ page, blocks: [noted('p')], verticalAlign: 'top' }] },
			measurer,
		);
		expect(at(at(at(result.pages, 0).columns, 0).blocks, 0).yPx).toBe(0);
	});

	it.each(['oddPage', 'evenPage'] as const)(
		'does not repeat notes on an inserted %s parity blank',
		(sectionBreak) => {
			const sections: LayoutSection[] = [{ page, blocks: [noted('a')] }];
			if (sectionBreak === 'evenPage') sections.push({ page, blocks: [noted('b')] });
			sections.push({ page, blocks: [paragraph('next')], break: sectionBreak });
			const result = layoutSections({ sections }, measurer);
			const blank = at(result.pages, result.pages.length - 2);
			const preceding = at(result.pages, result.pages.length - 3);
			expect(blank.columns.every((column) => column.blocks.length === 0)).toBe(true);
			expect(blank.footnotes).toBeUndefined();
			expect(blank.floats).toBeUndefined();
			expect(preceding.footnotes).toHaveLength(1);
			expect(blank.sectionIndex).toBe(preceding.sectionIndex);
			expect(blank.pageInSection).toBe(preceding.pageInSection + 1);
		},
	);
});
