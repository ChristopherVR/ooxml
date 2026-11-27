import { describe, expect, it } from 'vitest';
import type { LayoutParagraph, LayoutTable } from './input';
import { layoutSections } from './page-flow';

const measure = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
const paragraph = (id: string, widowControl: boolean): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	widowControl,
	runs: [{ text: `${id}a\n${id}b` }],
});
const table = (widowControl: boolean): LayoutTable => ({
	kind: 'table',
	id: 'table',
	widthPx: 200,
	rows: ['r1', 'r2', 'r3'].map((id) => ({
		cells: [
			{ widthPx: 100, xPx: 0, paragraphs: [paragraph(`${id}x`, widowControl)] },
			{
				widthPx: 100,
				xPx: 100,
				paragraphs: [{ kind: 'paragraph', id: `${id}y`, runs: [{ text: id }] }],
			},
		],
	})),
});
/** Each page's rows, as the lines of their first cell. */
function pages(widowControl: boolean) {
	const result = layoutSections(
		{
			sections: [
				{
					page: {
						widthPx: 200,
						heightPx: 80,
						marginTopPx: 0,
						marginRightPx: 0,
						marginBottomPx: 0,
						marginLeftPx: 0,
					},
					blocks: [table(widowControl)],
				},
			],
		},
		measure,
	);
	return result.pages.map((page) =>
		page.columns[0]!.blocks.flatMap((block) =>
			block.kind === 'table'
				? block.rows.map((row) => row.cells.map((cell) => cell.flatMap((box) => box.lines.length)))
				: [],
		),
	);
}

// Word 16 (`pdf-export/__fixtures__/word-reference/table-header*.docx`): with widow control the
// row whose two-line paragraph has room for one line moves to the next page whole, one-line cells
// included; with it off the row splits after the first line.
describe('table rows and widow control', () => {
	it('moves a row whole rather than leave one line of a cell paragraph', () => {
		expect(pages(true)).toEqual([
			[
				[[2], [1]],
				[[2], [1]],
			],
			[[[2], [1]]],
		]);
	});
	it('splits the row after the first line with widow control off', () => {
		expect(pages(false)).toEqual([
			[
				[[2], [1]],
				[[2], [1]],
				[[1], [1]],
			],
			[[[1], []]],
		]);
	});
});
