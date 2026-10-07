import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow';
import { createFakeMeasurer } from './measure';
import type { LayoutTableRow } from './input';
import type { LayoutTableBox } from './result';

const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 20 };
function row(id: string, isHeader: boolean, text = id): LayoutTableRow {
	return { isHeader, cells: [{ paragraphs: [{ kind: 'paragraph', id, runs: [{ text }] }] }] };
}
function paginate(rows: LayoutTableRow[]) {
	return layoutSections(
		{
			sections: [
				{
					page: {
						widthPx: 200,
						heightPx: 60,
						marginTopPx: 0,
						marginRightPx: 0,
						marginBottomPx: 0,
						marginLeftPx: 0,
					},
					blocks: [{ kind: 'table', id: 'table', rows }],
				},
			],
		},
		measurer,
	);
}
function fragments(rows: LayoutTableRow[]) {
	return paginate(rows).pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.filter((block): block is LayoutTableBox => block.kind === 'table'),
		),
	);
}
const repeatedIds = (rows: LayoutTableRow[]) =>
	fragments(rows).flatMap((fragment) =>
		fragment.rows
			.filter((r) => r.repeated)
			.flatMap((r) => r.cells.flatMap((cell) => cell.map((paragraph) => paragraph.blockId))),
	);

describe('Word table header pagination', () => {
	it('repeats only the initial contiguous header group', () => {
		expect(
			repeatedIds([
				row('h1', true),
				row('h2', true),
				row('body', false),
				row('ignored', true),
				row('end', false),
			]),
		).toEqual(['h1', 'h2', 'h1', 'h2']);
	});

	it('does not repeat a marked row when the first row is not a header', () => {
		expect(
			repeatedIds([
				row('body', false),
				row('ignored', true),
				row('body2', false),
				row('body3', false),
			]),
		).toEqual([]);
	});

	it('allows a noncontiguous header flag to split as an ordinary body row', () => {
		const result = fragments([row('body', false), row('ignored', true, 'a\nb\nc\nd')]);
		expect(result).toHaveLength(2);
		expect(result.flatMap((f) => f.rows).some((r) => r.repeated)).toBe(false);
		const ignored = result
			.flatMap((f) => f.rows.flatMap((r) => r.cells.flatMap((c) => c)))
			.filter((p) => p.blockId === 'ignored');
		expect(ignored).toHaveLength(2);
		expect(ignored.map((p) => p.lines.length)).toEqual([2, 2]);
	});
});

describe('table flow progress', () => {
	it('preserves body text when repeated headers leave no room for a line', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page: {
							widthPx: 200,
							heightPx: 20,
							marginTopPx: 0,
							marginBottomPx: 0,
							marginLeftPx: 0,
							marginRightPx: 0,
						},
						blocks: [
							{ kind: 'table', id: 'table', rows: [row('header', true), row('body', false)] },
						],
					},
				],
			},
			measurer,
		);
		const body = result.pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'table'
						? block.rows.flatMap((row) =>
								row.cells.flatMap((cell) =>
									cell.filter((paragraph) => paragraph.blockId === 'body'),
								),
							)
						: [],
				),
			),
		);
		expect(body).toHaveLength(1);
		expect(body[0]!.lines[0]!.fragments.map((fragment) => fragment.text).join('')).toBe('body');
		expect(result.approximations).toContain(
			'A table row cannot fit a line in an empty page/column and was placed without splitting.',
		);
	});
	it('reports and preserves an oversized splittable row when no line fits an empty column', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page: {
							widthPx: 200,
							heightPx: 10,
							marginTopPx: 0,
							marginBottomPx: 0,
							marginLeftPx: 0,
							marginRightPx: 0,
						},
						blocks: [
							{
								kind: 'table',
								id: 'table',
								rows: [
									{
										cells: [
											{
												paragraphs: [
													{ kind: 'paragraph', id: 'text', runs: [{ text: 'Oversized line' }] },
												],
											},
										],
									},
								],
							},
						],
					},
				],
			},
			createFakeMeasurer(),
		);
		expect(result.pages).toHaveLength(1);
		const table = result.pages[0]!.columns[0]!.blocks[0]!;
		if (table.kind !== 'table') throw new Error('Expected table');
		expect(
			table.rows[0]!.cells[0]![0]!.lines[0]!.fragments.map((fragment) => fragment.text).join(''),
		).toBe('Oversized line');
		expect(result.approximations).toContain(
			'A table row cannot fit a line in an empty page/column and was placed without splitting.',
		);
	});
});
