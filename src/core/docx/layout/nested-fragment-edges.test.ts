import { expect, it } from 'vitest';
import { nestedFragmentEdges, reflowNestedFragments } from './nested-fragment-edges';
import type { LayoutTable, LayoutTableRow } from './input';
const edge = (widthPx: number) => ({ widthPx, style: 'solid' as const, color: '#FF0000' });
const source = (): LayoutTable => ({
	kind: 'table',
	id: 't',
	widthPx: 200,
	rows: [0, 1].map((index) => ({
		cells: [0, 1].map((column) => ({
			xPx: column * 100,
			widthPx: 100,
			paragraphs: [
				{ kind: 'paragraph' as const, id: index + ':' + column, runs: [{ text: 'Line' }] },
			],
			borders: {
				top: edge(index ? 2 : 3),
				bottom: edge(index ? 1 : 2),
				left: edge(1),
				right: edge(1),
			},
		})),
	})),
});
it('retains original outer edges rather than an interior row edge', () => {
	const table = source();
	expect(nestedFragmentEdges(table)?.top[0]?.widthPx).toBe(3);
	expect(nestedFragmentEdges(table)?.bottom[0]?.widthPx).toBe(1);
});
it.each([
	'note',
	'keep',
	'fixed',
	'header',
	'merge',
	'nested',
	'center',
	'three',
	'outer-note',
	'auto-width',
	'multiple-tables',
])('keeps %s outside measured reflow admission', (kind) => {
	const table = source(),
		cell = table.rows[0]!.cells[0]!,
		p = cell.paragraphs[0]!;
	if (kind === 'auto-width') delete table.widthPx;
	if (kind === 'note') p.runs[0]!.footnoteReferenceId = '1';
	if (kind === 'keep') p.keepNext = true;
	if (kind === 'fixed') table.rows[0]!.heightPx = 50;
	if (kind === 'header') table.rows[0]!.isHeader = true;
	if (kind === 'merge') delete cell.borders!.bottom;
	if (kind === 'nested') cell.blocks = [source()];
	if (kind === 'center') cell.verticalAlign = 'center';
	if (kind === 'three') table.rows.push(structuredClone(table.rows[0]!));
	const outer: LayoutTableRow = {
		cells: [
			{
				paragraphs: [],
				blocks: [
					table,
					{
						kind: 'paragraph' as const,
						id: 'after',
						runs: [
							{ text: 'After', ...(kind === 'outer-note' ? { footnoteReferenceId: '1' } : {}) },
						],
					},
				],
			},
			{ paragraphs: [{ kind: 'paragraph' as const, id: 'other', runs: [{ text: 'Other' }] }] },
		],
	};
	if (kind === 'multiple-tables') outer.cells[1]!.paragraphs = [];
	if (kind === 'multiple-tables') outer.cells[1]!.blocks = [source()];
	expect(reflowNestedFragments(outer)).toBe(false);
});
