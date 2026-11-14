import { describe, expect, it } from 'vitest';
import { layoutDocument } from './layout';
import { FLOATING_TABLE_IN_FLOW_NOTE, FLOATING_TABLE_NOTE } from './flow-floating-table';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutParagraph,
	LayoutTable,
	LayoutTableFloating,
} from './input';
import type { TextMeasurer } from './measure';
import type { LayoutTableBox } from './result';

const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const page = {
	widthPx: 800,
	heightPx: 1000,
	marginTopPx: 100,
	marginRightPx: 100,
	marginBottomPx: 100,
	marginLeftPx: 100,
};
const input = (blocks: LayoutBlock[], more: LayoutDocumentInput['sections'] = []) => ({
	sections: [{ page, blocks }, ...more],
});
const paragraph = (id: string, text = 'text', extra: Partial<LayoutParagraph> = {}) => ({
	kind: 'paragraph' as const,
	id,
	runs: [{ text }],
	...extra,
});
const frame: LayoutTableFloating = {
	relativeFromH: 'margin',
	offsetXPx: 0,
	relativeFromV: 'paragraph',
	offsetYPx: 0,
	distance: { topPx: 0, rightPx: 10, bottomPx: 0, leftPx: 0 },
};
function floatingTable(rows: number, extra: Partial<LayoutTable> = {}, cell = {}): LayoutTable {
	return {
		kind: 'table',
		id: 't',
		widthPx: 200,
		rows: Array.from({ length: rows }, (_, index) => ({
			cells: [{ paragraphs: [paragraph(`c${index}`, 'cell')], widthPx: 200, ...cell }],
		})),
		floating: frame,
		...extra,
	};
}
const tables = (result: ReturnType<typeof layoutDocument>) =>
	result.pages.flatMap((p) =>
		p.columns.flatMap((column) =>
			column.blocks.filter((block): block is LayoutTableBox => block.kind === 'table'),
		),
	);
const words = Array.from({ length: 80 }, () => 'word').join(' ');

describe('floating table placement', () => {
	it('takes no room from the flow and wraps the next paragraph beside it', () => {
		const result = layoutDocument(input([floatingTable(3), paragraph('after', words)]), measurer);
		const [table] = tables(result);
		expect(table?.floating).toBeDefined();
		const after = result.pages[0]!.columns[0]!.blocks.find((block) => block.blockId === 'after');
		expect(after?.yPx).toBe(table?.yPx);
		const line = after?.kind === 'paragraph' ? after.lines[0] : undefined;
		// 200 px table plus its 10 px right distance.
		expect(line?.fragments[0]?.xPx).toBeGreaterThanOrEqual(210);
		expect(result.approximations).toContain(FLOATING_TABLE_NOTE);
	});

	it('lays a table taller than the page out in the flow, split across pages', () => {
		const result = layoutDocument(input([floatingTable(60), paragraph('after')]), measurer);
		const boxes = tables(result);
		expect(boxes.length).toBeGreaterThan(1);
		expect(boxes.every((box) => !box.floating)).toBe(true);
		expect(result.approximations).toContain(FLOATING_TABLE_IN_FLOW_NOTE);
	});

	it('keeps tables with footnote references or vertical merges in the flow', () => {
		const note = { id: 'n1', paragraphs: [paragraph('n1p', 'note')] };
		const withNote: LayoutTable = {
			...floatingTable(1),
			rows: [
				{ cells: [{ paragraphs: [paragraph('c', 'cell', { footnotes: [note] })], widthPx: 200 }] },
			],
		};
		const noted = layoutDocument(input([withNote, paragraph('after')]), measurer);
		expect(tables(noted)[0]?.floating).toBeUndefined();
		expect(noted.pages[0]!.footnotes?.length).toBe(1);
		const merged = layoutDocument(
			input([floatingTable(2, {}, { verticalMerge: 'restart' }), paragraph('after')]),
			measurer,
		);
		expect(tables(merged)[0]?.floating).toBeUndefined();
	});

	it('does not count a low page-anchored table as body content for footnotes', () => {
		const low: LayoutTable = floatingTable(2, {
			floating: { ...frame, relativeFromV: 'page', offsetYPx: 820 },
		});
		const note = { id: 'n1', paragraphs: [paragraph('n1p', 'note')] };
		const body = [paragraph('p', 'text', { footnotes: [note] })];
		const plain = layoutDocument(input(body), measurer);
		const floated = layoutDocument(input([low, ...body]), measurer);
		expect(tables(floated)[0]?.floating).toBeDefined();
		expect(floated.pages.length).toBe(1);
		expect(floated.pages[0]!.footnotes?.[0]?.yPx).toBe(plain.pages[0]!.footnotes?.[0]?.yPx);
	});

	it('lets a continuous section start below the flow, not below the table', () => {
		const low: LayoutTable = floatingTable(2, {
			floating: { ...frame, relativeFromV: 'page', offsetYPx: 600 },
		});
		const result = layoutDocument(
			input([low, paragraph('p')], [{ page, break: 'continuous', blocks: [paragraph('next')] }]),
			measurer,
		);
		expect(result.pages.length).toBe(1);
		const next = result.pages[0]!.columns.flatMap((column) => column.blocks).find(
			(block) => block.blockId === 'next',
		);
		expect(next?.yPx).toBeLessThan(500);
	});
});
