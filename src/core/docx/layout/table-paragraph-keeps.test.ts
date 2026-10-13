import { expect, it } from 'vitest';
import { signedTwips, twips } from '../units';
import type { LayoutParagraph, LayoutTable } from './input';
import { tableParagraphKeeps } from './table-paragraph-keeps';
import { layoutBorderedRow } from './table-border-layout';
import { nestedFrameProfile } from './nested-frame-profile';
import { splitRowAtHeight } from './table-layout';
import { layoutSections } from './page-flow';

const p = (id: string, lines: number, keepLines = false): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	keepLines,
	widowControl: false,
	lineSpacingRule: 'exact',
	lineSpacingTwips: signedTwips(240),
	runs: [{ text: Array.from({ length: lines }, (_, i) => id + i).join('\n') }],
});
const source = (prefix = false, side = 0): LayoutTable => ({
	kind: 'table',
	id: 'table',
	widthPx: 200,
	rows: [
		{
			cells: [0, 1].map((index) => ({
				widthPx: 100,
				xPx: index * 100,
				paragraphs:
					index === side
						? [...(prefix ? [p('prefix', 3)] : []), p('keeper', 4, true), p('suffix', 2)]
						: [p('neighbor', 9)],
				borders: Object.fromEntries(
					['top', 'bottom', 'left', 'right'].map((edge) => [
						edge,
						{ widthPx: 4, style: 'solid' as const, color: '#FF0000' },
					]),
				),
			})),
		},
	],
});
const measure = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
const layout = (table: LayoutTable, remaining?: Parameters<typeof layoutBorderedRow>[5]) =>
	layoutBorderedRow(table.rows[0]!, 200, measure, () => {}, nestedFrameProfile(table), remaining);

it.each([0, 1])('uses either first kept cell for whole-row startup: %i', (side) => {
	const table = source(false, side);
	expect(tableParagraphKeeps(table)!.startupMinimum(layout(table))).toBe(72);
});
it('defers the kept paragraph and following content, leaving the neighbor independent', () => {
	const table = source(true),
		policy = tableParagraphKeeps(table)!,
		row = layout(table);
	const cuts = policy.cuts(row, 88, true)!;
	expect([...cuts.keys()]).toEqual(['keeper', 'suffix']);
	const split = splitRowAtHeight(row, 88, undefined, cuts);
	expect(split.before.cells[0]!.map((box) => box.blockId)).toEqual(['prefix']);
	expect(split.before.cells[1]![0]!.lines).toHaveLength(5);
	const continued = layout(table, split.after!.cells);
	// An oversized paragraph gets one deferral, then splits at a fresh origin.
	expect(policy.cuts(continued, 40, true)).toBeUndefined();
	const next = splitRowAtHeight(continued, 40).after!;
	expect(next.cells[0]![0]!.lines[0]!.firstToken).toBeGreaterThan(0);
	expect(policy.cuts(next, 40, false)).toBeUndefined();
});
it.each([
	'note',
	'keep-next',
	'page-break',
	'nested',
	'height',
	'zero-height',
	'header',
	'cant-split',
	'grid-span',
	'vertical-merge',
	'alignment',
	'missing-border',
	'dashed-border',
	'width',
	'rows',
	'cells',
	'widows',
	'spacing-before',
	'spacing-after',
	'auto-spacing',
	'paragraph-frame',
	'duplicate-id',
	'unbordered',
])('refuses unmeasured keep profile: %s', (kind) => {
	const table = source(),
		row = table.rows[0]!,
		cell = row.cells[0]!,
		paragraph = cell.paragraphs[0]!;
	if (kind === 'note') paragraph.runs[0]!.footnoteReferenceId = '1';
	if (kind === 'keep-next') paragraph.keepNext = true;
	if (kind === 'page-break') paragraph.pageBreakBefore = true;
	if (kind === 'nested') cell.blocks = [source()];
	if (kind === 'height') row.heightPx = 20;
	if (kind === 'zero-height') row.heightSpecified = true;
	if (kind === 'header') row.isHeader = true;
	if (kind === 'cant-split') row.cantSplit = true;
	if (kind === 'grid-span') cell.gridSpan = 2;
	if (kind === 'vertical-merge') cell.verticalMerge = 'restart';
	if (kind === 'alignment') cell.verticalAlign = 'center';
	if (kind === 'missing-border') delete cell.borders!.bottom;
	if (kind === 'dashed-border') cell.borders!.bottom!.style = 'dashed';
	if (kind === 'width') delete table.widthPx;
	if (kind === 'rows') table.rows.push(structuredClone(row));
	if (kind === 'cells') row.cells.push(structuredClone(cell));
	if (kind === 'widows') paragraph.widowControl = true;
	if (kind === 'spacing-before') paragraph.spacingBeforeTwips = twips(20);
	if (kind === 'spacing-after') paragraph.spacingAfterTwips = twips(20);
	if (kind === 'auto-spacing') paragraph.lineSpacingRule = 'auto';
	if (kind === 'paragraph-frame') paragraph.shading = '#FF0000';
	if (kind === 'duplicate-id') row.cells[1]!.paragraphs[0]!.id = paragraph.id;
	if (kind === 'unbordered') for (const item of row.cells) delete item.borders;
	expect(tableParagraphKeeps(table)).toBeUndefined();
});
it('keeps source blocks authoritative and refuses unsupported keeping on an inserted direct story', () => {
	const table = source();
	table.rows[0]!.cells[0]!.blocks = [{ ...p('blocks', 3, true), keepNext: true }];
	expect(tableParagraphKeeps(table)).toBeUndefined();
});
it('uses blocks-only direct source stories for keep and continuation cuts', () => {
	const table = source(),
		cell = table.rows[0]!.cells[0]!;
	cell.blocks = cell.paragraphs;
	cell.paragraphs = [];
	const policy = tableParagraphKeeps(table)!;
	expect(policy.startupMinimum(layout(table))).toBe(72);
	expect(policy.cuts(layout(table), 40, false)?.has('keeper')).toBe(true);
});
it('retains fallback and reports keeping outside a single body column', () => {
	const table = source();
	table.rows[0]!.cells[0]!.paragraphs = [p('keeper', 3, true)];
	const result = layoutSections(
		{
			sections: [
				{
					page: {
						widthPx: 400,
						heightPx: 60,
						marginTopPx: 0,
						marginRightPx: 0,
						marginBottomPx: 0,
						marginLeftPx: 0,
					},
					columns: { count: 2, gapPx: 0 },
					blocks: [p('body', 1), table],
				},
			],
		},
		measure,
	);
	expect(result.approximations.some((warning) => warning.includes('Table-cell keep-lines'))).toBe(
		true,
	);
	const first = result.pages[0]!.columns[0]!.blocks.find((block) => block.kind === 'table');
	if (!first || first.kind !== 'table') throw new Error('Expected prior partial-row fallback');
	expect(first.rows[0]!.cells[0]![0]!.lines).toHaveLength(2);
});
