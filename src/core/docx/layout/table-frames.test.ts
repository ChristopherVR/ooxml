import { splitRowAtHeight } from './table-layout';
import { expect, it } from 'vitest';
import { clipTableFrames } from './table-frames';
import type { LayoutTableBox } from './result';
it('clips and rebases recursive decorative frames without mutating source geometry', () => {
	const inner: LayoutTableBox = {
		kind: 'table',
		blockId: 'inner',
		xPx: 12,
		widthPx: 30,
		yPx: 5,
		heightPx: 30,
		rows: [{ yPx: 0, heightPx: 30, repeated: false, cells: [[]] }],
	};
	const outer: LayoutTableBox = {
		...inner,
		blockId: 'outer',
		yPx: 10,
		heightPx: 50,
		rows: [{ yPx: 0, heightPx: 50, repeated: false, cells: [[]], cellTables: [[inner]] }],
	};
	const source = structuredClone(outer);
	const [fragment] = clipTableFrames([outer], 25, 50, 3);
	expect(fragment).toMatchObject({ xPx: 12, widthPx: 30, yPx: 3, heightPx: 25 });
	expect(fragment!.rows[0]!.cellTables![0]![0]).toMatchObject({ yPx: 0, heightPx: 20 });
	expect(outer).toEqual(source);
	expect(clipTableFrames([outer], 60, 80)).toEqual([]);
});

it('reserves nested frame edges again on a physical row continuation', () => {
	const paragraph = {
		kind: 'paragraph' as const,
		blockId: 'p',
		yPx: 4,
		heightPx: 48,
		tableFrameInsets: { top: 4, bottom: 4 },
		lines: [0, 16, 32].map((yPx) => ({
			yPx,
			heightPx: 16,
			firstToken: yPx / 16,
			sourceStart: yPx / 16,
			sourceEnd: yPx / 16 + 1,
			lastToken: yPx / 16,
			fragments: [],
			widthPx: 0,
			baselinePx: 12,
		})),
	};
	const row = {
		heightPx: 56,
		cells: [[paragraph]],
		geometry: [{ xPx: 0, widthPx: 80, paddingLeftPx: 0, paddingRightPx: 0, contentHeightPx: 56 }],
		cellTables: [
			[
				{
					kind: 'table' as const,
					blockId: 'inner',
					widthPx: 80,
					yPx: 0,
					heightPx: 56,
					rows: [{ yPx: 0, heightPx: 56, repeated: false, cells: [[]] }],
				},
			],
		],
	};
	const split = splitRowAtHeight(row, 34);
	expect(split.before.heightPx).toBe(24);
	expect(split.before.cells[0]![0]!.lines).toHaveLength(1);
	expect(split.before.cellTables![0]![0]!.heightPx).toBe(24);
	expect(split.after!.cells[0]![0]!.yPx).toBe(4);
	expect(split.after!.heightPx).toBe(40);
	expect(split.after!.cellTables![0]![0]!.yPx).toBe(0);
});

it('measures the longest parallel nested continuation regardless of paragraph ordering', () => {
	const lines = (count: number) =>
		Array.from({ length: count }, (_, i) => ({
			yPx: i * 16,
			heightPx: 16,
			firstToken: i,
			lastToken: i + 1,
			sourceStart: i,
			sourceEnd: i + 1,
			fragments: [],
		}));
	const long = {
		kind: 'paragraph' as const,
		blockId: 'long',
		yPx: 4,
		heightPx: 64,
		tableFrameInsets: { top: 4, bottom: 4 },
		lines: lines(4),
	};
	const short = { ...long, blockId: 'short', heightPx: 48, lines: lines(3) };
	for (const cells of [
		[long, short],
		[short, long],
	]) {
		const row = {
			heightPx: 72,
			cells: [cells],
			geometry: [
				{ xPx: 0, widthPx: 160, paddingLeftPx: 0, paddingRightPx: 0, contentHeightPx: 72 },
			],
			cellPadding: [{ top: 2, bottom: 3 }],
		};
		const { after } = splitRowAtHeight(row, 27);
		expect(after!.heightPx).toBe(61);
		expect(after!.geometry[0]!.contentHeightPx).toBe(61);
		expect(after!.cells[0]!.map((box) => [box.blockId, box.yPx, box.heightPx])).toContainEqual([
			'long',
			6,
			48,
		]);
	}
});
