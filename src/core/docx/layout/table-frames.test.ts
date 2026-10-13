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
