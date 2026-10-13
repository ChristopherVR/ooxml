import { describe, expect, it } from 'vitest';
import { snapshotEdits } from './edit-commands';

describe('position-following edits through the UI boundary', () => {
	for (const type of ['create-text-box', 'create-rectangle', 'create-ellipse'] as const) {
		const command = {
			type,
			pageId: '0',
			shapeId: '2',
			x: 3,
			y: 2,
			width: 1,
			height: 1,
			text: 'Value',
		};
		it(`${type} retains the owner reference through repeated snapshots`, () => {
			const edit = { ...command, followShapeId: '1' };
			expect(snapshotEdits(snapshotEdits([edit]))).toEqual([edit]);
			expect(snapshotEdits([command])).toEqual([command]);
		});
		it.each(['2', '1)!Width', '4294967296'])(
			`${type} rejects invalid owner %s`,
			(followShapeId) => {
				expect(() => snapshotEdits([{ ...command, followShapeId }])).toThrow(/owner/);
			},
		);
	}
});
