import { expect, it } from 'vitest';
import type { VisioEdit } from '../edit-commands';
import { snapshotEdits } from './edit-commands';

it('snapshots endpoint commands without caller extras or later mutation', () => {
	const command = {
		type: 'move-line-endpoint' as const,
		pageId: '0',
		shapeId: '1',
		endpoint: 'begin' as const,
		x: 2,
		y: 3,
		extra: new Map(),
	};
	const saved = snapshotEdits([command]);
	command.x = 99;
	expect(saved).toEqual([
		{ type: command.type, pageId: '0', shapeId: '1', endpoint: 'begin', x: 2, y: 3 },
	]);
});

it('rejects invalid endpoint identity and coordinates before dispatch', () => {
	for (const command of [
		{ type: 'move-line-endpoint', pageId: '0', shapeId: '1', endpoint: 'middle', x: 2, y: 3 },
		{ type: 'move-line-endpoint', pageId: '0', shapeId: '1', endpoint: 'end', x: NaN, y: 3 },
	])
		expect(() => snapshotEdits([command as VisioEdit])).toThrow();
});
