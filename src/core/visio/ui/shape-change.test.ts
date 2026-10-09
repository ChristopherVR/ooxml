import { expect, it } from 'vitest';
import type { VisioPage, VisioShape } from '../model';
import { demoDocument } from './demo-document';
import { visioChangeShapeCommand, visioChangeShapeRefusal } from './shape-change';

function page(change: Partial<VisioShape> = {}): VisioPage {
	const template = demoDocument.pages[0]!.shapes[0]!;
	const shape = { ...structuredClone(template), id: '7', kind: 'shape' as const, ...change };
	delete shape.masterId;
	Object.assign(shape, change);
	return { ...structuredClone(demoDocument.pages[0]!), shapes: [shape], connectors: [] };
}
const one = [{ id: '7' }];

it('admits one local 2D shape and builds the command', () => {
	expect(visioChangeShapeRefusal(page(), one)).toBeUndefined();
	expect(visioChangeShapeCommand(page(), one, 'star')).toEqual({
		type: 'change-shape',
		pageId: demoDocument.pages[0]!.id,
		shapeId: '7',
		shape: 'star',
	});
});

it.each([
	['nothing selected', page(), [], /exactly one/],
	['two shapes', page(), [{ id: '7' }, { id: '8' }], /exactly one/],
	['a nested shape', page(), [{ id: '99' }], /inside groups/],
	['a master', page({ masterId: '3' }), one, /Master shapes/],
	['a group', page({ kind: 'group' }), one, /Groups/],
	['a line', page({ kind: 'connector' }), one, /1D shapes/],
	['a picture', page({ kind: 'foreign' }), one, /Pictures/],
])('refuses %s', (_name, scene, selection, reason) => {
	expect(visioChangeShapeRefusal(scene, selection)).toMatch(reason);
	expect(visioChangeShapeCommand(scene, selection, 'circle')).toBeUndefined();
});

it('refuses glued shapes', () => {
	const scene = page();
	scene.connectors = [{ fromShapeId: '9', toShapeId: '7', fromCell: 'BeginX', toCell: 'PinX' }];
	expect(visioChangeShapeRefusal(scene, one)).toMatch(/Glued/);
});
