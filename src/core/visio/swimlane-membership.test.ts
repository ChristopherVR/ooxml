import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';

const rect = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width: 1,
	height: 0.6,
});

describe('shapes lying in a lane when a drawing loads', () => {
	it('join the lane they lie in, after its saved members and without touching list order', async () => {
		const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
		const shapes = await editVsdx(blank, [rect('1', 4, 6.8), rect('2', 6, 4.6), rect('9', 40, 40)]);
		// No reconcile: the lanes are written around the shapes and list none of them.
		const made = await editVsdx(shapes.bytes, [
			{
				type: 'create-swimlanes',
				pageId: '0',
				frameId: '50',
				laneIds: ['51', '52', '53'],
				headings: ['A', 'B', 'C'],
				box: { x: 6, y: 5, width: 8, height: 6.35 },
				direction: 'horizontal',
				title: 'Flow',
				listId: '54',
				phaseListId: '55',
			},
		]);
		const page = (await parseVsdx(made.bytes)).pages[0]!;
		const byId = new Map(page.shapes.map((shape) => [shape.id, shape]));
		const members = (id: string) => {
			const structure = byId.get(id)!.structure;
			return structure?.type === 'container' ? structure.memberIds : [];
		};
		const lanes = ['51', '52', '53'].map(members);
		expect(lanes.flat().sort()).toEqual(['1', '2']);
		expect(lanes.find((ids) => ids.includes('1'))).not.toContain('2');
		// The lists keep the saved order, and a shape off the page's lanes joins none.
		expect(byId.get('54')!.structure).toMatchObject({ memberIds: ['51', '52', '53'] });
		expect(lanes.flat()).not.toContain('9');
	});
});
