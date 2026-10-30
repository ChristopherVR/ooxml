import { describe, expect, it } from 'vitest';
import { editVsdx, parseVsdx, type VisioEdit, type VisioPage } from '../index';
import { createTemplateVsdx } from './templates';
import {
	visioDeleteSwimlane,
	visioInsertSeparator,
	visioInsertSwimlane,
	visioResizeSwimlane,
	visioSwimlaneInfo,
	visioSwimlaneOptions,
} from './swimlane';
import { visioContainerUpkeep } from './diagram-parts';

async function page(bytes: Uint8Array): Promise<VisioPage> {
	return (await parseVsdx(bytes)).pages[0]!;
}
const shapeBox = (shape: VisioPage['shapes'][number]) => {
	const [a, , , d, e, f] = shape.transform;
	return { minX: e, maxX: e + a * shape.width, minY: f, maxY: f + d * shape.height };
};
const find = (p: VisioPage, id: string) => p.shapes.find((shape) => shape.id === id)!;

/** The template's process shapes are stencil instances, so the planner must move them itself. */
async function template() {
	return createTemplateVsdx('cross-functional-flowchart');
}
async function apply(
	bytes: Uint8Array,
	edits: VisioEdit[] | VisioEdit[][] | undefined,
): Promise<Uint8Array> {
	expect(edits).toBeDefined();
	const batches = Array.isArray(edits![0]) ? (edits as VisioEdit[][]) : [edits as VisioEdit[]];
	let current = bytes;
	for (const batch of batches) current = (await editVsdx(current, batch)).bytes;
	return current;
}

describe('cross-functional flowchart planning', () => {
	it('finds the flowchart from the frame, a lane or a shape in a lane', async () => {
		const p = await page(await template());
		for (const id of ['50', '10', '2']) expect(visioSwimlaneInfo(p, id)?.frame.id).toBe('50');
		expect(visioSwimlaneInfo(p, '50')?.lanes.map((lane) => lane.id)).toEqual(['10', '11', '12']);
		expect(visioSwimlaneInfo(p, '999')).toBeUndefined();
	});

	it('inserts a lane after the selected one, carrying the stencil shapes below it', async () => {
		const bytes = await template();
		const before = await page(bytes);
		const next = await page(
			await apply(bytes, visioInsertSwimlane(before, '10', 'after', 'New function')),
		);
		const info = visioSwimlaneInfo(next, '50')!;
		expect(info.lanes).toHaveLength(4);
		expect(info.lanes[1]!.text?.plainText).toBe('New function');
		const lane = shapeBox(info.lanes[1]!);
		expect(lane.maxY - lane.minY).toBeCloseTo(1.8);
		// Sales and Fulfilment moved down by one lane, and the shapes in them with them.
		const drop = shapeBox(find(before, '2')).maxY - shapeBox(find(next, '2')).maxY;
		expect(drop).toBeCloseTo(1.8);
		expect(shapeBox(find(next, '4')).maxY).toBeCloseTo(shapeBox(find(before, '4')).maxY - 1.8);
		expect(shapeBox(find(next, '1')).maxY).toBeCloseTo(shapeBox(find(before, '1')).maxY);
		expect(shapeBox(find(next, '50')).minY).toBeCloseTo(shapeBox(find(before, '50')).minY - 1.8);
		// The glued connectors still end on their shapes.
		expect(next.connectors).toHaveLength(before.connectors.length);
	});

	it('makes a lane taller and pushes the lanes below', async () => {
		const bytes = await template();
		const before = await page(bytes);
		const next = await page(await apply(bytes, visioResizeSwimlane(before, '11', 3)));
		const lane = shapeBox(find(next, '11'));
		expect(lane.maxY - lane.minY).toBeCloseTo(3);
		expect(shapeBox(find(next, '12')).maxY).toBeCloseTo(lane.minY);
		expect(shapeBox(find(next, '4')).maxY).toBeCloseTo(shapeBox(find(before, '4')).maxY - 1.2);
		expect(shapeBox(find(next, '2')).maxY).toBeCloseTo(shapeBox(find(before, '2')).maxY);
	});

	it('deletes a lane with its shapes, instances included', async () => {
		const bytes = await template();
		const before = await page(bytes);
		const next = await page(await apply(bytes, visioDeleteSwimlane(before, '11')));
		expect(next.shapes.some((shape) => ['11', '2', '3'].includes(shape.id))).toBe(false);
		expect(visioSwimlaneInfo(next, '50')?.lanes.map((lane) => lane.id)).toEqual(['10', '12']);
		expect(shapeBox(find(next, '12')).maxY).toBeCloseTo(shapeBox(find(next, '10')).minY);
		expect(shapeBox(find(next, '4')).maxY).toBeCloseTo(shapeBox(find(before, '4')).maxY + 1.8);
		expect(visioDeleteSwimlane(next, '10')).toBeDefined();
	});

	it('refuses to delete the last lane', async () => {
		let bytes = await template();
		bytes = await apply(bytes, visioDeleteSwimlane(await page(bytes), '10'));
		bytes = await apply(bytes, visioDeleteSwimlane(await page(bytes), '11'));
		expect(visioDeleteSwimlane(await page(bytes), '12')).toBeUndefined();
	});

	it('adds a separator and toggles the title bar without losing the shapes', async () => {
		let bytes = await template();
		bytes = await apply(bytes, visioInsertSeparator(await page(bytes), '50'));
		const withBar = await page(bytes);
		expect(visioSwimlaneInfo(withBar, '50')?.frame.structure).toMatchObject({
			memberIds: expect.arrayContaining(['10', '11', '12']),
		});
		const hidden = await page(
			await apply(bytes, visioSwimlaneOptions(withBar, '50', { showTitle: false })),
		);
		expect(visioSwimlaneInfo(hidden, '50')?.options.showTitle).toBe(false);
		expect(shapeBox(find(hidden, '10')).maxY).toBeCloseTo(shapeBox(find(hidden, '50')).maxY);
		expect(shapeBox(find(hidden, '1')).maxY).toBeCloseTo(shapeBox(find(withBar, '1')).maxY + 0.35);
	});

	it('keeps lane membership when shapes are moved within the flowchart', async () => {
		const bytes = await template();
		const p = await page(bytes);
		const upkeep = visioContainerUpkeep(p, [{ type: 'move-shape', pageId: p.id, shapeId: '2' }]);
		expect(upkeep).toMatchObject({ type: 'reconcile-containers' });
	});
});
