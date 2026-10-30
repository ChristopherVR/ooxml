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

	it('deletes a lane that holds stencil shapes in one batch, so one undo step', async () => {
		const before = await page(await template());
		const batches = visioDeleteSwimlane(before, '11')!;
		expect(batches).toHaveLength(1);
		expect(batches[0]!.some((edit) => edit.type === 'delete-shape')).toBe(false);
	});

	it('turns the flowchart on its side and back, stencil shapes and connectors with it', async () => {
		const bytes = await template();
		const before = await page(bytes);
		const turned = await apply(
			bytes,
			visioSwimlaneOptions(before, '50', { direction: 'vertical' }),
		);
		const next = await page(turned);
		const info = visioSwimlaneInfo(next, '50')!;
		expect(info.options.direction).toBe('vertical');
		const lanes = info.lanes.map((lane) => shapeBox(lane));
		const old = visioSwimlaneInfo(before, '50')!.lanes.map((lane) => shapeBox(lane));
		lanes.forEach((lane, index) => {
			expect(lane.maxX - lane.minX).toBeCloseTo(old[index]!.maxY - old[index]!.minY);
			if (index) expect(lane.minX).toBeCloseTo(lanes[index - 1]!.maxX);
		});
		// Shape 2 (a stencil instance in Sales) sits inside the turned Sales lane.
		const sales = lanes[1]!;
		const two = shapeBox(find(next, '2'));
		expect(two.minX).toBeGreaterThanOrEqual(sales.minX - 1e-6);
		expect(two.maxX).toBeLessThanOrEqual(sales.maxX + 1e-6);
		expect(next.connectors).toHaveLength(before.connectors.length);
		// Turning the same way again is a no-op; turning back restores the lanes' places.
		expect(visioSwimlaneOptions(next, '50', { direction: 'vertical' })).toEqual([]);
		expect(visioSwimlaneOptions(next, '50', { direction: 'vertical', showTitle: false })).toBe(
			undefined,
		);
		const back = await page(
			await apply(turned, visioSwimlaneOptions(next, '50', { direction: 'horizontal' })),
		);
		for (const id of ['50', '10', '11', '12', '1', '2', '3', '4']) {
			const a = shapeBox(find(back, id));
			const b = shapeBox(find(before, id));
			expect(a.minX).toBeCloseTo(b.minX, 2);
			expect(a.maxY).toBeCloseTo(b.maxY, 2);
		}
	});

	it('grows the page with a frame that outgrows it, in the same batch', async () => {
		const bytes = await template();
		const before = await page(bytes);
		const frame = shapeBox(find(before, '50'));
		const edits = visioResizeSwimlane(before, '10', 3)!;
		// Making a lane taller grows the frame downwards: no page growth is needed.
		expect(edits.some((edit) => edit.type === 'set-page-size')).toBe(false);
		const turned = visioSwimlaneOptions(before, '50', { direction: 'vertical' })!;
		const growth = turned.find((edit) => edit.type === 'set-page-size');
		const wide = shapeBox(find(before, '10')).maxY - shapeBox(find(before, '10')).minY;
		const total = ['10', '11', '12']
			.map((id) => shapeBox(find(before, id)))
			.reduce((sum, lane) => sum + (lane.maxY - lane.minY), 0);
		expect(wide).toBeGreaterThan(0);
		if (frame.minX + total > before.width) expect(growth).toBeDefined();
		const next = await page(await apply(bytes, turned));
		expect(shapeBox(find(next, '50')).maxX).toBeLessThanOrEqual(next.width + 1e-6);
		expect(shapeBox(find(next, '50')).maxY).toBeLessThanOrEqual(next.height + 1e-6);
	});

	it('makes the page room when a lane is widened past its edge, in one batch', async () => {
		const start = await apply(
			await template(),
			visioSwimlaneOptions(await page(await template()), '50', { direction: 'vertical' }),
		);
		const before = await page(start);
		const edits = visioResizeSwimlane(before, '10', 30)!;
		expect(edits[0]).toMatchObject({ type: 'set-page-size', pageId: before.id });
		const next = await page(await apply(start, edits));
		expect(next.width).toBeGreaterThan(before.width);
		expect(next.height).toBeCloseTo(before.height);
		expect(shapeBox(find(next, '50')).maxX).toBeLessThanOrEqual(next.width);
		expect(shapeBox(find(next, '10')).maxX - shapeBox(find(next, '10')).minX).toBeCloseTo(30);
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
		expect(visioSwimlaneInfo(withBar, '50')?.lanes.map((lane) => lane.id)).toEqual([
			'10',
			'11',
			'12',
		]);
		const phases = withBar.shapes
			.map((shape) => shape.structure)
			.find((structure) => structure?.type === 'list' && structure.kind === 'phase');
		expect(phases).toMatchObject({ memberIds: [expect.any(String)] });
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
