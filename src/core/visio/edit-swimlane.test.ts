import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';
import type { VisioShape } from './model';

const rect = (shapeId: string, x: number, y: number, width = 1, height = 0.6): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width,
	height,
});
const create = (direction: 'horizontal' | 'vertical' = 'horizontal'): VisioEdit => ({
	type: 'create-swimlanes',
	pageId: '0',
	frameId: '50',
	laneIds: ['51', '52', '53'],
	headings: ['A', 'B', 'C'],
	box: { x: 6, y: 5, width: 8, height: 6.35 },
	direction,
	title: 'Flow',
});
const reconcile = (...shapeIds: string[]): VisioEdit => ({
	type: 'reconcile-containers',
	pageId: '0',
	shapeIds,
});
async function run(bytes: Uint8Array, edits: VisioEdit[]): Promise<Uint8Array> {
	return (await editVsdx(bytes, edits)).bytes;
}
async function shapes(bytes: Uint8Array): Promise<Map<string, VisioShape>> {
	const page = (await parseVsdx(bytes)).pages[0]!;
	return new Map(page.shapes.map((shape) => [shape.id, shape]));
}
const box = (shape: VisioShape) => {
	const [a, , , d, e, f] = shape.transform;
	return { minX: e, maxX: e + a * shape.width, minY: f, maxY: f + d * shape.height };
};
const members = (map: Map<string, VisioShape>, id: string): string[] => {
	const structure = map.get(id)?.structure;
	return structure?.type === 'container' ? structure.memberIds : [];
};
/** The lanes of a flowchart in order, through its Swimlane List. */
const laneIds = (map: Map<string, VisioShape>, frameId: string): string[] => {
	for (const id of members(map, frameId)) {
		const structure = map.get(id)?.structure;
		if (structure?.type === 'list' && structure.kind === 'swimlane') return structure.memberIds;
	}
	return [];
};
async function built(direction: 'horizontal' | 'vertical' = 'horizontal') {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	// Process shapes first, then the lanes behind them, then membership by reconcile.
	const start = await run(blank, [rect('1', 4, 6.8), rect('2', 6, 4.6), rect('3', 8, 2.7)]);
	return run(start, [create(direction), reconcile('1', '2', '3')]);
}

describe('cross-functional flowchart', () => {
	it('stacks lanes edge to edge inside a frame that spans them', async () => {
		const map = await shapes(await built());
		const frame = box(map.get('50')!);
		const lanes = ['51', '52', '53'].map((id) => box(map.get(id)!));
		expect(lanes[0]!.maxY).toBeCloseTo(frame.maxY - 0.35);
		expect(lanes[1]!.maxY).toBeCloseTo(lanes[0]!.minY);
		expect(lanes[2]!.maxY).toBeCloseTo(lanes[1]!.minY);
		expect(lanes[2]!.minY).toBeCloseTo(frame.minY);
		for (const lane of lanes) {
			expect(lane.minX).toBeCloseTo(frame.minX);
			expect(lane.maxX).toBeCloseTo(frame.maxX);
		}
		expect(map.get('50')!.structure).toMatchObject({
			type: 'container',
			swimlane: { role: 'frame', options: { direction: 'horizontal', showTitle: true } },
		});
		expect(map.get('51')!.structure).toMatchObject({ swimlane: { role: 'lane' } });
	});

	it('makes the shapes dropped into a lane its members', async () => {
		const map = await shapes(await built());
		expect(members(map, '51')).toContain('1');
		expect(members(map, '52')).toContain('2');
		expect(members(map, '53')).toContain('3');
	});

	it('pushes the lanes below and grows the frame when a lane is made taller', async () => {
		const before = await shapes(await built());
		const grown = await shapes(
			await run(await built(), [
				{ type: 'resize-swimlane', pageId: '0', frameId: '50', laneId: '51', size: 3 },
			]),
		);
		const lane0 = box(grown.get('51')!);
		expect(lane0.maxY - lane0.minY).toBeCloseTo(3);
		expect(lane0.maxY).toBeCloseTo(box(before.get('51')!).maxY);
		expect(box(grown.get('52')!).maxY).toBeCloseTo(lane0.minY);
		const frame = box(grown.get('50')!);
		expect(frame.maxY).toBeCloseTo(box(before.get('50')!).maxY);
		expect(frame.minY).toBeLessThan(box(before.get('50')!).minY);
		expect(frame.minY).toBeCloseTo(box(grown.get('53')!).minY);
		// The shapes of the lanes below went down with their lanes; the first lane's did not.
		const delta = box(grown.get('52')!).maxY - box(before.get('52')!).maxY;
		expect(delta).toBeLessThan(0);
		expect(box(grown.get('2')!).maxY - box(before.get('2')!).maxY).toBeCloseTo(delta);
		expect(box(grown.get('1')!).maxY).toBeCloseTo(box(before.get('1')!).maxY);
	});

	it('inserts a lane after the selected one and shifts the rest', async () => {
		const before = await shapes(await built());
		const map = await shapes(
			await run(await built(), [
				{
					type: 'insert-swimlane',
					pageId: '0',
					frameId: '50',
					shapeId: '60',
					relativeTo: '51',
					position: 'after',
					heading: 'New',
				},
			]),
		);
		expect(laneIds(map, '50')).toEqual(['51', '60', '52', '53']);
		const added = box(map.get('60')!);
		expect(added.maxY).toBeCloseTo(box(map.get('51')!).minY);
		expect(box(map.get('52')!).maxY).toBeCloseTo(added.minY);
		expect(box(before.get('52')!).maxY - box(map.get('52')!).maxY).toBeCloseTo(
			added.maxY - added.minY,
		);
		expect(box(map.get('50')!).minY).toBeCloseTo(box(map.get('53')!).minY);
	});

	it('inserts a lane before the selected one', async () => {
		const map = await shapes(
			await run(await built(), [
				{
					type: 'insert-swimlane',
					pageId: '0',
					frameId: '50',
					shapeId: '60',
					relativeTo: '51',
					position: 'before',
					heading: 'First',
				},
			]),
		);
		expect(box(map.get('60')!).maxY).toBeCloseTo(box(map.get('50')!).maxY - 0.35);
		expect(box(map.get('51')!).maxY).toBeCloseTo(box(map.get('60')!).minY);
	});

	it('deletes a lane with its shapes and closes the gap', async () => {
		const before = await shapes(await built());
		const map = await shapes(
			await run(await built(), [
				{ type: 'delete-swimlane', pageId: '0', frameId: '50', laneId: '52' },
			]),
		);
		expect(map.has('52')).toBe(false);
		expect(map.has('2')).toBe(false);
		expect(map.has('1')).toBe(true);
		expect(box(map.get('53')!).maxY).toBeCloseTo(box(map.get('51')!).minY);
		expect(laneIds(map, '50')).toEqual(['51', '53']);
		expect(box(map.get('50')!).minY).toBeGreaterThan(box(before.get('50')!).minY);
		expect(box(map.get('50')!).minY).toBeCloseTo(box(map.get('53')!).minY);
	});

	it('keeps the last lane', async () => {
		const bytes = await run(await built(), [
			{ type: 'delete-swimlane', pageId: '0', frameId: '50', laneId: '51' },
			{ type: 'delete-swimlane', pageId: '0', frameId: '50', laneId: '52' },
		]);
		await expect(
			run(bytes, [{ type: 'delete-swimlane', pageId: '0', frameId: '50', laneId: '53' }]),
		).rejects.toThrow(/last lane/);
	});

	it('adds a phase separator that spans the lanes and follows their growth', async () => {
		const bytes = await run(await built(), [
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '70', at: 6 },
			{ type: 'resize-swimlane', pageId: '0', frameId: '50', laneId: '53', size: 3 },
		]);
		const map = await shapes(bytes);
		const bar = box(map.get('70')!);
		expect(bar.maxX - bar.minX).toBeCloseTo(0.04);
		expect(bar.minY).toBeCloseTo(box(map.get('53')!).minY);
		expect(bar.maxY).toBeCloseTo(box(map.get('51')!).maxY);
	});

	it('records title, label, separator and style options', async () => {
		const bytes = await run(await built(), [
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '70' },
			{
				type: 'set-swimlane-options',
				pageId: '0',
				frameId: '50',
				showTitle: false,
				showSeparators: false,
				rotateLabels: true,
				style: 2,
			},
		]);
		const map = await shapes(bytes);
		expect(map.get('50')!.structure).toMatchObject({
			swimlane: {
				options: { showTitle: false, showSeparators: false, rotateLabels: true, style: 2 },
			},
		});
		// With the title bar gone the first lane reaches the top of the frame.
		expect(box(map.get('51')!).maxY).toBeCloseTo(box(map.get('50')!).maxY);
		const back = await shapes(
			await run(bytes, [
				{ type: 'set-swimlane-options', pageId: '0', frameId: '50', showTitle: true },
			]),
		);
		expect(box(back.get('51')!).maxY).toBeCloseTo(box(back.get('50')!).maxY - 0.35);
	});

	it('lays vertical flows out left to right', async () => {
		const map = await shapes(await built('vertical'));
		const frame = box(map.get('50')!);
		const lanes = ['51', '52', '53'].map((id) => box(map.get(id)!));
		expect(lanes[0]!.minX).toBeCloseTo(frame.minX);
		expect(lanes[1]!.minX).toBeCloseTo(lanes[0]!.maxX);
		expect(lanes[2]!.maxX).toBeCloseTo(frame.maxX);
		const grown = await shapes(
			await run(await built('vertical'), [
				{ type: 'resize-swimlane', pageId: '0', frameId: '50', laneId: '51', size: 4 },
			]),
		);
		expect(box(grown.get('52')!).minX).toBeCloseTo(box(grown.get('51')!).maxX);
		expect(box(grown.get('50')!).maxX).toBeCloseTo(box(grown.get('53')!).maxX);
	});

	it('refuses a frame that is not a flowchart', async () => {
		await expect(
			run(await built(), [{ type: 'delete-swimlane', pageId: '0', frameId: '1', laneId: '51' }]),
		).rejects.toThrow();
	});

	it('adds separators when the lanes are the last shapes on the page', async () => {
		// The separator used to be inserted before itself there, which corrupted the page's XML.
		let bytes = await fixture({ pages: [{ id: '0', contents: '' }] });
		for (const edit of [
			create(),
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '70', at: 5, label: 'P1' },
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '71', at: 8, label: 'P2' },
		] as VisioEdit[])
			bytes = await run(bytes, [edit]);
		const all = await shapes(bytes);
		expect(all.has('70')).toBe(true);
		expect(all.has('71')).toBe(true);
	});
});
