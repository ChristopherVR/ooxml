import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';
import type { VisioShape } from './model';

const rect = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width: 1,
	height: 0.6,
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
	listId: '54',
	phaseListId: '55',
});
const run = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
async function built(direction: 'horizontal' | 'vertical' = 'horizontal') {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	const start = await run(blank, [rect('1', 4, 6.8), rect('2', 6, 4.6), rect('3', 8, 2.7)]);
	return run(start, [
		create(direction),
		{ type: 'reconcile-containers', pageId: '0', shapeIds: ['1', '2', '3'] },
	]);
}
const pageXml = async (bytes: Uint8Array) =>
	await (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
const relationships = (xml: string, id: string) =>
	new RegExp(`<Shape ID="${id}"[\\s\\S]*?N="Relationships" V="0" F="([^"]*)"`).exec(xml)?.[1];
async function shapes(bytes: Uint8Array): Promise<Map<string, VisioShape>> {
	return new Map((await parseVsdx(bytes)).pages[0]!.shapes.map((shape) => [shape.id, shape]));
}
const box = (shape: VisioShape) => {
	const [a, , , d, e, f] = shape.transform;
	return { minX: e, maxX: e + a * shape.width, minY: f, maxY: f + d * shape.height };
};

describe('cross-functional flowchart in Visio structure', () => {
	it('writes the frame, Swimlane List, Phase List and lanes with Visio relationship formulas', async () => {
		const xml = await pageXml(await built());
		// Read from a cross-functional diagram Visio 16 built: kind 1 members, 4 member of, 2 list
		// members, 5 list member of, 11 the sibling list.
		expect(relationships(xml, '50')).toBe(
			'SUM(DEPENDSON(1,Sheet.55!SheetRef(),Sheet.54!SheetRef()))',
		);
		expect(relationships(xml, '54')).toBe(
			'SUM(DEPENDSON(4,Sheet.50!SheetRef()),DEPENDSON(2,Sheet.51!SheetRef(),Sheet.52!SheetRef(),Sheet.53!SheetRef()),DEPENDSON(11,Sheet.55!SheetRef()))',
		);
		expect(relationships(xml, '55')).toBe(
			'SUM(DEPENDSON(4,Sheet.50!SheetRef()),DEPENDSON(11,Sheet.54!SheetRef()))',
		);
		expect(relationships(xml, '52')).toBe(
			'SUM(DEPENDSON(5,Sheet.54!SheetRef()),DEPENDSON(1,Sheet.2!SheetRef()))',
		);
		const empty = await pageXml(
			await run(await built(), [{ type: 'delete-shape', pageId: '0', shapeId: '2' }]),
		);
		expect(relationships(empty, '52')).toBe('SUM(DEPENDSON(5,Sheet.54!SheetRef()),DEPENDSON(1))');
		const map = await shapes(await built());
		expect(map.get('54')!.structure).toMatchObject({
			type: 'list',
			kind: 'swimlane',
			direction: 2,
			memberIds: ['51', '52', '53'],
		});
		expect(map.get('55')!.structure).toMatchObject({ type: 'list', kind: 'phase', direction: 0 });
		expect(map.get('51')!.structure).toMatchObject({ swimlane: { role: 'lane' } });
	});

	it('keeps the lists spanning the lanes and the lane list in stacking order', async () => {
		const bytes = await run(await built(), [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '50',
				shapeId: '60',
				relativeTo: '51',
				position: 'before',
				heading: 'New',
			},
		]);
		const map = await shapes(bytes);
		expect(map.get('54')!.structure).toMatchObject({ memberIds: ['60', '51', '52', '53'] });
		const list = box(map.get('54')!);
		const frame = box(map.get('50')!);
		expect(list.minY).toBeCloseTo(frame.minY);
		expect(list.maxY).toBeCloseTo(frame.maxY - 0.35);
		const xml = await pageXml(bytes);
		expect(relationships(xml, '60')).toContain('DEPENDSON(5,Sheet.54!SheetRef())');
	});

	it('drops a deleted lane from the list and the back references', async () => {
		const xml = await pageXml(
			await run(await built(), [
				{ type: 'delete-swimlane', pageId: '0', frameId: '50', laneId: '52' },
			]),
		);
		expect(relationships(xml, '54')).toBe(
			'SUM(DEPENDSON(4,Sheet.50!SheetRef()),DEPENDSON(2,Sheet.51!SheetRef(),Sheet.53!SheetRef()),DEPENDSON(11,Sheet.55!SheetRef()))',
		);
	});

	it('names each phase: the bar carries its label and the Phase List its order', async () => {
		const bytes = await run(await built(), [
			{
				type: 'insert-separator',
				pageId: '0',
				frameId: '50',
				shapeId: '71',
				at: 8,
				label: 'Build',
			},
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '70', at: 5, label: 'Plan' },
		]);
		const map = await shapes(bytes);
		expect(map.get('70')!.text?.plainText).toBe('Plan');
		expect(map.get('71')!.text?.plainText).toBe('Build');
		expect(map.get('55')!.structure).toMatchObject({ memberIds: ['70', '71'] });
		expect(map.get('70')!.structure).toMatchObject({ swimlane: { role: 'separator' } });
		const unnamed = await shapes(
			await run(bytes, [
				{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '72', at: 7 },
			]),
		);
		expect(unnamed.get('72')!.text?.plainText).toBe('Phase 3');
	});

	it('turns a horizontal flowchart on its side and back, carrying the shapes', async () => {
		const start = await built();
		const before = await shapes(start);
		const turned = await run(start, [
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '70', at: 6 },
			{ type: 'set-swimlane-options', pageId: '0', frameId: '50', direction: 'vertical' },
		]);
		const map = await shapes(turned);
		expect(map.get('50')!.structure).toMatchObject({
			swimlane: { options: { direction: 'vertical' } },
		});
		const lanes = ['51', '52', '53'].map((id) => box(map.get(id)!));
		const old = ['51', '52', '53'].map((id) => box(before.get(id)!));
		const top = box(map.get('50')!).maxY - 0.35;
		lanes.forEach((lane, index) => {
			// Each lane keeps its size along the stacking axis: the old height is the new width.
			expect(lane.maxX - lane.minX).toBeCloseTo(old[index]!.maxY - old[index]!.minY);
			expect(lane.maxY).toBeCloseTo(top);
			if (index) expect(lane.minX).toBeCloseTo(lanes[index - 1]!.maxX);
		});
		expect(lanes[0]!.minX).toBeCloseTo(box(before.get('50')!).minX);
		expect(map.get('54')!.structure).toMatchObject({ direction: 0, memberIds: ['51', '52', '53'] });
		expect(map.get('55')!.structure).toMatchObject({ direction: 2 });
		// Shape 1 sat 0.7 inches below the top of lane 51 and 1.5 from its left: it turns with it.
		const one = box(map.get('1')!);
		const oldOne = box(before.get('1')!);
		const oldLane = old[0]!;
		const centre = { x: (one.minX + one.maxX) / 2, y: (one.minY + one.maxY) / 2 };
		const oldCentre = { x: (oldOne.minX + oldOne.maxX) / 2, y: (oldOne.minY + oldOne.maxY) / 2 };
		expect(centre.x - lanes[0]!.minX).toBeCloseTo(oldLane.maxY - oldCentre.y);
		expect(lanes[0]!.maxY - centre.y).toBeCloseTo(oldCentre.x - oldLane.minX);
		// The phase bar is now horizontal and spans the lanes.
		const bar = box(map.get('70')!);
		expect(bar.maxY - bar.minY).toBeCloseTo(0.04);
		expect(bar.minX).toBeCloseTo(lanes[0]!.minX);
		expect(bar.maxX).toBeCloseTo(lanes[2]!.maxX);

		const back = await shapes(
			await run(turned, [
				{ type: 'set-swimlane-options', pageId: '0', frameId: '50', direction: 'horizontal' },
			]),
		);
		for (const id of ['50', '51', '52', '53', '1', '2', '3']) {
			const a = box(back.get(id)!);
			const b = box(before.get(id)!);
			expect([a.minX, a.maxX, a.minY, a.maxY].map((v) => Number(v.toFixed(3)))).toEqual(
				[b.minX, b.maxX, b.minY, b.maxY].map((v) => Number(v.toFixed(3))),
			);
		}
	});

	it('brings a flowchart saved as plain containers over to lists when it is edited', async () => {
		const zip = await JSZip.loadAsync(await built());
		let xml = await zip.file('visio/pages/page1.xml')!.async('string');
		// Back to the earlier form: no lists, the lanes held by the frame under kind 4.
		xml = xml.replace(/<Shape ID="5[45]"[\s\S]*?<\/Shape>/g, '');
		xml = xml.replace(/<Cell N="Relationships"[^>]*F="SUM\(DEPENDSON\(5,[^"]*"\/>/g, '');
		xml = xml.replace(
			/(<Shape ID="50"[\s\S]*?N="Relationships" V="0" F=")[^"]*"/,
			'$1SUM(DEPENDSON(4,Sheet.51!SheetRef(),Sheet.52!SheetRef(),Sheet.53!SheetRef()))"',
		);
		zip.file('visio/pages/page1.xml', xml);
		const legacy = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await shapes(legacy);
		expect(loaded.get('51')!.structure).toMatchObject({ swimlane: { role: 'lane' } });
		const bytes = await run(legacy, [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '50',
				shapeId: '60',
				relativeTo: '53',
				position: 'after',
				heading: 'New',
			},
		]);
		const map = await shapes(bytes);
		const lists = [...map.values()].filter((shape) => shape.structure?.type === 'list');
		expect(lists.map((shape) => shape.structure)).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ kind: 'swimlane', memberIds: ['51', '52', '53', '60'] }),
				expect.objectContaining({ kind: 'phase' }),
			]),
		);
	});
});
