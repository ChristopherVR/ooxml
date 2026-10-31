import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { visioSwimlaneInfo, visioSwimlaneOptions } from './ui/swimlane';

/** The cross-functional diagram Visio 16 built from its own template (see the .md next to it). */
const native = async () =>
	(
		await parseVsdx(
			new Uint8Array(
				readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
			),
		)
	).pages[0]!;

describe('a cross-functional flowchart Visio built', () => {
	it('is recognised: frame, Swimlane List with its lanes in order, Phase List with its separator', async () => {
		const page = await native();
		const byId = new Map(page.shapes.map((shape) => [shape.id, shape]));
		expect(byId.get('1')!.structure).toMatchObject({
			type: 'container',
			memberIds: ['5', '4'],
			swimlane: { role: 'frame', options: { direction: 'horizontal' } },
		});
		expect(byId.get('4')!.structure).toMatchObject({
			type: 'list',
			kind: 'swimlane',
			direction: 2,
			memberIds: ['9', '6'],
		});
		expect(byId.get('5')!.structure).toMatchObject({
			type: 'list',
			kind: 'phase',
			direction: 0,
			memberIds: ['12'],
		});
		expect(byId.get('6')!.structure).toMatchObject({ swimlane: { role: 'lane' } });
		expect(byId.get('12')!.structure).toMatchObject({ swimlane: { role: 'separator' } });
	});

	it('takes lane edits: a new lane is a copy of its neighbour, the chart restacks', async () => {
		const bytes = new Uint8Array(
			readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
		);
		const result = await editVsdx(bytes, [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '1',
				shapeId: '30',
				relativeTo: '6',
				position: 'after',
				heading: 'New',
			},
		]);
		const page = (await parseVsdx(result.bytes)).pages[0]!;
		const byId = new Map(page.shapes.map((shape) => [shape.id, shape]));
		expect(byId.get('4')!.structure).toMatchObject({ type: 'list', memberIds: ['9', '6', '30'] });
		expect(byId.get('30')!.masterId).toBe(byId.get('6')!.masterId);
		const info = visioSwimlaneInfo(page, '30');
		expect(info?.lanes.map((lane) => lane.id)).toEqual(['9', '6', '30']);
		// Lanes edge to edge, the frame grown by the new lane's height.
		const tops = info!.lanes.map((lane) => lane.transform[5] + lane.height);
		const heights = info!.lanes.map((lane) => lane.height);
		expect(tops[1]! - heights[1]!).toBeCloseTo(tops[2]!, 5);
		expect(tops[0]! - heights[0]!).toBeCloseTo(tops[1]!, 5);
	});

	it('resizes and deletes lanes of a chart Visio built', async () => {
		const bytes = new Uint8Array(
			readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
		);
		const before = (await parseVsdx(bytes)).pages[0]!;
		const frameHeight = before.shapes.find((shape) => shape.id === '1')!.height;
		const laneHeight = before.shapes.find((shape) => shape.id === '6')!.height;
		const grown = await editVsdx(bytes, [
			{ type: 'resize-swimlane', pageId: '0', frameId: '1', laneId: '6', size: laneHeight + 1 },
		]);
		const page = (await parseVsdx(grown.bytes)).pages[0]!;
		const byId = new Map(page.shapes.map((shape) => [shape.id, shape]));
		expect(byId.get('6')!.height).toBeCloseTo(laneHeight + 1, 5);
		expect(byId.get('1')!.height).toBeCloseTo(frameHeight + 1, 5);
		expect(byId.get('4')!.height).toBeGreaterThan(before.shapes.find((s) => s.id === '4')!.height);
		const removed = await editVsdx(bytes, [
			{ type: 'delete-swimlane', pageId: '0', frameId: '1', laneId: '9' },
		]);
		const after = (await parseVsdx(removed.bytes)).pages[0]!;
		expect(after.shapes.some((shape) => shape.id === '9')).toBe(false);
		expect(after.shapes.find((shape) => shape.id === '4')!.structure).toMatchObject({
			memberIds: ['6'],
		});
		expect(after.shapes.find((shape) => shape.id === '1')!.height).toBeLessThan(frameHeight);
	});

	it('takes a phase and an option (the cell-by-cell checks are in edit-swimlane-cff.test.ts)', async () => {
		const bytes = new Uint8Array(
			readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
		);
		const result = await editVsdx(bytes, [
			{ type: 'insert-separator', pageId: '0', frameId: '1', shapeId: '30', at: 3 },
			{ type: 'set-swimlane-options', pageId: '0', frameId: '1', style: 2 },
		]);
		const page = (await parseVsdx(result.bytes)).pages[0]!;
		expect(page.shapes.find((shape) => shape.id === '5')!.structure).toMatchObject({
			memberIds: ['30', '12'],
		});
		expect(page.shapes.find((shape) => shape.id === '1')!.structure).toMatchObject({
			swimlane: { options: { style: 2 } },
		});
	});

	it('is marked native, and its option edits move no lane shapes of their own', async () => {
		const page = await native();
		const frame = page.shapes.find((shape) => shape.id === '1')!;
		expect(frame.structure).toMatchObject({ swimlane: { role: 'frame', native: true } });
		for (const change of [{ showTitle: false }, { showSeparators: false }, { style: 2 }]) {
			const edits = visioSwimlaneOptions(page, '1', change);
			expect(edits?.map((edit) => edit.type)).toEqual(['set-swimlane-options']);
		}
	});

	it('plans edits on it: the lanes come from the Swimlane List, top lane first', async () => {
		const page = await native();
		const info = visioSwimlaneInfo(page, '6');
		expect(info?.frame.id).toBe('1');
		expect(info?.lanes.map((lane) => lane.id)).toEqual(['9', '6']);
		expect(visioSwimlaneInfo(page, '12')?.frame.id).toBe('1');
	});
});
