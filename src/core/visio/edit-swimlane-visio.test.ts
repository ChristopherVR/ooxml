import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { visioSwimlaneInfo } from './ui/swimlane';

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

	it('is left unedited by the lane commands: its lanes are master groups', async () => {
		const bytes = new Uint8Array(
			readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
		);
		await expect(
			editVsdx(bytes, [
				{
					type: 'insert-swimlane',
					pageId: '0',
					frameId: '1',
					shapeId: '30',
					relativeTo: '6',
					position: 'after',
					heading: 'New',
				},
			]),
		).rejects.toThrow();
	});

	it('plans edits on it: the lanes come from the Swimlane List, top lane first', async () => {
		const page = await native();
		const info = visioSwimlaneInfo(page, '6');
		expect(info?.frame.id).toBe('1');
		expect(info?.lanes.map((lane) => lane.id)).toEqual(['9', '6']);
		expect(visioSwimlaneInfo(page, '12')?.frame.id).toBe('1');
	});
});
