import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';
import { loadCffMasters } from './cff-masters';
import { openEditablePackage } from './edit-package';
import { DEFAULTS } from './package-common';

const rect = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width: 1,
	height: 0.6,
});
const create: VisioEdit = {
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
};
const run = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
async function built() {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	const start = await run(blank, [rect('1', 4, 6.8), rect('2', 6, 4.6), rect('3', 8, 2.7)]);
	return run(start, [create, { type: 'reconcile-containers', pageId: '0', shapeIds: ['1', '2'] }]);
}
const text = async (bytes: Uint8Array, path: string) =>
	await (await JSZip.loadAsync(bytes)).file(path)!.async('string');
const masterSet = async (bytes: Uint8Array) =>
	await loadCffMasters((await openEditablePackage(bytes, DEFAULTS, () => undefined)).pkg);
/** Shape id to its Master attribute, from the saved page. */
async function masterOf(bytes: Uint8Array): Promise<Map<string, string | null>> {
	const xml = await text(bytes, 'visio/pages/page1.xml');
	const map = new Map<string, string | null>();
	for (const match of xml.matchAll(/<Shape ID="(\d+)"([^>]*)>/g))
		map.set(match[1]!, /Master="(\d+)"/.exec(match[2]!)?.[1] ?? null);
	return map;
}

describe('cross-functional flowchart as instances of the flowchart masters', () => {
	it('adds the five hidden masters and makes every part of the chart an instance', async () => {
		const bytes = await built();
		const masters = await text(bytes, 'visio/masters/masters.xml');
		for (const name of ['CFF Container', 'Swimlane List', 'Phase List', 'Swimlane', 'Separator'])
			expect(masters).toMatch(new RegExp(`NameU="${name}"[^>]*Hidden="1"`));
		const ids = await masterOf(bytes);
		const set = await masterSet(bytes);
		expect(set.empty).toBe(true);
		expect(ids.get('50')).toBe(set.authoredId('frame'));
		expect(ids.get('54')).toBe(set.authoredId('swimlane-list'));
		expect(ids.get('55')).toBe(set.authoredId('phase-list'));
		for (const lane of ['51', '52', '53']) expect(ids.get(lane)).toBe(set.authoredId('lane'));
		expect(ids.get('1')).toBeNull();
		const parsed = await parseVsdx(bytes);
		expect(parsed.pages[0]!.shapes.find((shape) => shape.id === '50')!.structure).toMatchObject({
			type: 'container',
			swimlane: { role: 'frame' },
		});
	});

	it('keeps new lanes and separators instances, and moves the frame', async () => {
		const edited = await run(await built(), [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '50',
				shapeId: '60',
				relativeTo: '53',
				position: 'after',
				heading: 'D',
			},
			{ type: 'insert-separator', pageId: '0', frameId: '50', shapeId: '61' },
		]);
		const set = await masterSet(edited);
		const ids = await masterOf(edited);
		expect(ids.get('60')).toBe(set.authoredId('lane'));
		expect(ids.get('61')).toBe(set.authoredId('separator'));
		const moved = await run(edited, [
			{ type: 'move-shape', pageId: '0', shapeId: '50', x: 6.5, y: 5 },
		]);
		expect((await masterOf(moved)).get('50')).toBe(set.authoredId('frame'));
		const again = await text(moved, 'visio/masters/masters.xml');
		expect(again.match(/<Master /g)).toHaveLength(5);
	});
});
