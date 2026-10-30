import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { visioGlueSites } from './edit-connector-glue';
import { transform } from './geometry';

/**
 * Visio 16 (COM) drew a group G (4) of two rectangles (1 at 2..3, 2 at 4..5, y 3..4), a free
 * rectangle 3 at 8..9 x 6..7 and a Dynamic connector 5 glued from member 1 to shape 3. It then
 * set the group's PinX +3, its Width x1.5 and Height x2 (about the centre pin) and its Angle to
 * 0.5 rad, reading the connector's BeginX/BeginY/EndX/EndY back after each step.
 */
const source = new Uint8Array(
	readFileSync(new URL('./__fixtures__/group-glued-connector.vsdx', import.meta.url)),
);
const visio = {
	start: [2.5, 4, 8, 6.5],
	moved: [5.5, 4, 8, 6.5],
	resized: [5.00000000000002, 4.5, 8, 6.5],
	rotated: [5.18362615716445, 3.92035561941825, 8, 6.5],
};

async function ends(bytes: Uint8Array): Promise<number[]> {
	const xml = await (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
	const shape = /<Shape ID=['"]5['"][\s\S]*?<\/Shape>/.exec(xml)![0];
	return ['BeginX', 'BeginY', 'EndX', 'EndY'].map((name) =>
		Number(new RegExp(`<Cell N=['"]${name}['"] V=['"]([^'"]*)['"]`).exec(shape)![1]),
	);
}
const expectEnds = (actual: number[], expected: number[]) =>
	expected.forEach((value, index) => expect(actual[index], `end ${index}`).toBeCloseTo(value, 6));

describe('a connector glued to a group member follows the group (Visio 16)', () => {
	const run = (bytes: Uint8Array, edit: VisioEdit) => editVsdx(bytes, [edit]).then((r) => r.bytes);
	it('reads the member through its group at the start', async () => {
		expectEnds(await ends(source), visio.start);
	});
	it('moves with the group', async () => {
		const moved = await run(source, {
			type: 'move-shape',
			pageId: '0',
			shapeId: '4',
			x: 6.5,
			y: 3.5,
		});
		expectEnds(await ends(moved), visio.moved);
	});
	it('follows a resize and a rotation of the group', async () => {
		const moved = await run(source, {
			type: 'move-shape',
			pageId: '0',
			shapeId: '4',
			x: 6.5,
			y: 3.5,
		});
		const resized = await run(moved, {
			type: 'resize-shape',
			pageId: '0',
			shapeId: '4',
			width: 4.5,
			height: 2,
			anchor: { x: 0.5, y: 0.5 },
		});
		expectEnds(await ends(resized), visio.resized);
		const rotated = await run(resized, {
			type: 'rotate-shape',
			pageId: '0',
			shapeId: '4',
			angle: 0.5,
		});
		expectEnds(await ends(rotated), visio.rotated);
	});
});

describe('dynamic glue sites of a turned shape (Visio 16)', () => {
	it('sit where the lines through the centre cross the outline', () => {
		// A 1.5 x 2 rectangle at pin (2.75, 4) turned 0.5 rad; Visio read the glued ends back.
		const sites = visioGlueSites({
			width: 1.5,
			height: 2,
			transform: transform(2.75, 4, 0.75, 1, 0.5),
		});
		expect(sites[2]!.x).toBeCloseTo(2.75, 6);
		expect(sites[2]!.y).toBeCloseTo(5.13949392732455, 6);
		expect(sites[1]!.x).toBeCloseTo(3.60462044549341, 6);
		expect(sites[1]!.y).toBeCloseTo(4, 6);
		expect(sites[0]!.x).toBeCloseTo(2.75, 6);
		expect(sites[0]!.y).toBeCloseTo(2.86050607267545, 6);
	});
});
