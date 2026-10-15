import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { VisioPackage } from './package';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';
import { attribute } from './sheet';
import { connectorPolyline } from './edit-connector-follow';
import { routeCrossings } from './connector-route-avoid';

const box = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width: 1,
	height: 1,
});
const connect = (shapeId: string, begin: string, end: string): VisioEdit => ({
	type: 'create-line',
	pageId: '0',
	shapeId,
	beginX: 0,
	beginY: 0,
	endX: 9,
	endY: 9,
	connect: { begin, end },
	route: 'right-angle',
});
const move = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId,
	x,
	y,
});
async function route(bytes: Uint8Array, id: string) {
	const root = (await VisioPackage.open(bytes)).readXml('visio/pages/page1.xml');
	const shape = Array.from((await root).getElementsByTagName('Shape')).find(
		(node) => attribute(node, 'ID') === id,
	)!;
	return connectorPolyline(shape)!;
}
const around = (x: number, y: number) => ({
	minX: x - 0.5,
	minY: y - 0.5,
	maxX: x + 0.5,
	maxY: y + 0.5,
});

/** 1 and 2 joined by connector 3; 4 and 5 joined by connector 6, which makes both placeable. */
async function drawing(): Promise<Uint8Array> {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	const shapes = await editVsdx(blank, [
		box('1', 2, 4),
		box('2', 8, 4),
		box('4', 5, 9),
		box('5', 8, 9),
		box('7', 5, 7),
	]);
	const glued = await editVsdx(shapes.bytes, [connect('6', '4', '5')]);
	return (await editVsdx(glued.bytes, [connect('3', '1', '2')])).bytes;
}

describe('a shape moved onto a right-angle connector', () => {
	it('reroutes the connector around a placeable shape', async () => {
		const start = await drawing();
		expect(routeCrossings(await route(start, '3'), [around(5, 4)])).toBe(1);
		const moved = await editVsdx(start, [move('4', 5, 4)]);
		const after = await route(moved.bytes, '3');
		expect(routeCrossings(after, [around(5, 4)])).toBe(0);
		// Still joined to its shapes at both ends.
		expect(after[0]!.x).toBeLessThan(2.6);
		expect(after.at(-1)!.x).toBeGreaterThan(7.4);
		expect(moved.changedParts).toContain('visio/pages/page1.xml');
	});

	it('also does so when the shape is resized onto the path', async () => {
		const start = await drawing();
		const resized = await editVsdx(start, [
			{ type: 'resize-shape', pageId: '0', shapeId: '4', width: 1, height: 6 },
		]);
		const model = await parseVsdx(resized.bytes);
		const four = model.pages[0]!.shapes.find((item) => item.id === '4')!;
		const { pinX, pinY } = four.rotation!;
		expect(four.height).toBe(6);
		const area = {
			minX: pinX - four.width / 2,
			minY: pinY - four.height / 2,
			maxX: pinX + four.width / 2,
			maxY: pinY + four.height / 2,
		};
		expect(routeCrossings(await route(resized.bytes, '3'), [area])).toBe(0);
	});

	it('leaves the connector alone for a shape that is not placeable', async () => {
		const start = await drawing();
		const before = await route(start, '3');
		const moved = await editVsdx(start, [move('7', 5, 4)]);
		expect(await route(moved.bytes, '3')).toEqual(before);
	});

	it('leaves connectors the shape does not touch as they were', async () => {
		const start = await drawing();
		const before = await route(start, '3');
		const moved = await editVsdx(start, [move('4', 5, 10)]);
		expect(await route(moved.bytes, '3')).toEqual(before);
	});

	it('does not reroute the shape connector glued to the moved shape onto itself', async () => {
		const start = await drawing();
		const moved = await editVsdx(start, [move('5', 8, 4)]);
		// 6 follows 5 as before; 3 runs 1 to 2 through 5's new place, so it goes around.
		expect(routeCrossings(await route(moved.bytes, '3'), [around(8, 4)])).toBeLessThanOrEqual(1);
	});
});
