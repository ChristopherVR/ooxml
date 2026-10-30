import { expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from '../edit';
import { parseVsdx } from '../parser';
import { cell, fixture, rectangle, shape } from '../test-fixtures';
import type { VisioShape } from '../model';
import { visioResizeSelectionDrag } from './selection-resize';

const box = (shapeId: string, x: number, y: number, width = 1, height = 1): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width,
	height,
});
async function base(commands: VisioEdit[]): Promise<Uint8Array> {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	return (await editVsdx(blank, commands)).bytes;
}
const page = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!;
const find = (shapes: readonly VisioShape[], id: string): VisioShape =>
	shapes.flatMap((shape) => [shape, ...find2(shape)]).find((shape) => shape.id === id)!;
const find2 = (shape: VisioShape): VisioShape[] =>
	shape.children.flatMap((child) => [child, ...find2(child)]);

/** Drag the east handle of the selection by `dx` page inches and commit it. */
async function resizeEast(bytes: Uint8Array, ids: string[], dx: number, handle: 'e' | 'ne' = 'e') {
	const before = await page(bytes);
	const result = visioResizeSelectionDrag(before, ids, handle, { x: 0, y: 0 }, { x: dx, y: 0 })!;
	expect(result).toBeDefined();
	const saved = await editVsdx(bytes, result.commands);
	return { before, result, after: await page(saved.bytes) };
}

it('resizes a rotated shape by the squared cosines of its axes and keeps its angle', async () => {
	const angle = Math.PI / 6;
	const bytes = await base([
		box('1', 2, 2, 1, 2),
		box('2', 6, 3),
		{ type: 'rotate-shape', pageId: '0', shapeId: '1', angle },
	]);
	const { before, after, result } = await resizeEast(bytes, ['1', '2'], 2, 'e');
	// The east handle only scales x: the selection spans the rotated shape's corners to the box.
	const left = Math.min(...[before.shapes[0]!, before.shapes[1]!].map((s) => bounds(s).left));
	const right = Math.max(...[before.shapes[0]!, before.shapes[1]!].map((s) => bounds(s).right));
	const sx = (right - left + 2) / (right - left);
	const rotated = find(after.shapes, '1'),
		original = find(before.shapes, '1');
	expect(rotated.width).toBeCloseTo(
		original.width * (Math.cos(angle) ** 2 * sx + Math.sin(angle) ** 2),
		9,
	);
	expect(rotated.height).toBeCloseTo(
		original.height * (Math.sin(angle) ** 2 * sx + Math.cos(angle) ** 2),
		9,
	);
	expect(rotated.rotation!.pinX).toBeCloseTo(left + (original.rotation!.pinX - left) * sx, 9);
	expect(rotated.rotation!.pinY).toBeCloseTo(original.rotation!.pinY, 9);
	expect(rotated.transform[0]).toBeCloseTo(original.transform[0], 12);
	expect(rotated.transform[1]).toBeCloseTo(original.transform[1], 12);
	expect(result.members.map((member) => member.id)).toEqual(['1', '2']);
});

function bounds(shape: VisioShape) {
	const [a, , c, , e] = shape.transform;
	const xs: number[] = [];
	for (const [lx, ly] of [
		[0, 0],
		[shape.width, 0],
		[0, shape.height],
		[shape.width, shape.height],
	] as const) {
		xs.push(a * lx + c * ly + e);
	}
	return { left: Math.min(...xs), right: Math.max(...xs) };
}

it('resizes a group beside a shape: members follow the group and keep their proportions', async () => {
	const bytes = await base([
		box('1', 2, 2),
		box('2', 3.5, 3),
		box('3', 8, 2),
		{ type: 'group-shapes', pageId: '0', shapeId: '9', memberIds: ['1', '2'] },
	]);
	const { before, after } = await resizeEast(bytes, ['9', '3'], 5, 'e');
	const group = find(before.shapes, '9');
	const left = group.transform[4],
		right = find(before.shapes, '3').transform[4] + 1;
	const sx = (right - left + 5) / (right - left);
	const grown = find(after.shapes, '9');
	expect(grown.width).toBeCloseTo(group.width * sx, 9);
	for (const id of ['1', '2']) {
		const old = find(before.shapes, id),
			next = find(after.shapes, id);
		expect(next.width).toBeCloseTo(old.width * sx, 9);
		expect(next.height).toBeCloseTo(old.height, 9);
	}
	expect(find(after.shapes, '3').width).toBeCloseTo(sx, 9);
});

it('reroutes connectors glued to resized selection members', async () => {
	const blank = await base([box('1', 2, 2), box('2', 5, 2)]);
	const bytes = (
		await editVsdx(blank, [
			{
				type: 'create-line',
				pageId: '0',
				shapeId: '3',
				beginX: 0,
				beginY: 0,
				endX: 9,
				endY: 9,
				connect: { begin: '1', end: '2' },
			},
		])
	).bytes;
	const { before, after } = await resizeEast(bytes, ['1', '2'], 4, 'e');
	const sx = 2;
	const right = find(after.shapes, '1'),
		left = find(after.shapes, '2');
	expect(right.width).toBeCloseTo(find(before.shapes, '1').width * sx, 9);
	const line = find(after.shapes, '3');
	// The connector still spans from the first shape's edge to the second shape's edge.
	const lineLeft = line.transform[4],
		lineRight = line.transform[4] + line.width;
	expect(lineLeft).toBeCloseTo(right.rotation!.pinX + right.width / 2, 9);
	expect(lineRight).toBeCloseTo(left.rotation!.pinX - left.width / 2, 9);
});

const layered = async (locked: boolean) => {
	const member = (id: string, x: number) =>
		shape(
			id,
			cell('Width', 1) +
				cell('Height', 1) +
				cell('PinX', x) +
				cell('PinY', 2) +
				cell('LocPinX', 0.5) +
				cell('LocPinY', 0.5) +
				cell('LayerMember', '0') +
				rectangle,
		);
	return fixture({
		pages: [
			{
				id: '0',
				pageCells: `<Section N="Layer" IX="0"><Row IX="0">${cell('Name', 'Walls')}${cell('Lock', locked ? 1 : 0)}</Row></Section>`,
				contents: `<Shapes>${member('1', 2)}${member('2', 5)}</Shapes>`,
			},
		],
	});
};

it.each([false, true])(
	'resizes layered shapes unless their layer is locked (locked %s)',
	async (locked) => {
		const bytes = await layered(locked);
		const before = await page(bytes);
		const result = visioResizeSelectionDrag(
			before,
			['1', '2'],
			'e',
			{ x: 0, y: 0 },
			{ x: 3, y: 0 },
		);
		if (locked) {
			expect(result).toBeUndefined();
			// The core refuses the same edit for callers that skip the interface check.
			await expect(
				editVsdx(bytes, [
					{
						type: 'resize-shape',
						pageId: '0',
						shapeId: '1',
						width: 2,
						height: 1,
						anchor: { x: 0, y: 0 },
					},
				]),
			).rejects.toThrow(/locked layer/);
			return;
		}
		expect(result).toBeDefined();
		const after = await page((await editVsdx(bytes, result!.commands)).bytes);
		expect(find(after.shapes, '1').width).toBeGreaterThan(1);
		expect(find(after.shapes, '2').width).toBeGreaterThan(1);
	},
);
