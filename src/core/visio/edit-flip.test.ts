import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture, cell, shape, rectangle } from './test-fixtures';
import { snapshotEdits } from './ui/edit-commands';
const flip = (
	axis: 'horizontal' | 'vertical' = 'horizontal',
): Extract<VisioEdit, { type: 'flip-shape' }> => ({
	type: 'flip-shape',
	pageId: '0',
	shapeId: '1',
	axis,
});
const target = (extra = '') =>
	shape(
		'1',
		cell('Width', 2) + cell('Height', 1) + cell('PinX', 2) + cell('PinY', 3) + rectangle + extra,
	);
it('flips twice, preserves the source pin and snapshots its axis', async () => {
	const bytes = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${target(cell('Angle', Math.PI / 6))}</Shapes>` }],
	});
	const before = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
	for (const axis of ['horizontal', 'vertical'] as const) {
		const first = await editVsdx(bytes, [flip(axis)]);
		const result = (await parseVsdx(first.bytes)).pages[0]!.shapes[0]!;
		expect(result.rotation!.angle).toBeCloseTo(-Math.PI / 6, 12);
		expect(result.rotation!.pinX).toBe(before.rotation!.pinX);
		expect(result.rotation!.pinY).toBe(before.rotation!.pinY);
		expect(result.geometry).toEqual(before.geometry);
		const second = (await parseVsdx((await editVsdx(first.bytes, [flip(axis)])).bytes)).pages[0]!
			.shapes[0]!;
		expect(second.transform).toEqual(before.transform);
	}
	const command = flip();
	const copied = snapshotEdits([command]);
	command.axis = 'vertical';
	expect(copied).toEqual([flip()]);
});
for (const extra of [
	cell('FlipX', 0, 'GUARD(FALSE)'),
	cell('FlipX', 0, 'Width/1in'),
	cell('FlipX', 2),
	'<Cell N="FlipX" V="0" U="IN"/>',
	cell('Angle', 0, 'GUARD(0rad)'),
	cell('LockRotate', 1),
	cell('FlipX', 0) + cell('User.Test', 0, 'NOW()+FlipX'),
])
	it(`refuses unsupported source ${extra}`, async () => {
		const bytes = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${target(extra)}</Shapes>` }],
		});
		await expect(editVsdx(bytes, [flip()])).rejects.toThrow();
	});
for (const variable of ['VISIO_NATIVE_FLIP_HORIZONTAL_DIR', 'VISIO_NATIVE_FLIP_VERTICAL_DIR'])
	it.skipIf(!process.env[variable]).each(['rectangle', 'ellipse'])(
		`matches native ${variable} %s`,
		async (kind) => {
			const directory = process.env[variable]!;
			const evidence = JSON.parse(await readFile(join(directory, 'evidence.json'), 'utf8'));
			const id = evidence.flipped[kind].shapeId;
			const source = await readFile(join(directory, 'flip-source.vsdx'));
			const expected = (
				await parseVsdx(await readFile(join(directory, 'flipped.vsdx')))
			).pages[0]!.shapes.find((shape) => shape.id === id)!;
			const result = await editVsdx(source, [
				{ type: 'flip-shape', pageId: '0', shapeId: id, axis: evidence.flip.toLowerCase() },
			]);
			const actual = (await parseVsdx(result.bytes)).pages[0]!.shapes.find(
				(shape) => shape.id === id,
			)!;
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
			expect(actual.rotation).toEqual(expected.rotation);
			for (let i = 0; i < 6; i++)
				expect(actual.transform[i]).toBeCloseTo(expected.transform[i]!, 12);
		},
	);
