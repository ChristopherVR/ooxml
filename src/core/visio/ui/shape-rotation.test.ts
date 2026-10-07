import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from '../parser';
import { editVsdx } from '../edit';
import { demoDocument } from './demo-document';
import { visioLocalRotationShape, visioQuarterTurnCommand } from './shape-rotation';
it('shares local rotation admission and preserves the sign at half-turn boundaries', () => {
	const page = structuredClone(demoDocument.pages[0]!);
	const shape = page.shapes[0]!;
	shape.rotation = { pinX: 3, pinY: 5, angle: Math.PI / 2 };
	expect(visioQuarterTurnCommand(page, shape.id, 'left')!.angle).toBe(Math.PI);
	shape.rotation.angle = -Math.PI / 2;
	expect(visioQuarterTurnCommand(page, shape.id, 'right')!.angle).toBe(-Math.PI);
	shape.rotation.angle = Math.PI / 2 + 4e-15;
	expect(visioQuarterTurnCommand(page, shape.id, 'left')!.angle).toBe(Math.PI);
	shape.rotation.angle = Math.PI / 2 + 1e-7;
	expect(visioQuarterTurnCommand(page, shape.id, 'left')!.angle).toBeLessThan(0);
	expect(visioLocalRotationShape(page, shape.id)).toBe(shape);
	for (const unsupported of [
		{ kind: 'connector' as const },
		{ kind: 'group' as const },
		{ masterId: '0' },
		{ hidden: true },
		{ width: 0 },
		{ height: 0 },
		{ children: [structuredClone(shape)] },
	]) {
		const target = { ...shape, ...unsupported };
		const scene = { ...page, shapes: [target] };
		expect(visioLocalRotationShape(scene, shape.id)).toBeUndefined();
		expect(visioQuarterTurnCommand(scene, shape.id, 'left')).toBeUndefined();
	}
	expect(
		visioLocalRotationShape(
			{
				...page,
				connectors: [{ fromShapeId: shape.id, toShapeId: 'x', fromCell: 'BeginX', toCell: 'PinX' }],
			},
			shape.id,
		),
	).toBeUndefined();
	delete shape.rotation;
	expect(visioQuarterTurnCommand(page, shape.id, 'left')).toBeUndefined();
});
for (const variable of [
	'VISIO_NATIVE_QUARTER_LEFT_DIR',
	'VISIO_NATIVE_QUARTER_RIGHT_PIN_DIR',
	'VISIO_NATIVE_QUARTER_WRAP_DIR',
	'VISIO_NATIVE_QUARTER_HALF_LEFT_DIR',
	'VISIO_NATIVE_QUARTER_HALF_RIGHT_DIR',
	'VISIO_NATIVE_QUARTER_DEPENDENT_RIGHT_DIR',
	'VISIO_NATIVE_QUARTER_DEPENDENT_WRAP_DIR',
]) {
	it.skipIf(!process.env[variable]).each(['rectangle', 'ellipse'])(
		`matches native Selection.Rotate ${variable} %s`,
		async (kind) => {
			const directory = process.env[variable]!;
			const source = await readFile(join(directory, 'quarter-source.vsdx'));
			const model = await parseVsdx(source),
				page = model.pages[0]!;
			const native = await parseVsdx(await readFile(join(directory, 'quarter-turned.vsdx')));
			const evidence = JSON.parse(await readFile(join(directory, 'evidence.json'), 'utf8'));
			const reference = evidence.quarterTurned[kind];
			const command = visioQuarterTurnCommand(
				page,
				reference.shapeId,
				evidence.quarterTurn === 'Left' ? 'left' : 'right',
			)!;
			expect(command.angle).toBeCloseTo(reference.cells.Angle.value, 12);
			const saved = await editVsdx(source, [command]);
			const actual = (await parseVsdx(saved.bytes)).pages[0]!.shapes.find(
				(shape) => shape.id === reference.shapeId,
			)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
			expect(actual.rotation!.angle).toBeCloseTo(reference.cells.Angle.value, 12);
			expect(actual.rotation!.pinX).toBeCloseTo(expected.rotation!.pinX, 12);
			expect(actual.rotation!.pinY).toBeCloseTo(expected.rotation!.pinY, 12);
			const ratio = native.pages[0]!.drawingToPageScale ?? 1;
			for (let i = 0; i < 6; i++) {
				expect(actual.transform[i]).toBeCloseTo(expected.transform[i]!, 12);
				expect(actual.transform[i]).toBeCloseTo(reference.transform[i] * (i >= 4 ? ratio : 1), 12);
			}
		},
	);
}
