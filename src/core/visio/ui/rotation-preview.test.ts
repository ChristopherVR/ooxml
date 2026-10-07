import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from '../parser';
import { transform } from '../geometry';
import { demoDocument } from './demo-document';
import { visioRotationPreviewTransform } from './rotation-preview';
it.each([false, true])(
	'preserves a custom pivot and flip %s through signed preview turns',
	(flip) => {
		const source = structuredClone(demoDocument.pages[0]!.shapes[0]!);
		source.rotation = { pinX: 3, pinY: 5, angle: Math.PI / 6 };
		source.transform = transform(3, 5, 0.3, 0.7, Math.PI / 6, flip, !flip);
		const before = structuredClone(source);
		for (const angle of [-Math.PI / 4, Math.PI / 2, (7 * Math.PI) / 6]) {
			const actual = visioRotationPreviewTransform(source, angle)!;
			const expected = transform(3, 5, 0.3, 0.7, angle, flip, !flip);
			for (let i = 0; i < 6; i++) expect(actual[i]).toBeCloseTo(expected[i]!, 12);
		}
		expect(source).toEqual(before);
		expect(visioRotationPreviewTransform(source, source.rotation.angle)).toBe(source.transform);
		expect(visioRotationPreviewTransform(source, Number.NaN)).toBeUndefined();
		for (const field of ['pinX', 'pinY', 'angle'] as const) {
			const malformed = { ...source, rotation: { ...source.rotation, [field]: Number.NaN } };
			expect(visioRotationPreviewTransform(malformed, 0)).toBeUndefined();
		}
		delete source.rotation;
		expect(visioRotationPreviewTransform(source, 0)).toBeUndefined();
	},
);
for (const variable of [
	'VISIO_NATIVE_ROTATE_DIR',
	'VISIO_NATIVE_ROTATE_DOUBLE_DIR',
	'VISIO_NATIVE_ROTATE_HALF_DIR',
	'VISIO_NATIVE_ROTATE_TRIPLE_DIR',
	'VISIO_NATIVE_ROTATE_PIN_DIR',
]) {
	it.skipIf(!process.env[variable]).each(['rectangle', 'ellipse'])(
		`matches native preview pose ${variable} %s`,
		async (kind) => {
			const directory = process.env[variable];
			if (!directory) return;
			const source = await parseVsdx(
				await readFile(
					join(
						directory,
						variable === 'VISIO_NATIVE_ROTATE_PIN_DIR'
							? 'rotation-source.vsdx'
							: 'ellipse-edited.vsdx',
					),
				),
			);
			const native = await parseVsdx(await readFile(join(directory, 'rotated.vsdx')));
			const evidence = JSON.parse(await readFile(join(directory, 'evidence.json'), 'utf8'));
			const reference = evidence.rotated[kind];
			const shape = source.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
			const preview = visioRotationPreviewTransform(shape, reference.cells.Angle.value)!;
			for (let i = 0; i < 6; i++) expect(preview[i]).toBeCloseTo(expected.transform[i]!, 12);
		},
	);
}
