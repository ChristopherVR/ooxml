import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

for (const variable of ['VISIO_NATIVE_LINE_WIDTH_BEGIN_DIR', 'VISIO_NATIVE_LINE_WIDTH_END_DIR']) {
	const directory = process.env[variable];
	it.skipIf(!directory)(`records fixed-Width native endpoint semantics (${variable})`, async () => {
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: {
				shapeId: string;
				endpoint: 'Begin' | 'End';
				endpointBefore: Record<string, { value: number; formula: string }>;
				endpointAfter: Record<string, { value: number; formula: string }>;
				endpointTransform: number[];
			}[];
		};
		expect(evidence.cases).toHaveLength(4);
		const source = await readFile(join(directory!, 'resized.vsdx'));
		const original = await parseVsdx(source);
		const native = await parseVsdx(await readFile(join(directory!, 'endpoint.vsdx')));
		for (const item of evidence.cases) {
			const before = item.endpointBefore;
			const after = item.endpointAfter;
			expect(after.Width).toEqual(before.Width);
			expect(after.Width!.value).toBe(4);
			const opposite = item.endpoint === 'Begin' ? 'End' : 'Begin';
			for (const axis of ['X', 'Y'])
				expect(after[opposite + axis]!.value).toBe(before[opposite + axis]!.value);
			const shape = native.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
			for (let i = 0; i < 6; i++)
				expect(shape.transform[i]).toBeCloseTo(item.endpointTransform[i]!, 12);
			expect(shape.width).toBe(4);
			// Raw endpoint cells and displayed geometry diverge after an explicit Width override.
			const [a, b, , , x, y] = shape.transform;
			expect(Math.hypot(x - after.BeginX!.value, y - after.BeginY!.value)).toBeGreaterThan(0.1);
			expect(
				Math.hypot(
					x + a * shape.width - after.EndX!.value,
					y + b * shape.width - after.EndY!.value,
				),
			).toBeGreaterThan(0.1);
			await expect(
				editVsdx(source, [
					{
						type: 'move-line-endpoint',
						pageId: original.pages[0]!.id,
						shapeId: item.shapeId,
						endpoint: item.endpoint === 'Begin' ? 'begin' : 'end',
						x: after[item.endpoint + 'X']!.value,
						y: after[item.endpoint + 'Y']!.value,
					},
				]),
			).rejects.toThrow('Endpoint editing requires native derived transform formulas');
		}
		expect(source).toEqual(await readFile(join(directory!, 'resized.vsdx')));
	});
}
