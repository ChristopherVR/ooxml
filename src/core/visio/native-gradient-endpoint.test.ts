import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { parseVsdx } from './parser';

for (const name of [
	'VISIO_NATIVE_GRADIENT_BEGIN_ENDPOINT_DIR',
	'VISIO_NATIVE_GRADIENT_END_ENDPOINT_DIR',
]) {
	const directory = process.env[name];
	it.skipIf(!directory)(`matches native gradient endpoint edits (${name})`, async () => {
		const bytes = await readFile(join(directory!, 'gradient-raster-before.vsdx'));
		const original = await parseVsdx(bytes);
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).replace(/^\uFEFF/, ''),
		) as {
			cases: {
				shapeId: string;
				endpointEdit: { endpoint: 'begin' | 'end'; x: number; y: number };
				nativeTransform: number[];
				nativeShapeWidth: number;
			}[];
		};
		expect(evidence.cases).toHaveLength(4);
		const saved = await editVsdx(
			bytes,
			evidence.cases.map((item) => ({
				type: 'move-line-endpoint',
				pageId: original.pages[0]!.id,
				shapeId: item.shapeId,
				...item.endpointEdit,
			})),
		);
		const result = await parseVsdx(saved.bytes);
		const native = await parseVsdx(await readFile(join(directory!, 'gradient-raster.vsdx')));
		expect(result.pages[0]!.shapes).toHaveLength(5);
		for (const item of evidence.cases) {
			const actual = result.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
			for (let i = 0; i < 6; i++)
				expect(actual.transform[i]).toBeCloseTo(item.nativeTransform[i]!, 12);
			expect(actual.width).toBeCloseTo(item.nativeShapeWidth, 12);
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
			expect(actual.style.lineGradient?.type).toBe('linear');
		}
		expect(result.pages[0]!.shapes[4]).toEqual(original.pages[0]!.shapes[4]);
		const before = await VisioPackage.open(bytes);
		const after = await VisioPackage.open(saved.bytes);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(after.paths().sort()).toEqual(before.paths().sort());
		for (const path of before.paths())
			if (!saved.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
	});
}
