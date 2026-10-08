import { expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { attribute, children } from './sheet';

interface TextBoxCase {
	pageId: string;
	text: string;
	drawingScale: number;
	native: {
		shapeId: string;
		text: string;
		values: Record<string, { value: number; formula: string }>;
		transform: number[];
	};
}
for (const variable of ['VISIO_NATIVE_TEXT_BOX_DIR', 'VISIO_NATIVE_TEXT_BOX_CUSTOM_DIR']) {
	const directory = process.env[variable];
	it.skipIf(!directory)(
		`matches native fixed text-box styles, geometry and drawing scales (${variable})`,
		async () => {
			const original = await readFile(join(directory!, 'original.vsdx'));
			const native = await readFile(join(directory!, 'native.vsdx'));
			const evidence = JSON.parse(
				(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim().replace(/^\ufeff/, ''),
			) as { cases: TextBoxCase[] };
			const expected = await parseVsdx(native);
			const result = await editVsdx(
				original,
				evidence.cases.map((item) => ({
					type: 'create-text-box',
					pageId: item.pageId,
					shapeId: item.native.shapeId,
					x: item.native.values.PinX!.value,
					y: item.native.values.PinY!.value,
					width: item.native.values.Width!.value,
					height: item.native.values.Height!.value,
					text: item.text,
				})),
			);
			const actual = await parseVsdx(result.bytes);
			expect(actual.pages.length).toBe(evidence.cases.length);
			for (const [index, item] of evidence.cases.entries()) {
				const page = actual.pages[index]!,
					shape = page.shapes[0]!,
					reference = expected.pages[index]!.shapes[0]!;
				expect(page.shapes.length).toBe(1);
				expect(shape.text.plainText).toBe(item.text);
				expect(shape.text).toEqual(reference.text);
				expect(shape.style).toEqual(reference.style);
				expect(shape.geometry).toEqual(reference.geometry);
				expect(shape.style).toMatchObject({ fill: 'none', linePattern: 0 });
				for (let i = 0; i < 6; i++)
					expect(shape.transform[i]).toBeCloseTo(
						item.native.transform[i]! / (i < 4 ? 1 : item.drawingScale),
						12,
					);
			}
			const before = await VisioPackage.open(original),
				after = await VisioPackage.open(result.bytes),
				nativePackage = await VisioPackage.open(native);
			for (const path of before.paths()) {
				if (!result.changedParts.includes(path)) {
					expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
					continue;
				}
				const root = await after.readXml(path),
					nativeRoot = await nativePackage.readXml(path);
				const shape = children(children(root, 'Shapes')[0], 'Shape')[0]!,
					reference = children(children(nativeRoot, 'Shapes')[0], 'Shape')[0]!;
				for (const name of ['LineStyle', 'FillStyle', 'TextStyle'])
					expect(attribute(shape, name)).toBe(attribute(reference, name));
			}
			await writeFile(join(directory!, 'core.vsdx'), result.bytes);
		},
		30_000,
	);
}
