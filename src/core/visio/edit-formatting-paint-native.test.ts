import { expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioShapeFormatEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';

const directory = process.env['VISIO_NATIVE_PAINT_FORMATTING_DIR'];
it.skipIf(!directory)(
	'matches native paint formatting across three drawing scales',
	async () => {
		const original = await readFile(join(directory!, 'original.vsdx'));
		const native = await readFile(join(directory!, 'native.vsdx'));
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim().replace(/^\ufeff/, ''),
		) as { cases: { command: VisioShapeFormatEdit }[] };
		const result = await editVsdx(
			original,
			evidence.cases.map((item) => item.command),
		);
		const actual = await parseVsdx(result.bytes),
			expected = await parseVsdx(native);
		expect(actual.pages.length).toBe(evidence.cases.length);
		for (let index = 0; index < actual.pages.length; index++) {
			const shape = actual.pages[index]!.shapes[0]!,
				reference = expected.pages[index]!.shapes[0]!;
			expect({ ...shape.style, lineWidth: 0 }, `page ${index}`).toEqual({
				...reference.style,
				lineWidth: 0,
			});
			expect(shape.style.lineWidth).toBeCloseTo(reference.style.lineWidth, 14);
			expect(shape.text).toEqual(reference.text);
			expect(shape.geometry).toEqual(reference.geometry);
			expect(shape.transform).toEqual(reference.transform);
		}
		const before = await VisioPackage.open(original),
			after = await VisioPackage.open(result.bytes);
		for (const path of before.paths())
			if (!result.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		await writeFile(join(directory!, 'core.vsdx'), result.bytes);
	},
	30_000,
);
