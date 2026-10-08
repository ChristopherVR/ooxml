import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

const directory = process.env.VISIO_NATIVE_NEW_DRAWING_DIR;
/** Factory files and native reopen/save output from the two scripts in scripts/. */
describe.skipIf(!directory)('native new drawing acceptance', () => {
	it('matches page, geometry and text after native creation and resaving at three paper sizes', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as {
			application: string;
			cases: {
				name: string;
				width: number;
				height: number;
				initialShapes: number;
				shape: { font: string; text: string; 'Char.Size': number };
			}[];
		};
		expect(evidence.application).toBe('Microsoft Visio');
		expect(evidence.cases).toHaveLength(6);
		for (const item of evidence.cases) {
			const blank = await createVsdx({ width: item.width, height: item.height });
			const drawn = await editVsdx(blank, [
				{
					type: 'create-rectangle',
					pageId: '0',
					shapeId: '1',
					x: 2,
					y: 3,
					width: 2,
					height: 1,
					text: 'New diagram',
				},
			]);
			const core = (await parseVsdx(drawn.bytes)).pages[0]!;
			const native = (await parseVsdx(await readFile(join(directory!, `${item.name}-native.vsdx`))))
				.pages[0]!;
			expect(native.width).toBeCloseTo(core.width, 10);
			expect(native.height).toBeCloseTo(core.height, 10);
			expect(native.shapes).toHaveLength(1);
			expect(native.shapes[0]!.transform).toEqual(core.shapes[0]!.transform);
			// COM Text assignment appends a paragraph terminator; reopening existing core text does not.
			expect(native.shapes[0]!.text.plainText).toBe(
				core.shapes[0]!.text.plainText + (item.initialShapes === 0 ? '\n' : ''),
			);
			expect(item.shape.font).toBe('Calibri');
			expect(item.shape['Char.Size']).toBeCloseTo(12 / 72, 10);
			expect(item.initialShapes).toBe(item.name.endsWith('-blank') ? 0 : 1);
		}
	});
});
