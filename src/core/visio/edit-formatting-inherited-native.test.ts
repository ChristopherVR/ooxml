import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

const directory = process.env.VISIO_NATIVE_INHERITED_FORMATTING_DIR;
describe.skipIf(!directory)('native acceptance of explicit inherited text cell variants', () => {
	it('matches native overrides of cached and cacheless Inh cells without losing mixed font sizes', async () => {
		const bytes = new Uint8Array(await readFile(join(directory!, 'original.vsdx')));
		const model = await parseVsdx(bytes);
		const native = await parseVsdx(new Uint8Array(await readFile(join(directory!, 'native.vsdx'))));
		expect(model.pages[0]!.shapes).toHaveLength(3);
		for (const shape of model.pages[0]!.shapes) {
			expect(shape.text.runs.every((run) => run.italic && !run.bold)).toBe(true);
			expect(shape.text.paragraphs?.[0]?.horizontalAlign).toBe('right');
		}
		const saved = await editVsdx(
			bytes,
			model.pages[0]!.shapes.map((shape) => ({
				type: 'format-text',
				pageId: model.pages[0]!.id,
				shapeId: shape.id,
				bold: true,
				horizontalAlign: 'left',
			})),
		);
		const actual = await parseVsdx(saved.bytes);
		for (const [index, shape] of actual.pages[0]!.shapes.entries()) {
			const reference = native.pages[0]!.shapes[index]!;
			expect(shape.text, shape.id).toEqual(reference.text);
			expect(shape.style).toEqual(reference.style);
			expect(shape.geometry).toEqual(reference.geometry);
			expect(shape.transform).toEqual(reference.transform);
		}
		await writeFile(join(directory!, 'core.vsdx'), saved.bytes);
	});
});
