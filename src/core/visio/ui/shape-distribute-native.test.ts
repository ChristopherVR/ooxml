import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from '../edit';
import { parseVsdx } from '../parser';
import { visioArrangeCommands } from './shape-arrange';

const directory = process.env.VISIO_NATIVE_DISTRIBUTION_DIR;
describe.skipIf(!directory)('native Visio overlapping and tied distribution oracle', () => {
	it('matches all native pins and preserves unrelated properties in one source transaction', async () => {
		const original = new Uint8Array(await readFile(join(directory!, 'original.vsdx')));
		const source = await parseVsdx(original);
		const native = await parseVsdx(new Uint8Array(await readFile(join(directory!, 'native.vsdx'))));
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: { pageId: string; axis: 'horizontal' | 'vertical'; selection: string[] }[];
		};
		const edits: VisioEdit[] = [];
		for (const item of evidence.cases) {
			const page = source.pages.find((page) => page.id === item.pageId)!;
			const commands = visioArrangeCommands(page, item.selection, {
				type: 'distribute',
				axis: item.axis,
			});
			expect(commands, page.name).toBeDefined();
			edits.push(...commands!);
		}
		const saved = await editVsdx(original, edits);
		const actual = await parseVsdx(saved.bytes);
		expect(actual.pages.map((page) => page.id)).toEqual(native.pages.map((page) => page.id));
		for (const [index, page] of actual.pages.entries()) {
			const expected = native.pages[index]!;
			expect(
				page.shapes.map((shape) => shape.id),
				page.name,
			).toEqual(expected.shapes.map((shape) => shape.id));
			for (const shape of page.shapes) {
				const reference = expected.shapes.find((candidate) => candidate.id === shape.id)!;
				const context = `${page.name}/${shape.id}`;
				expect(shape.rotation!.pinX, `${context} pinX`).toBeCloseTo(reference.rotation!.pinX, 8);
				expect(shape.rotation!.pinY, `${context} pinY`).toBeCloseTo(reference.rotation!.pinY, 8);
				shape.transform.forEach((value, index) =>
					expect(value, context).toBeCloseTo(reference.transform[index]!, 8),
				);
				expect(shape.geometry, context).toEqual(reference.geometry);
				expect(shape.style, context).toEqual(reference.style);
				const { transform: textTransform, ...text } = shape.text;
				const { transform: expectedTransform, ...expectedText } = reference.text;
				expect(text, context).toEqual(expectedText);
				textTransform.forEach((value, index) =>
					expect(value, context).toBeCloseTo(expectedTransform[index]!, 8),
				);
			}
		}
		await writeFile(join(directory!, 'core.vsdx'), saved.bytes);
	});
});
