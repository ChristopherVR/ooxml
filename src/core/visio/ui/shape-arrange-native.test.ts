import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from '../edit';
import { parseVsdx } from '../parser';
import { visioArrangeCommands, type VisioArrangement } from './shape-arrange';

const directory = process.env.VISIO_NATIVE_ARRANGEMENT_DIR;
describe.skipIf(!directory)('native Visio align/distribute oracle', () => {
	it('matches native pins, transforms and dimensions for every axis at three page scales', async () => {
		const original = new Uint8Array(await readFile(join(directory!, 'original.vsdx')));
		const source = await parseVsdx(original);
		const native = await parseVsdx(new Uint8Array(await readFile(join(directory!, 'native.vsdx'))));
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: { pageId: string; action: string; selection: string[] }[];
		};
		const edits: VisioEdit[] = [];
		for (const item of evidence.cases) {
			const page = source.pages.find((page) => page.id === item.pageId)!;
			const action: VisioArrangement =
				item.action === 'horizontal' || item.action === 'vertical'
					? { type: 'distribute', axis: item.action }
					: {
							type: 'align',
							edge: item.action as Extract<VisioArrangement, { type: 'align' }>['edge'],
						};
			const commands = visioArrangeCommands(page, item.selection, action);
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
			// Visio computes alignment extents with float32 local dimensions. Do not round
			// model coordinates: bound the source/anchor boxes' quantization independently.
			const scale = source.pages[index]!.drawingToPageScale ?? 1;
			const rounding = Math.max(
				...source.pages[index]!.shapes.map((shape) => {
					const widthError =
						Math.abs(Math.fround(shape.width / scale) - shape.width / scale) * scale;
					const heightError =
						Math.abs(Math.fround(shape.height / scale) - shape.height / scale) * scale;
					const [a, b, c, d] = shape.transform;
					return Math.max(
						Math.abs(a) * widthError + Math.abs(c) * heightError,
						Math.abs(b) * widthError + Math.abs(d) * heightError,
					);
				}),
			);
			const tolerance = 2 * rounding + 1e-11;
			for (const shape of page.shapes) {
				const reference = expected.shapes.find((candidate) => candidate.id === shape.id)!;
				expect(
					Math.abs(shape.rotation!.pinX - reference.rotation!.pinX),
					`${page.name}/${shape.id} pinX`,
				).toBeLessThanOrEqual(tolerance);
				expect(
					Math.abs(shape.rotation!.pinY - reference.rotation!.pinY),
					`${page.name}/${shape.id} pinY`,
				).toBeLessThanOrEqual(tolerance);
				shape.transform.forEach((coefficient, index) =>
					expect(Math.abs(coefficient - reference.transform[index]!)).toBeLessThanOrEqual(
						tolerance,
					),
				);
				expect(shape.geometry).toEqual(reference.geometry);
				expect(shape.text).toEqual(reference.text);
				expect(shape.style).toEqual(reference.style);
			}
		}
		await writeFile(join(directory!, 'core.vsdx'), saved.bytes);
	});
});
