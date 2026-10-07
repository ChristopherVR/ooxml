import { expect, it } from 'vitest';
import reference from './__fixtures__/fill-patterns-native.json';
import { fillPatternPixels } from './fill-pattern';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { cell, fixture, rectangle, shape } from './test-fixtures';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

it('matches every pixel of all 23 native opaque and translucent pattern tiles', () => {
	for (const item of reference.cases) {
		expect(
			Buffer.from(fillPatternPixels(item.pattern, '#ff0000', '#0000ff', 1, 1)!).toString('hex'),
			`opaque ${item.pattern}`,
		).toBe(item.pixels);
		expect(
			Buffer.from(fillPatternPixels(item.pattern, '#1b8bd3', '#e73d53', 0.8, 0.5)!).toString('hex'),
			`alpha ${item.pattern}`,
		).toBe(item.alphaPixels);
	}
});

const nativeDirectory = process.env.VISIO_NATIVE_FILL_PATTERNS_DIR;

it('reports the remaining oblique raster fidelity gap while accepting quarter-turn hatch axes', async () => {
	const document = await parseVsdx(
		await fixture({
			pages: [Math.PI / 2, Math.PI / 6].map((angle, index) => ({
				id: String(index),
				contents: `<Shapes>${shape('1', rectangle + cell('FillPattern', 2) + cell('FillForegnd', '#ff0000') + cell('FillBkgnd', '#0000ff') + cell('Angle', angle))}</Shapes>`,
			})),
		}),
	);
	const warnings = document.diagnostics.filter((item) => item.code === 'unverified-hatch-angle');
	expect(warnings).toHaveLength(1);
	expect(warnings[0]!.pageId).toBe('1');
});
it.skipIf(!nativeDirectory)(
	'moves and saves the genuine 23-page native drawing without changing its pattern resources',
	async () => {
		const bytes = new Uint8Array(await readFile(join(nativeDirectory!, 'fill-patterns.vsdx')));
		const original = await parseVsdx(bytes);
		const saved = await editVsdx(bytes, [
			{
				type: 'move-shape',
				pageId: original.pages[0]!.id,
				shapeId: original.pages[0]!.shapes[0]!.id,
				x: 2.25,
				y: 1.5,
			},
		]);
		const reopened = await parseVsdx(saved.bytes);
		expect(reopened.pages).toHaveLength(23);
		for (let index = 0; index < 23; index++)
			expect(reopened.pages[index]!.shapes[0]!.style.fillPattern!.bytes).toEqual(
				original.pages[index]!.shapes[0]!.style.fillPattern!.bytes,
			);
		await writeFile(join(nativeDirectory!, 'core-fill-patterns.vsdx'), saved.bytes);
	},
);

it('normalizes all native pattern codes and retains tiles through a move/save/reparse', async () => {
	const source = await fixture({
		pages: reference.cases.map((item) => ({
			id: String(item.pattern),
			contents: `<Shapes>${shape('1', rectangle + cell('Width', 2) + cell('Height', 1) + cell('PinX', 1) + cell('PinY', 1) + cell('FillPattern', item.pattern) + cell('FillForegnd', '#ff0000') + cell('FillBkgnd', '#0000ff'))}</Shapes>`,
		})),
	});
	const original = await parseVsdx(source);
	const saved = await editVsdx(source, [
		{ type: 'move-shape', pageId: '2', shapeId: '1', x: 2, y: 2 },
	]);
	const reopened = await parseVsdx(saved.bytes);
	expect(original.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(false);
	for (let index = 0; index < reference.cases.length; index++) {
		const before = original.pages[index]!.shapes[0]!.style;
		expect(before.fillPattern).toMatchObject({
			width: 1 / 12,
			height: 1 / 12,
			mimeType: 'image/png',
			pixelWidth: 8,
			pixelHeight: 8,
		});
		expect(reopened.pages[index]!.shapes[0]!.style.fillPattern!.bytes).toEqual(
			before.fillPattern!.bytes,
		);
		expect(before.fillOpacity).toBe(1);
	}
});
