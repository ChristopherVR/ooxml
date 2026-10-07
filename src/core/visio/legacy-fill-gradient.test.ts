import { expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { assertViewableDocument } from './ui/scene-validation';
import { copySnapshotScene } from './ui/snapshot-scene';
import { cell, fixture, rectangle, shape } from './test-fixtures';

const nativeRadial = process.env.VISIO_NATIVE_RADIAL_FILLS_DIR;
const nativeRadialAlpha = process.env.VISIO_NATIVE_RADIAL_FILLS_ALPHA_DIR;
it('preserves all classic region triangles, stops and snapshot allocations through scaling and edits', async () => {
	const angles = [
		[90, 360],
		[180, 90],
		[270, 360],
		[270, 180],
		[180, 360, 270, 90],
	];
	const source = await fixture({
		pages: angles.map((_, index) => ({
			id: String(index),
			pageCells: cell('DrawingScale', 2) + cell('PageScale', 1),
			contents: `<Shapes>${shape('1', rectangle + cell('Width', 2) + cell('Height', 1) + cell('FillPattern', index + 31) + cell('FillForegnd', '#ff0000') + cell('FillBkgnd', '#0000ff') + cell('FillForegndTrans', 0.25) + cell('FillBkgndTrans', 0.6))}</Shapes>`,
		})),
	});
	const original = await parseVsdx(source);
	const saved = await editVsdx(source, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 2 },
	]);
	const reopened = await parseVsdx(saved.bytes);
	const copy = copySnapshotScene(original);
	for (const [index, expectedAngles] of angles.entries()) {
		const paint = original.pages[index]!.shapes[0]!.style.fillGradient!;
		if (paint.type !== 'regions') throw new Error('Expected region gradient.');
		expect(paint.regions.map((region) => region.angle)).toEqual(expectedAngles);
		expect(paint.stops.map((stop) => stop.opacity)).toEqual([0.75, 0.4]);
		expect(reopened.pages[index]!.shapes[0]!.style.fillGradient).toEqual(paint);
		expect(copy.pages[index]!.shapes[0]!.style.fillGradient).toEqual(paint);
	}
	assertViewableDocument(original);
	assertViewableDocument(copy);
	const paint = original.pages[0]!.shapes[0]!.style.fillGradient!;
	if (paint.type !== 'regions') throw new Error('Expected region gradient.');
	paint.regions[0]!.angle = NaN;
	expect(() => assertViewableDocument(original)).toThrow('gradient region angle');
	assertViewableDocument(copy);
});

const nativeRegions = process.env.VISIO_NATIVE_REGION_FILLS_DIR;
const nativeRegionsAlpha = process.env.VISIO_NATIVE_REGION_FILLS_ALPHA_DIR;
it.skipIf(!nativeRegions || !nativeRegionsAlpha)(
	'preserves genuine native region paint through move/save/reparse',
	async () => {
		for (const directory of [nativeRegions!, nativeRegionsAlpha!]) {
			const bytes = new Uint8Array(await readFile(join(directory, 'fill-patterns.vsdx')));
			const original = await parseVsdx(bytes);
			expect(original.pages).toHaveLength(5);
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
			for (let index = 0; index < 5; index++) {
				expect(original.pages[index]!.shapes[0]!.style.fillGradient?.type).toBe('regions');
				expect(reopened.pages[index]!.shapes[0]!.style.fillGradient).toEqual(
					original.pages[index]!.shapes[0]!.style.fillGradient,
				);
			}
			await writeFile(join(directory, 'core-fill-patterns.vsdx'), saved.bytes);
		}
	},
);
it.skipIf(!nativeRadial || !nativeRadialAlpha)(
	'preserves genuine native radial paint through move/save/reparse',
	async () => {
		for (const directory of [nativeRadial!, nativeRadialAlpha!]) {
			const bytes = new Uint8Array(await readFile(join(directory, 'fill-patterns.vsdx')));
			const original = await parseVsdx(bytes);
			expect(original.pages).toHaveLength(5);
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
			for (let index = 0; index < 5; index++) {
				expect(original.pages[index]!.shapes[0]!.style.fillGradient?.type).toBe('radial');
				expect(reopened.pages[index]!.shapes[0]!.style.fillGradient).toEqual(
					original.pages[index]!.shapes[0]!.style.fillGradient,
				);
			}
			await writeFile(join(directory, 'core-fill-patterns.vsdx'), saved.bytes);
		}
	},
);

it('retains native radial centers, independent alpha and normalized radii through save and snapshots', async () => {
	const centers = [
		[0, 1],
		[1, 1],
		[0, 0],
		[1, 0],
		[0.5, 0.5],
	];
	const bytes = await fixture({
		pages: centers.map((_, index) => ({
			id: String(index),
			pageCells: cell('DrawingScale', 2) + cell('PageScale', 1),
			contents: `<Shapes>${shape('1', rectangle + cell('Width', 2) + cell('Height', 1) + cell('FillPattern', index + 36) + cell('FillForegnd', '#00ff00') + cell('FillBkgnd', '#0000ff') + cell('FillForegndTrans', 0.25) + cell('FillBkgndTrans', 0.6))}</Shapes>`,
		})),
	});
	const document = await parseVsdx(bytes);
	const saved = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 2 },
	]);
	const reopened = await parseVsdx(saved.bytes);
	const snapshot = copySnapshotScene(document);
	for (const [index, center] of centers.entries()) {
		const expected = {
			type: 'radial',
			center,
			radius: index === 4 ? 0.73 : 1.4,
			stops: [
				{ offset: 0, color: '#00ff00', opacity: 0.75 },
				{ offset: 1, color: '#0000ff', opacity: 0.4 },
			],
		};
		expect(document.pages[index]!.shapes[0]!.style.fillGradient).toEqual(expected);
		expect(reopened.pages[index]!.shapes[0]!.style.fillGradient).toEqual(expected);
		expect(snapshot.pages[index]!.shapes[0]!.style.fillGradient).toEqual(expected);
	}
	assertViewableDocument(document);
	assertViewableDocument(snapshot);
	const radial = document.pages[0]!.shapes[0]!.style.fillGradient!;
	if (radial.type !== 'radial') throw new Error('Expected radial fill.');
	radial.radius = NaN;
	expect(() => assertViewableDocument(document)).toThrow('gradient radius');
	assertViewableDocument(snapshot);
});

it('preserves independent foreground and background alpha without a colored layer', async () => {
	const document = await parseVsdx(
		await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + cell('FillPattern', 26) + cell('FillForegnd', '#00ff00') + cell('FillBkgnd', '#0000ff') + cell('FillForegndTrans', 0.25) + cell('FillBkgndTrans', 0.6))}</Shapes>`,
				},
			],
		}),
	);
	const style = document.pages[0]!.shapes[0]!.style;
	expect(style.fillOpacity).toBe(1);
	expect(style.fillGradient?.stops).toEqual([
		{ offset: 0, color: '#0000ff', opacity: 0.4 },
		{ offset: 0.5, color: '#00ff00', opacity: 0.75 },
		{ offset: 1, color: '#0000ff', opacity: 0.4 },
	]);
	expect(document.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(false);
});

it('keeps unsupported pattern and incomplete modern gradient diagnostics', async () => {
	for (const settings of [
		cell('FillPattern', 41),
		cell('FillPattern', 25) + cell('FillGradientEnabled', 1),
	]) {
		const document = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', rectangle + settings + cell('FillForegnd', '#00ff00') + cell('FillBkgnd', '#0000ff'))}</Shapes>`,
					},
				],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.style.fillGradient).toBeUndefined();
		expect(document.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(
			true,
		);
	}
});
