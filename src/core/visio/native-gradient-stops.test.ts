import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import {
	visioRenderedGradientStops,
	visioRenderedGradientStopCount,
} from './native-gradient-stops';
import { demoDocument } from './ui/demo-document';
import { assertViewableDocument } from './ui/scene-validation';
import { copySnapshotScene } from './ui/snapshot-scene';
import { normalizeVisioPageGeometry } from './page-scale';
import type { VisioLinearGradient } from './model';

const paint = (): VisioLinearGradient => ({
	type: 'linear',
	start: [0, 0],
	end: [2, 0],
	interpolation: 'sigma-gamma22',
	stops: [
		{ offset: 0, color: '#ff0000', opacity: 1 },
		{ offset: 1, color: '#0000ff', opacity: 1 },
	],
});

describe('native opaque gradient sampling', () => {
	it('retains source stops and uses the observed brighter quarter colors', () => {
		const original = paint(),
			stops = visioRenderedGradientStops(original);
		expect(stops).toHaveLength(256);
		expect(stops[0]).toEqual({ offset: 0, color: '#FF0000', opacity: 1 });
		expect(stops[32]!.color).toBe('#FA003F');
		expect(stops[64]!.color).toBe('#EE0069');
		expect(stops[255]).toEqual({ offset: 1, color: '#0000FF', opacity: 1 });
		expect(original.stops).toHaveLength(2);
		delete original.interpolation;
		expect(visioRenderedGradientStops(original)).toBe(original.stops);
	});
	it('preserves interpolation and local radial coordinates across snapshots and scaling', () => {
		const model = structuredClone(demoDocument),
			shape = model.pages[0]!.shapes[0]!;
		shape.style.fillGradient = {
			type: 'radial',
			coordinateSpace: 'local',
			center: [2, 1],
			radius: Math.hypot(2, 1),
			interpolation: 'sigma-gamma22',
			stops: paint().stops,
		};
		const copy = copySnapshotScene(model);
		expect(copy.pages[0]!.shapes[0]!.style.fillGradient).toEqual(shape.style.fillGradient);
		normalizeVisioPageGeometry(copy.pages[0]!.shapes, 0.5, () => {});
		expect(copy.pages[0]!.shapes[0]!.style.fillGradient).toMatchObject({
			center: [1, 0.5],
			radius: Math.hypot(1, 0.5),
			coordinateSpace: 'local',
			interpolation: 'sigma-gamma22',
		});
		expect(shape.style.fillGradient.center).toEqual([2, 1]);
	});
	it('counts generated stops against scene limits and rejects unsupported native profiles', () => {
		const model = structuredClone(demoDocument),
			shape = model.pages[0]!.shapes[0]!;
		shape.style.fillGradient = paint();
		model.pages = [model.pages[0]!];
		expect(visioRenderedGradientStopCount(shape.style.fillGradient)).toBe(256);
		model.pages[0]!.shapes = Array.from({ length: 390 }, (_, id) => ({ ...shape, id: String(id) }));
		assertViewableDocument(model);
		model.pages[0]!.shapes.push({ ...shape, id: 'over-budget' });
		expect(() => assertViewableDocument(model)).toThrow('gradient stop limits');
		model.pages[0]!.shapes = [shape];
		shape.style.fillGradient.stops[0]!.opacity = 0.5;
		expect(() => assertViewableDocument(model)).toThrow('native gradient interpolation');
		expect(() => visioRenderedGradientStops(shape.style.fillGradient!)).toThrow(
			'native gradient interpolation',
		);
	});
});

const directory = process.env.VISIO_NATIVE_GRADIENT_RASTER_DIR;
it.skipIf(!directory)('preserves all 68 genuine native raster cases through editing', async () => {
	const bytes = await readFile(join(directory!, 'gradient-raster.vsdx'));
	const before = await parseVsdx(bytes),
		page = before.pages[0]!;
	expect(page.shapes).toHaveLength(68);
	const saved = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: page.id, shapeId: page.shapes[0]!.id, x: 2.25, y: 1.5 },
	]);
	const after = await parseVsdx(saved.bytes);
	for (let i = 0; i < 68; i++)
		expect(after.pages[0]!.shapes[i]!.style.fillGradient).toEqual(
			page.shapes[i]!.style.fillGradient,
		);
	expect(page.shapes.filter((shape) => shape.style.fillGradient?.interpolation)).toHaveLength(17);
});
