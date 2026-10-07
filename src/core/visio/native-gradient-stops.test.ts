import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import {
	visioRenderedGradientStops,
	visioRenderedGradientStopCount,
	visioStrokeGradient,
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
	it('projects oblique stroke paint through physical dimensions without altering its source cache', () => {
		const original: VisioLinearGradient = {
			...paint(),
			start: [0, 1],
			end: [1, 1],
			boundingBoxAngle: -45,
		};
		const rendered = visioStrokeGradient(original, 0.2, [4, 2]);
		if (rendered?.type !== 'linear') throw new Error('Expected linear stroke paint.');
		expect(rendered.boundingBoxAngle).toBeUndefined();
		expect(rendered.start[0]).toBeCloseTo(0.4);
		expect(rendered.start[1]).toBeCloseTo(2.6);
		expect(rendered.end[0]).toBeCloseTo(3.6);
		expect(rendered.end[1]).toBeCloseTo(-0.6);
		expect(rendered.stops).toBe(original.stops);
		expect(original.boundingBoxAngle).toBe(-45);
		const scaled = visioStrokeGradient(original, 0.2, [2, 1]);
		if (scaled?.type !== 'linear') throw new Error('Expected scaled stroke paint.');
		expect(scaled.start[0]).toBeCloseTo(0.15);
		expect(scaled.start[1]).toBeCloseTo(1.35);
	});
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

for (const [name, count, profiles] of [
	['VISIO_NATIVE_GRADIENT_RASTER_DIR', 68, 17],
	['VISIO_NATIVE_EXTENDED_GRADIENT_RASTER_DIR', 84, 21],
	['VISIO_NATIVE_ROTATED_GRADIENT_RASTER_DIR', 32, 8],
	['VISIO_NATIVE_OBLIQUE_FILL_RASTER_DIR', 4, 1],
	['VISIO_NATIVE_ROTATED_OBLIQUE_FILL_RASTER_DIR', 4, 1],
] as const) {
	const directory = process.env[name];
	it.skipIf(!directory)(
		`preserves all ${count} genuine native raster cases through editing`,
		async () => {
			const bytes = await readFile(join(directory!, 'gradient-raster.vsdx'));
			const before = await parseVsdx(bytes),
				page = before.pages[0]!;
			expect(page.shapes).toHaveLength(count);
			const saved = await editVsdx(bytes, [
				{ type: 'move-shape', pageId: page.id, shapeId: page.shapes[0]!.id, x: 2.25, y: 1.5 },
			]);
			const after = await parseVsdx(saved.bytes);
			for (let i = 0; i < count; i++)
				expect(after.pages[0]!.shapes[i]!.style.fillGradient).toEqual(
					page.shapes[i]!.style.fillGradient,
				);
			expect(page.shapes.filter((shape) => shape.style.fillGradient?.interpolation)).toHaveLength(
				profiles,
			);
		},
	);
}
