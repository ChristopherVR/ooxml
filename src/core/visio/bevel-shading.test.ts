import { describe, expect, it } from 'vitest';
import { VISIO_ROTATION_PRESETS, visioRotation3dMatrix } from './bevel-3d';
import {
	BEVEL_LIGHTING_TYPES,
	bevelHeightAt,
	bevelHeightTable,
	bevelLight,
	bevelMaterial,
	bevelProfile,
	isUndrawnBevelMaterial,
	normalBevelLighting,
} from './bevel-shading';
import { parseVsdx } from './parser';
import { createProjection, projectionMatrix } from './solid-3d';
import { cell, fixture, rectangle, shape } from './test-fixtures';

const read = async (extra: string) => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 2) + cell('Height', 1) + rectangle + extra)}</Shapes>`,
			},
		],
	});
	return (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
};

describe('3-D Rotation gallery', () => {
	it('has unique ids and labels in the Office gallery order', () => {
		const ids = VISIO_ROTATION_PRESETS.map((preset) => preset.id);
		expect(new Set(ids).size).toBe(ids.length);
		expect(ids.slice(0, 5)).toEqual([
			'off',
			'isometric-left-down',
			'isometric-right-up',
			'isometric-top-up',
			'isometric-bottom-down',
		]);
		expect(
			VISIO_ROTATION_PRESETS.find((preset) => preset.id === 'perspective-heroic-left'),
		).toEqual(expect.objectContaining({ label: 'Perspective Heroic Left' }));
	});
	it('keeps the Office camera values with the signs Visio uses', () => {
		const preset = (id: string) => VISIO_ROTATION_PRESETS.find((item) => item.id === id)!;
		// PowerPoint: isometricTopUp reads x 45.28, y -35.4, z -60.16; X and Z are negated for Visio.
		expect(preset('isometric-top-up')).toMatchObject({ x: -45.28, y: -35.4, z: 60.16 });
		// perspectiveLeft reads x -20 with a 45 degree field of view.
		expect(preset('perspective-left')).toMatchObject({ x: 20, y: 0, z: 0, perspective: 45 });
		expect(preset('perspective-heroic-extreme-right').perspective).toBe(80);
	});
	it('gives every parallel preset no perspective and every perspective preset one', () => {
		for (const preset of VISIO_ROTATION_PRESETS) {
			if (preset.id.startsWith('perspective-')) expect(preset.perspective).toBeGreaterThan(0);
			else expect(preset.perspective).toBe(0);
		}
	});
});

describe('bevel profiles', () => {
	it('defines every gallery type and falls back to a chamfer', () => {
		for (let type = 1; type <= 12; type++) {
			const points = bevelProfile(type);
			expect(points[0]!.t).toBe(0);
			expect(points[points.length - 1]!.t).toBe(1);
			for (let i = 1; i < points.length; i++)
				expect(points[i]!.t).toBeGreaterThanOrEqual(points[i - 1]!.t);
		}
		expect(bevelProfile(99)).toEqual(bevelProfile(5));
	});
	it('interpolates and clamps heights', () => {
		const angle = bevelProfile(5);
		expect(bevelHeightAt(angle, 0.25)).toBeCloseTo(0.25);
		expect(bevelHeightAt(angle, -1)).toBe(0);
		expect(bevelHeightAt(angle, 2)).toBe(1);
	});
	it('turns a blurred alpha back into the profile height', () => {
		const table = bevelHeightTable(5);
		expect(table.values).toHaveLength(128);
		expect(Math.min(...table.values)).toBe(0);
		expect(Math.max(...table.values)).toBe(1);
		// Below one half is outside the edge, and the face is reached inside the bevel width.
		expect(table.values[0]).toBe(0);
		expect(table.values[127]).toBe(1);
		for (let i = 1; i < table.values.length; i++)
			expect(table.values[i]!).toBeGreaterThanOrEqual(table.values[i - 1]!);
		// The Angle chamfer rises its whole height; an inset (Soft Round) falls instead.
		expect(table.rise).toBeCloseTo(1);
		expect(bevelHeightTable(6).values[127]).toBeLessThan(bevelHeightTable(6).values[64]! + 1e-9);
	});
	it('keeps each type distinct', () => {
		const tables = Array.from({ length: 12 }, (_, i) => bevelHeightTable(i + 1).values.join());
		expect(new Set(tables).size).toBe(12);
	});
});

describe('bevel lighting and materials', () => {
	it('reads 0 and out-of-range lighting as three point', () => {
		expect(normalBevelLighting(0)).toBe(1);
		expect(normalBevelLighting(16)).toBe(1);
		expect(normalBevelLighting(7.5)).toBe(1);
		expect(normalBevelLighting(15)).toBe(15);
		expect(BEVEL_LIGHTING_TYPES).toBe(15);
	});
	it('fits a light from the top for the default lighting', () => {
		const light = bevelLight(0);
		// Page frame, y down: the light is above (an azimuth between -180 and 0) and well off the plane.
		expect(light.azimuth).toBeLessThan(0);
		expect(light.azimuth).toBeGreaterThan(-180);
		expect(light.elevation).toBeGreaterThan(30);
		expect(light.face).toBeCloseTo(1.04);
		// The fit reproduces the measured top and bottom facets (brightness 1.09 and 0.51).
		const facet = (towardsTop: number) => {
			const az = (light.azimuth * Math.PI) / 180;
			const el = (light.elevation * Math.PI) / 180;
			const ly = Math.cos(el) * Math.sin(az);
			const lz = Math.sin(el);
			return light.ambient + light.diffuse * (lz * Math.SQRT1_2 - towardsTop * ly * Math.SQRT1_2);
		};
		expect(facet(1)).toBeCloseTo(1.09, 1);
		expect(facet(-1)).toBeCloseTo(0.51, 1);
	});
	it('lights from the left for lighting type 5', () => {
		const light = bevelLight(5);
		expect(Math.abs(light.azimuth)).toBeGreaterThan(135);
	});
	it('maps materials', () => {
		expect(bevelMaterial(0)).toEqual({ contrast: 1, lift: 0 });
		expect(bevelMaterial(7).contrast).toBe(0);
		expect(bevelMaterial(10).lift).toBeGreaterThan(bevelMaterial(9).lift);
		expect(isUndrawnBevelMaterial(8)).toBe(true);
		expect(isUndrawnBevelMaterial(0)).toBe(false);
	});
	it('reads the lighting and material cells with the bevel', async () => {
		const plain = await read(
			cell('BevelTopType', 5) + cell('BevelTopWidth', 0.2) + cell('BevelTopHeight', 0.2),
		);
		expect(plain.style.bevel).toEqual({ type: 5, width: 0.2, height: 0.2 });
		const lit = await read(
			cell('BevelTopType', 5) +
				cell('BevelTopWidth', 0.2) +
				cell('BevelTopHeight', 0.2) +
				cell('BevelLightingType', 4) +
				cell('BevelMaterialType', 9),
		);
		expect(lit.style.bevel).toMatchObject({ lighting: 4, material: 9 });
		const wild = await read(
			cell('BevelTopType', 5) + cell('BevelLightingType', 40) + cell('BevelMaterialType', 40),
		);
		expect(wild.style.bevel).not.toHaveProperty('lighting');
		expect(wild.style.bevel).not.toHaveProperty('material');
	});
});

describe('projectionMatrix', () => {
	it('is the parallel matrix about the centre when there is no perspective', () => {
		const turn = { x: 0.4, y: -0.3, z: 0.2, perspective: 0 };
		const projection = createProjection(turn, { x: 1, y: 0.5 }, 1);
		const [a, b, c, d, e, f] = projectionMatrix(projection, 2, 1);
		const [pa, pb, pc, pd] = visioRotation3dMatrix(turn);
		for (const [got, want] of [
			[a, pa],
			[b, pb],
			[c, pc],
			[d, pd],
		] as const)
			expect(got).toBeCloseTo(want, 9);
		// The centre stays where it was.
		expect(a * 1 + c * 0.5 + e).toBeCloseTo(1, 9);
		expect(b * 1 + d * 0.5 + f).toBeCloseTo(0.5, 9);
	});
	it('keeps the centre fixed and foreshortens the far side under perspective', () => {
		const turn = { x: 0.6, y: 0, z: 0, perspective: Math.PI / 3 };
		const projection = createProjection(turn, { x: 1, y: 0.5 }, 1);
		const [a, b, c, d, e, f] = projectionMatrix(projection, 2, 1);
		expect(a * 1 + c * 0.5 + e).toBeCloseTo(1, 9);
		expect(b * 1 + d * 0.5 + f).toBeCloseTo(0.5, 9);
		// X turns the shape about the vertical axis, so the width shrinks and the height stays near.
		expect(a).toBeLessThan(1);
		expect(a).toBeGreaterThan(0.5);
		expect(d).toBeGreaterThan(0.9);
		expect(Math.abs(b)).toBeLessThan(0.05);
	});
});
