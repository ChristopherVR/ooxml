import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { generatedTheme, themeFixture } from './theme-fixtures';
import { cell, fixture, shape, rectangle } from './test-fixtures';
import { snapshotFormatting } from './edit-formatting-commands';
import { visioRotation3dMatrix } from './bevel-3d';
import { createProjection, extrudePaths, flattenPath, projectPathData } from './solid-3d';

const deg = (value: number) => (value * Math.PI) / 180;
const centre = { x: 1.5, y: 1 };
const turn = (x: number, y: number, z: number, perspective: number) => ({ x, y, z, perspective });

describe('Visio 3-D projection', () => {
	it('matches the parallel matrix when there is no perspective', () => {
		const rotation = turn(deg(30), deg(-20), deg(15), 0);
		const [a, b, c, d] = visioRotation3dMatrix(rotation);
		const projection = createProjection(rotation, centre);
		for (const [x, y] of [
			[0, 0],
			[3, 2],
			[0.4, 1.7],
		] as const) {
			const [dx, dy] = [x - centre.x, y - centre.y];
			const point = projection.point(x, y);
			expect(point.x).toBeCloseTo(centre.x + a * dx + c * dy, 9);
			expect(point.y).toBeCloseTo(centre.y + b * dx + d * dy, 9);
		}
	});
	it('foreshortens by the measured camera distance (right edge 1.17 times the left at 60 degrees)', () => {
		// Visio 16, 3 x 2 in shape, RotationXAngle 35 deg, Perspective 60 deg: edge heights 270 and
		// 316 px at 144 px per inch.
		const projection = createProjection(turn(deg(35), 0, 0, deg(60)), centre);
		const left = Math.abs(projection.point(0, 0).y - projection.point(0, 2).y);
		const right = Math.abs(projection.point(3, 0).y - projection.point(3, 2).y);
		expect(right / left).toBeGreaterThan(1.15);
		expect(right / left).toBeLessThan(1.19);
		expect(projection.point(1.5, 1).x).toBeCloseTo(1.5, 9);
	});
	it('clamps the field of view at 120 degrees and scales with the unit', () => {
		const wide = createProjection(turn(deg(35), 0, 0, deg(170)), centre).point(3, 0).y;
		const limit = createProjection(turn(deg(35), 0, 0, deg(120)), centre).point(3, 0).y;
		expect(wide).toBeCloseTo(limit, 9);
		// A local unit of half an inch puts the camera twice as far in local units.
		const near = createProjection(turn(deg(35), 0, 0, deg(60)), centre, 1).point(3, 0).y;
		const far = createProjection(turn(deg(35), 0, 0, deg(60)), centre, 0.5).point(3, 0).y;
		expect(Math.abs(far - 1)).toBeLessThan(Math.abs(near - 1));
	});
	it('flattens curves and projects path data', () => {
		const flat = flattenPath('M 0 0 C 1 0 2 1 2 2 L 0 2 Z')!;
		expect(flat[0]!.closed).toBe(true);
		expect(flat[0]!.points).toHaveLength(14);
		expect(flattenPath('M 0 0 L')).toBeUndefined();
		const projected = projectPathData(
			'M 0 0 L 3 0 L 3 2 L 0 2 Z',
			createProjection(turn(deg(35), 0, 0, deg(60)), centre),
		)!;
		expect(projected.startsWith('M ')).toBe(true);
		expect(projected.endsWith(' Z')).toBe(true);
	});
});

describe('Visio extrusion', () => {
	const box = 'M 0 0 L 3 0 L 3 2 L 0 2 Z';
	const sidesOf = (x: number, y: number) =>
		extrudePaths([box], createProjection(turn(deg(x), deg(y), 0, 0), centre), 0.4)!;
	const centreOf = (points: readonly { x: number; y: number }[]) => ({
		x: points.reduce((sum, p) => sum + p.x, 0) / points.length,
		y: points.reduce((sum, p) => sum + p.y, 0) / points.length,
	});
	it('shows the right side when RotationXAngle is positive and the left when negative', () => {
		const right = sidesOf(35, 0).sides;
		expect(right).toHaveLength(1);
		expect(centreOf(right[0]!.points).x).toBeGreaterThan(2.5);
		const left = sidesOf(-35, 0).sides;
		expect(left).toHaveLength(1);
		expect(centreOf(left[0]!.points).x).toBeLessThan(0.5);
	});
	it('shows the top for a positive RotationYAngle and the bottom for a negative one', () => {
		const top = sidesOf(0, 35).sides;
		expect(top).toHaveLength(1);
		expect(centreOf(top[0]!.points).y).toBeGreaterThan(1.5);
		const bottom = sidesOf(0, -35).sides;
		expect(bottom).toHaveLength(1);
		expect(centreOf(bottom[0]!.points).y).toBeLessThan(0.5);
	});
	it('offsets the back outline by depth * sin(angle) like Visio 16 (30 pt, 35 deg is 34 px at 144 ppi)', () => {
		const projection = createProjection(turn(0, deg(35), 0, 0), centre);
		const { back } = extrudePaths([box], projection, 0.4)!;
		// 0.4 in * sin 35 deg
		expect(back[0]![0]!.y - projection.point(0, 0).y).toBeCloseTo(0.4 * Math.sin(deg(35)), 6);
	});
	it('draws nothing for a face-on shape and shades the right side darker than the bottom', () => {
		expect(sidesOf(0, 0).sides).toHaveLength(0);
		const both = extrudePaths([box], createProjection(turn(deg(35), deg(-35), 0, 0), centre), 0.4)!;
		expect(both.sides).toHaveLength(2);
		const shades = both.sides.map((face) => [centreOf(face.points).x, face.shade] as const);
		const [right, bottom] = [...shades].sort((a, b) => b[0] - a[0]);
		expect(right![1]).toBeLessThan(bottom![1]);
	});
	it('leaves open paths and malformed data out', () => {
		const projection = createProjection(turn(deg(35), 0, 0, 0), centre);
		expect(extrudePaths(['M 0 0 L 3 0'], projection, 0.4)).toEqual({ back: [], sides: [] });
		expect(extrudePaths(['M 0 0 L'], projection, 0.4)).toBeUndefined();
	});
});

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const plain = (extra = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 3) + cell('Height', 2) + rectangle + extra)}</Shapes>`,
			},
		],
	});
const readStyle = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!.style;

describe('Visio extrusion and contour cells', () => {
	it('writes BevelDepth and BevelContour cells in points and reads them back with a bevel', async () => {
		const saved = await editVsdx(await plain(), [
			{
				...target,
				bevel: 'circle',
				rotation3d: 'isometric-right-up',
				extrusionDepth: 20,
				extrusionColor: '#ff0000',
				contourSize: 6,
				contourColor: '#008000',
			},
		]);
		const style = await readStyle(saved.bytes);
		expect(style.extrusion).toEqual({ depth: 20 / 72, color: '#ff0000' });
		expect(style.contour).toEqual({ size: 6 / 72, color: '#008000' });
	});
	it('ignores a depth and a contour on a shape with no top bevel (Visio 16 draws nothing)', async () => {
		const saved = await editVsdx(await plain(), [
			{ ...target, extrusionDepth: 20, contourSize: 6, contourColor: '#008000' },
		]);
		const style = await readStyle(saved.bytes);
		expect(style.extrusion).toBeUndefined();
		expect(style.contour).toBeUndefined();
	});
	it('keeps a bevel of type 1-12 with no size so a solid shape is recognised', async () => {
		const style = await readStyle(await plain(cell('BevelTopType', 1)));
		expect(style.bevel).toEqual({ type: 1, width: 0, height: 0 });
	});
	it('refuses depths and colours outside the supported range', () => {
		expect(() => snapshotFormatting({ ...target, extrusionDepth: -1 })).toThrow();
		expect(() => snapshotFormatting({ ...target, extrusionDepth: 9999 })).toThrow();
		expect(() => snapshotFormatting({ ...target, contourColor: 'red' })).toThrow();
	});
	it('offers the perspective presets of the gallery', () => {
		expect(() =>
			snapshotFormatting({ ...target, rotation3d: 'perspective-heroic-left' }),
		).not.toThrow();
	});
});

describe('Themed bevels', () => {
	const themed = async (sp3d: string) => {
		const theme = generatedTheme().replace(
			'</a:fmtScheme>',
			`<a:effectStyleLst><a:effectStyle><a:effectLst/>${sp3d}</a:effectStyle></a:effectStyleLst></a:fmtScheme>`,
		);
		const bytes = await themeFixture({
			theme,
			contents:
				cell('QuickStyleEffectsMatrix', 1) +
				cell('BevelTopType', 'Themed', 'THEMEVAL()') +
				cell('BevelTopWidth', 'Themed', 'THEMEVAL()') +
				cell('BevelTopHeight', 'Themed', 'THEMEVAL()') +
				cell('BevelContourSize', 'Themed', 'THEMEVAL()') +
				cell('BevelContourColor', 'Themed', 'THEMEVAL()') +
				cell('BevelDepthSize', 0.25),
		});
		return readStyle(bytes);
	};
	it('reads the top bevel, contour and depth colour defer from the effect style sp3d', async () => {
		const style = await themed(
			'<a:sp3d contourW="38100"><a:bevelT w="76200" h="38100" prst="angle"/><a:contourClr><a:srgbClr val="336699"/></a:contourClr></a:sp3d>',
		);
		expect(style.bevel).toEqual({ type: 5, width: 76_200 / 914_400, height: 38_100 / 914_400 });
		expect(style.contour).toEqual({ size: 38_100 / 914_400, color: '#336699' });
		expect(style.extrusion?.depth).toBe(0.25);
	});
	it('draws no bevel when the theme carries only extrusion and contour colours', async () => {
		const style = await themed(
			'<a:sp3d><a:extrusionClr><a:srgbClr val="FFFFFF"/></a:extrusionClr></a:sp3d>',
		);
		expect(style.bevel).toBeUndefined();
		expect(style.contour).toBeUndefined();
		expect(style.extrusion).toBeUndefined();
	});
});
