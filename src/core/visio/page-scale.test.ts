import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { cell, fixture, shape, rectangle } from './test-fixtures';
import { normalizeVisioPageGeometry, scaleVisioGeometryPath } from './page-scale';
import { assertViewableDocument } from './ui/scene-validation';
import { visioPageEditToDrawing } from './ui/page-edit';
import evidence from './__fixtures__/page-scales-native.json';

const scaledFixture = (drawingScale = 2, pageScale = 1) =>
	fixture({
		pages: [
			{
				id: '0',
				pageCells: cell('DrawingScale', drawingScale) + cell('PageScale', pageScale),
				contents: `<Shapes>${shape('1', cell('Width', 4) + cell('Height', 2) + cell('PinX', 3) + cell('PinY', 3) + cell('LineWeight', 0.01) + rectangle)}</Shapes>`,
			},
		],
	});

describe('page drawing coordinates to paper inches', () => {
	it.each(evidence.cases)(
		'normalizes geometry independently of style at $drawingScale:$pageScale',
		async (row) => {
			const model = await parseVsdx(await scaledFixture(row.drawingScale, row.pageScale));
			const page = model.pages[0]!,
				rectangle = page.shapes[0]!;
			expect(page.width).toBeCloseTo(row.nativePaperWidth, 12);
			expect(rectangle.width).toBeCloseTo(row.nativeRectangleWidth, 12);
			expect(rectangle.height).toBeCloseTo(row.nativeRectangleHeight, 12);
			expect(rectangle.transform.slice(4)).toEqual([row.nativeRectangleX, row.nativeRectangleY]);
			expect(rectangle.style.lineWidth).toBe(row.nativeLineStroke);
		},
	);
	it('retains the source rotation pin separately from an off-centre local pivot', async () => {
		const model = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						pageCells: cell('DrawingScale', 2) + cell('PageScale', 1),
						contents: `<Shapes>${shape('1', cell('Width', 4) + cell('Height', 2) + cell('PinX', 3) + cell('PinY', 5) + cell('LocPinX', 1) + cell('LocPinY', 1.5) + cell('Angle', Math.PI / 6) + rectangle)}</Shapes>`,
					},
				],
			}),
		);
		const rotated = model.pages[0]!.shapes[0]!;
		expect(rotated.rotation).toEqual({ pinX: 1.5, pinY: 2.5, angle: Math.PI / 6 });
		expect(rotated.transform[4]).not.toBe(rotated.rotation!.pinX);
	});
	it.each(['pinX', 'pinY', 'angle'] as const)(
		'rejects a nonfinite source rotation %s',
		async (field) => {
			const model = await parseVsdx(await scaledFixture());
			model.pages[0]!.shapes[0]!.rotation![field] = Number.NaN;
			expect(() => assertViewableDocument(model)).toThrow('rotation');
		},
	);
	it('scales arc radii and endpoints while keeping angles, flags and exponents correct', () => {
		expect(scaleVisioGeometryPath('M 1e-4 2 A 4 6 45 0 1 8 10 C 1 2 3 4 5 6 Z', 0.5)).toBe(
			'M 0.00005 1 A 2 3 45 0 1 4 5 C 0.5 1 1.5 2 2.5 3 Z',
		);
	});
	it('scales nested placements without changing shared raster bytes or physical metrics', async () => {
		const parent = (await parseVsdx(await scaledFixture(1))).pages[0]!.shapes[0]!;
		const child = structuredClone(parent);
		parent.children = [child];
		child.transform = [0, 1, -1, 0, 6, 8];
		const image = {
			mimeType: 'image/png' as const,
			bytes: new Uint8Array([1, 2]),
			pixelWidth: 1,
			pixelHeight: 1,
			x: 2,
			y: 4,
			width: 6,
			height: 8,
		};
		parent.image = image;
		child.image = image;
		child.style.fillGradient = { type: 'linear', start: [2, 4], end: [6, 8], stops: [] };
		parent.style.fillGradient = {
			type: 'linear',
			start: [0, 1],
			end: [1, 1],
			boundingBoxAngle: -30,
			stops: [],
		};
		const fontSize = child.text.fontSize;
		const margins = { ...child.text.margins };
		normalizeVisioPageGeometry([parent], 0.5, () => {});
		expect(child.transform).toEqual([0, 1, -1, 0, 3, 4]);
		expect(child.style.fillGradient).toMatchObject({ start: [1, 2], end: [3, 4] });
		expect(parent.style.fillGradient).toMatchObject({
			start: [0, 1],
			end: [1, 1],
			boundingBoxAngle: -30,
		});
		expect(parent.image).not.toBe(image);
		expect(child.image).not.toBe(image);
		expect(child.image).toMatchObject({ x: 1, y: 2, width: 3, height: 4 });
		expect(parent.image!.bytes).toBe(image.bytes);
		expect(image.width).toBe(6);
		expect(child.text.fontSize).toBe(fontSize);
		expect(child.text.margins).toEqual(margins);
	});
	it.each([0, -2])('reports invalid scale %s without dividing the scene by it', async (scale) => {
		const model = await parseVsdx(await scaledFixture(scale));
		expect(model.pages[0]!.width).toBe(8.5);
		expect(model.diagnostics.some((d) => d.code === 'invalid-page-drawing-scale')).toBe(true);
	});
	it('round-trips page-inch UI geometry while keeping raw editing in drawing inches', async () => {
		const original = await scaledFixture();
		const page = (await parseVsdx(original)).pages[0]!;
		const command = visioPageEditToDrawing(page, {
			type: 'resize-shape',
			pageId: '0',
			shapeId: '1',
			width: 3,
			height: 2,
		});
		expect(command).toMatchObject({ width: 6, height: 4 });
		const saved = await editVsdx(original, [command]);
		const reloaded = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect([reloaded.width, reloaded.height]).toEqual([3, 2]);
		const before = await JSZip.loadAsync(original),
			after = await JSZip.loadAsync(saved.bytes);
		for (const part of Object.keys(before.files).filter(
			(part) => !before.files[part]!.dir && !saved.changedParts.includes(part),
		))
			expect(await after.file(part)!.async('uint8array')).toEqual(
				await before.file(part)!.async('uint8array'),
			);
		expect(await after.file('visio/pages/pages.xml')!.async('string')).toContain(
			'N="DrawingScale" V="2"',
		);
		const direct = await editVsdx(original, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 6, height: 4 },
		]);
		expect((await parseVsdx(direct.bytes)).pages[0]!.shapes[0]!.width).toBe(3);
	});
});

const native = process.env.VISIO_NATIVE_PAGE_SCALES_DIR;
describe.skipIf(!native)('native-authored page scale corpus', () => {
	it('matches geometry and physical text/stroke metrics for every native page', async () => {
		const model = await parseVsdx(await readFile(resolve(native!, 'page-scales.vsdx')));
		for (const row of evidence.cases) {
			const page = model.pages.find((p) => p.id === row.pageId)!;
			const rectangle = page.shapes.find((s) => s.id === row.rectangleId)!;
			const rounded = page.shapes.find((s) => s.id === row.roundedId)!;
			expect(Number(rounded.geometry[0]!.path.match(/A ([^ ]+)/)![1])).toBeCloseTo(
				row.nativeRoundedRadius,
				12,
			);
			expect(page.width).toBeCloseTo(row.nativePaperWidth, 12);
			expect(rectangle.width).toBeCloseTo(row.nativeRectangleWidth, 12);
			expect(rectangle.height).toBeCloseTo(row.nativeRectangleHeight, 12);
			expect(rectangle.transform.slice(4)).toEqual([row.nativeRectangleX, row.nativeRectangleY]);
			expect(rectangle.text.fontSize).toBe(row.fontSize);
			expect(rectangle.text.margins.left + rectangle.text.paragraphs![0]!.indentLeft).toBe(
				row.nativeTextX,
			);
		}
	});
});
