import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures';

const stops = (paint: 'Fill' | 'Line', alpha = 0) =>
	section(
		`${paint}Gradient`,
		row(
			0,
			'',
			cell('GradientStopPosition', 0) +
				cell('GradientStopColor', '#ff0000') +
				cell('GradientStopColorTrans', alpha),
		) +
			row(
				1,
				'',
				cell('GradientStopPosition', 1) +
					cell('GradientStopColor', '#0000ff') +
					cell('GradientStopColorTrans', alpha),
			),
	);
async function source(overrides = '', pageCells = '') {
	return fixture({
		pages: [
			{
				id: '0',
				pageCells,
				contents: `<Shapes>${shape(
					'1',
					rectangle +
						cell('Width', 4) +
						cell('Height', 2) +
						cell('LinePattern', 1) +
						cell('LineColor', '#123456') +
						cell('LineGradientEnabled', 1) +
						cell('LineGradientDir', 0) +
						cell('LineGradientAngle', 0) +
						cell('RotateGradientWithShape', 1) +
						cell('UseGroupGradient', 0) +
						stops('Line') +
						overrides,
				)}</Shapes>`,
			},
		],
	});
}

it('parses independent saved fill and line gradients through the shared parser', async () => {
	const parsed = await parseVsdx(
		await source(
			cell('FillGradientEnabled', 1) +
				cell('FillGradientDir', 0) +
				cell('FillGradientAngle', 0) +
				stops('Fill', 0.5),
		),
	);
	const style = parsed.pages[0]!.shapes[0]!.style;
	expect(style.lineGradient).toMatchObject({
		type: 'linear',
		start: [0, 1],
		end: [4, 1],
		interpolation: 'sigma-gamma22',
	});
	expect(style.lineGradient?.stops.map((stop) => stop.opacity)).toEqual([1, 1]);
	expect(style.fillGradient?.stops.map((stop) => stop.opacity)).toEqual([0.5, 0.5]);
	expect(style.lineColor).toBe('#123456');
	expect(parsed.diagnostics.some((item) => item.code === 'unverified-gradient-raster')).toBe(true);
	expect(parsed.diagnostics.some((item) => item.code === 'unsupported-gradient')).toBe(false);
});

it.each([1, 8, 13])(
	'reports unsupported line direction %i without using fill geometry',
	async (direction) => {
		const parsed = await parseVsdx(await source(cell('LineGradientDir', direction)));
		expect(parsed.pages[0]!.shapes[0]!.style.lineGradient).toBeUndefined();
		expect(parsed.diagnostics.some((item) => item.code === 'unsupported-saved-line-gradient')).toBe(
			true,
		);
	},
);

it('disables line gradient paint when the line is hidden', async () => {
	const parsed = await parseVsdx(await source(cell('LinePattern', 0)));
	expect(parsed.pages[0]!.shapes[0]!.style.lineGradient).toBeUndefined();
});

it.each([0, Math.PI / 2, Math.PI / 4])(
	'retains the paint angle for a height-zero line at %s',
	async (angle) => {
		const parsed = await parseVsdx(
			await source(cell('Height', 0) + cell('LineGradientAngle', angle)),
		);
		expect(parsed.pages[0]!.shapes[0]!.style.lineGradient).toMatchObject({
			type: 'linear',
			start: [0, 1],
			end: [1, 1],
			boundingBoxAngle: (-angle * 180) / Math.PI,
		});
		expect(parsed.diagnostics.some((item) => item.code === 'unsupported-saved-line-gradient')).toBe(
			false,
		);
	},
);

it('rejects a point-sized stroke and retains zero-area fill rejection', async () => {
	const point = await parseVsdx(await source(cell('Width', 0) + cell('Height', 0)));
	expect(point.pages[0]!.shapes[0]!.style.lineGradient).toBeUndefined();
	const fill = await parseVsdx(
		await source(
			cell('Height', 0) +
				cell('FillGradientEnabled', 1) +
				cell('FillGradientDir', 0) +
				cell('FillGradientAngle', 0) +
				stops('Fill'),
		),
	);
	expect(fill.pages[0]!.shapes[0]!.style.fillGradient).toBeUndefined();
	expect(fill.pages[0]!.shapes[0]!.style.lineGradient).toBeDefined();
});

it('reports the solid arrow marker fallback', async () => {
	const parsed = await parseVsdx(await source(cell('EndArrow', 1)));
	expect(parsed.pages[0]!.shapes[0]!.style.lineGradient).toBeDefined();
	expect(parsed.diagnostics.some((item) => item.code === 'unsupported-gradient-arrows')).toBe(true);
});

it('scales physical line endpoints with page geometry', async () => {
	const parsed = await parseVsdx(await source('', cell('DrawingScale', 2) + cell('PageScale', 1)));
	expect(parsed.pages[0]!.shapes[0]!.style.lineGradient).toMatchObject({
		start: [0, 0.5],
		end: [2, 0.5],
	});
});

it('uses the native colored-layer stroke override without destroying source stops', async () => {
	const bytes = await source(
		cell('LayerMember', '0'),
		section('Layer', row(0, '', cell('Color', '#00ff00') + cell('ColorTrans', 0.4))),
	);
	const parsed = await parseVsdx(bytes);
	const style = parsed.pages[0]!.shapes[0]!.style;
	expect(style.lineGradient).toBeUndefined();
	expect(style.lineColor).toBe('#00ff00');
	expect(style.lineOpacity).toBe(0.6);
	const saved = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 2 },
	]);
	expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style).toEqual(style);
});

it('preserves saved stroke paint through editing and package reparse', async () => {
	const bytes = await source();
	const original = await parseVsdx(bytes);
	const saved = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 2 },
	]);
	expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style.lineGradient).toEqual(
		original.pages[0]!.shapes[0]!.style.lineGradient,
	);
});

for (const name of ['VISIO_NATIVE_LINE_GRADIENT_DIR', 'VISIO_NATIVE_1D_LINE_GRADIENT_DIR']) {
	const nativeDirectory = process.env[name];
	it.skipIf(!nativeDirectory)(
		`preserves all genuine native line stop caches through editing (${name})`,
		async () => {
			const bytes = new Uint8Array(await readFile(join(nativeDirectory!, 'gradient-raster.vsdx')));
			const original = await parseVsdx(bytes);
			const shapes = original.pages[0]!.shapes;
			const oneDimensional = name === 'VISIO_NATIVE_1D_LINE_GRADIENT_DIR';
			expect(shapes).toHaveLength(oneDimensional ? 5 : 4);
			if (oneDimensional) {
				expect(shapes.slice(0, 4).every((shape) => shape.height === 0)).toBe(true);
				await expect(
					editVsdx(bytes, [
						{
							type: 'move-shape',
							pageId: original.pages[0]!.id,
							shapeId: shapes[0]!.id,
							x: 4,
							y: 3,
						},
					]),
				).rejects.toThrow('Only local 2D shapes');
			}
			expect(shapes.slice(0, 4).every((shape) => shape.style.lineGradient?.type === 'linear')).toBe(
				true,
			);
			expect(
				shapes.filter((shape) => shape.style.lineGradient?.interpolation === 'sigma-gamma22'),
			).toHaveLength(1);
			const saved = await editVsdx(bytes, [
				{
					type: 'move-shape',
					pageId: original.pages[0]!.id,
					shapeId: shapes[oneDimensional ? 4 : 0]!.id,
					x: 4,
					y: 3,
				},
			]);
			expect(
				(await parseVsdx(saved.bytes)).pages[0]!.shapes.map((shape) => shape.style.lineGradient),
			).toEqual(shapes.map((shape) => shape.style.lineGradient));
		},
	);
}
