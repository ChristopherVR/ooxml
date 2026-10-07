import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from './edit';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures';

const stop = (
	index: number,
	position: string | number,
	color = '#ff0000',
	transparency: string | number = 0,
) =>
	row(
		index,
		'',
		cell('GradientStopPosition', position) +
			cell('GradientStopColor', color) +
			cell('GradientStopColorTrans', transparency),
	);
const settings =
	cell('FillPattern', 1) +
	cell('FillForegnd', '#123456') +
	cell('FillGradientEnabled', 1) +
	cell('FillGradientDir', 0) +
	cell('FillGradientAngle', 0) +
	cell('RotateGradientWithShape', 1) +
	cell('UseGroupGradient', 0);

for (const [name, type, count] of [
	['LINEAR_VERTICAL', 'linear', 6],
	['LINEAR_REVERSE', 'linear', 6],
	['LINEAR_OBLIQUE', 'linear', 6],
	['LINEAR_OBLIQUE_ALPHA', 'linear', 6],
	['SAVED_RADIAL', 'radial', 7],
	['SAVED_RADIAL_ALPHA', 'radial', 7],
	['SAVED_REGIONS', 'regions', 5],
	['SAVED_REGIONS_ALPHA', 'regions', 5],
] as const) {
	const directory = process.env[`VISIO_NATIVE_${name}_DIR`];
	it.skipIf(!directory)(`preserves genuine ${name} gradients through core edits`, async () => {
		const bytes = new Uint8Array(await readFile(join(directory!, 'fill-patterns.vsdx')));
		const original = await parseVsdx(bytes);
		expect(original.pages).toHaveLength(count);
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
		for (let index = 0; index < count; index++) {
			expect(original.pages[index]!.shapes[0]!.style.fillGradient?.type).toBe(type);
			expect(reopened.pages[index]!.shapes[0]!.style.fillGradient).toEqual(
				original.pages[index]!.shapes[0]!.style.fillGradient,
			);
		}
		await writeFile(join(directory!, 'core-fill-patterns.vsdx'), saved.bytes);
	});
}
async function parse(overrides = '', stops = stop(0, 0) + stop(1, 1, '#0000ff'), document = '') {
	const parsed = await parseVsdx(
		await fixture({
			document,
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + cell('Width', 4) + cell('Height', 2) + settings + overrides + section('FillGradient', stops))}</Shapes>`,
				},
			],
		}),
	);
	return { style: parsed.pages[0]!.shapes[0]!.style, diagnostics: parsed.diagnostics };
}

it.each([
	[1, [4, 0], Math.hypot(4, 2)],
	[2, [0, 0], Math.hypot(4, 2)],
	[3, [2, 1], Math.hypot(2, 1)],
	[4, [2, 2], Math.hypot(2, 2)],
	[5, [2, 0], Math.hypot(2, 2)],
	[6, [4, 2], Math.hypot(4, 2)],
	[7, [0, 2], Math.hypot(4, 2)],
])(
	'normalizes native saved radial direction %s independently of the linear angle',
	async (direction, center, radius) => {
		const { style, diagnostics } = await parse(
			cell('FillGradientDir', direction as number) + cell('FillGradientAngle', 'Themed'),
		);
		expect(style.fillGradient).toMatchObject({
			type: 'radial',
			coordinateSpace: 'local',
			center,
			radius,
		});
		expect(style.fillGradient?.stops).toHaveLength(2);
		expect(diagnostics.some((item) => item.code === 'unsupported-saved-fill-gradient')).toBe(false);
	},
);

it.each([
	[Math.PI / 2, [2, 2], [2, 0]],
	[(3 * Math.PI) / 2, [2, 0], [2, 2]],
	[-Math.PI / 2, [2, 0], [2, 2]],
	[1.5707963267949, [2, 2], [2, 0]],
])('normalizes native vertical angle %s with saved decimal rounding', async (angle, start, end) => {
	const { style } = await parse(cell('FillGradientAngle', angle as number));
	expect(style.fillGradient).toMatchObject({ type: 'linear', start, end });
});

it.each([
	[8, [270, 180]],
	[9, [270, 360]],
	[10, [180, 360, 270, 90]],
	[11, [180, 90]],
	[12, [90, 360]],
])('normalizes native saved rectangular direction %s', async (direction, angles) => {
	const { style, diagnostics } = await parse(
		cell('FillGradientDir', direction as number) + cell('FillGradientAngle', 'Themed'),
		stop(0, 0, '#ff0000', 0.2) + stop(1, 0.5, '#00ff00', 0.4) + stop(2, 1, '#0000ff', 0.5),
	);
	const gradient = style.fillGradient;
	expect(gradient?.type).toBe('regions');
	if (gradient?.type !== 'regions') throw new Error('Expected a rectangular gradient.');
	expect(gradient.regions.map((region) => region.angle)).toEqual(angles);
	expect(gradient.stops).toEqual([
		{ offset: 0, color: '#ff0000', opacity: 0.8 },
		{ offset: 0.5, color: '#00ff00', opacity: 0.6 },
		{ offset: 1, color: '#0000ff', opacity: 0.5 },
	]);
	expect(diagnostics.some((item) => item.code === 'unsupported-saved-fill-gradient')).toBe(false);
});

it.each([
	[Math.PI / 6, -30],
	[0.5235987755983, -30],
	[-Math.PI / 4, -315],
	[(5 * Math.PI) / 4, -225],
])('retains native bounding-box rotation for oblique angle %s', async (angle, expected) => {
	const { style, diagnostics } = await parse(cell('FillGradientAngle', angle));
	expect(style.fillGradient).toMatchObject({
		type: 'linear',
		start: [0, 1],
		end: [1, 1],
		boundingBoxAngle: expected,
	});
	expect(diagnostics.some((item) => item.code === 'unsupported-saved-fill-gradient')).toBe(false);
});

it('ignores inherited wholly themed stop tails while diagnosing partly themed active rows', async () => {
	const tail = Array.from({ length: 8 }, (_, index) =>
		stop(index + 2, 'Themed', 'Themed', 'Themed'),
	).join('');
	const result = await parse('', stop(0, 0) + stop(1, 1, '#0000ff') + tail);
	expect(result.style.fillGradient?.stops).toHaveLength(2);
	expect(result.diagnostics.some((item) => item.code === 'unsupported-saved-fill-gradient')).toBe(
		false,
	);
	const incomplete = await parse(
		'',
		stop(0, 0) + stop(1, 1, '#0000ff') + stop(2, 0.5, 'Themed', 'Themed'),
	);
	expect(incomplete.style.fillGradient).toBeUndefined();
	expect(
		incomplete.diagnostics.some((item) => item.code === 'unsupported-saved-fill-gradient'),
	).toBe(true);
});
describe('saved horizontal fill gradients', () => {
	it('normalizes complete saved stop caches into local inches', async () => {
		const { style, diagnostics } = await parse();
		expect(style.fillGradient).toEqual({
			type: 'linear',
			interpolation: 'sigma-gamma22',
			start: [0, 1],
			end: [4, 1],
			stops: [
				{ offset: 0, color: '#ff0000', opacity: 1 },
				{ offset: 1, color: '#0000ff', opacity: 1 },
			],
		});
		expect(diagnostics.some((d) => d.code.startsWith('unsupported-'))).toBe(false);
		expect(diagnostics.map((d) => d.code)).toContain('unverified-gradient-raster');
	});
	it.each([Math.PI, -Math.PI, 3 * Math.PI, 3.14159265358979])(
		'accepts sign-independent reversed horizontal angle %s',
		async (angle) => {
			const { style } = await parse(cell('FillGradientAngle', angle));
			expect(style.fillGradient).toMatchObject({ start: [4, 1], end: [0, 1] });
		},
	);
	it.each([2 * Math.PI, -2 * Math.PI])('normalizes complete turns %s', async (angle) => {
		const { style } = await parse(cell('FillGradientAngle', angle));
		expect(style.fillGradient).toMatchObject({ start: [0, 1], end: [4, 1] });
	});
	it('uses per-stop transparency without multiplying foreground transparency twice', async () => {
		const { style } = await parse(
			cell('FillForegndTrans', 0.6),
			stop(0, 0, '#ff0000', 0.6) + stop(1, 1, '#0000ff', 0.2),
		);
		expect(style.fillOpacity).toBe(1);
		expect(style.fillGradient?.stops.map((s) => s.opacity)).toEqual([0.4, 0.8]);
	});
	it('resolves document palette and standard palette stop colors', async () => {
		const { style } = await parse(
			'',
			stop(0, 0, '1') + stop(1, 1, '24'),
			'<Colors><ColorEntry IX="24" RGB="#aabbcc"/></Colors>',
		);
		expect(style.fillGradient?.stops.map((s) => s.color)).toEqual(['#ffffff', '#aabbcc']);
	});
	it('honors enabled saved stops even when the legacy fill pattern is non-solid', async () => {
		const { style, diagnostics } = await parse(cell('FillPattern', 30));
		expect(style.fillGradient).toBeDefined();
		expect(diagnostics.some((d) => d.code === 'unsupported-fill-pattern')).toBe(false);
	});
	it('preserves coincident and non-boundary stops', async () => {
		const { style, diagnostics } = await parse('', stop(0, 0.2) + stop(1, 0.2) + stop(2, 0.8));
		expect(style.fillGradient?.stops.map((s) => s.offset)).toEqual([0.2, 0.2, 0.8]);
		expect(diagnostics.map((d) => d.code)).not.toContain('unverified-gradient-raster');
	});
	it('ignores rows beyond the first ten and deleted rows', async () => {
		const rows = Array.from({ length: 10 }, (_, i) => stop(i, i / 10)).join('');
		const { style } = await parse('', '<Row IX="50" Del="1"/>' + rows + stop(10, 'bad'));
		expect(style.fillGradient?.stops).toHaveLength(10);
	});
	it.each([
		cell('FillGradientDir', -1),
		cell('FillGradientDir', 13),
		cell('FillGradientDir', 1.5),
		cell('FillGradientDir', 'Themed'),
		cell('FillGradientAngle', 'invalid'),
		cell('FillGradientAngle', 'Themed'),
		cell('FillGradientAngle', 'bad'),
		cell('RotateGradientWithShape', 0),
		cell('RotateGradientWithShape', 'Themed'),
		cell('UseGroupGradient', 1),
		cell('Width', 0),
		cell('Height', 0),
	])('keeps unsupported settings diagnosed: %s', async (overrides) => {
		const { style, diagnostics } = await parse(overrides);
		expect(style.fillGradient).toBeUndefined();
		expect(style.fill).toBe('#123456');
		expect(diagnostics.map((d) => d.code)).toContain('unsupported-saved-fill-gradient');
	});
	it.each([
		stop(0, 0),
		stop(0, 0.8) + stop(1, 0.2),
		stop(0, 0) + stop(1, 1) + row(2, '', ''),
		stop(0, -0.1) + stop(1, 1),
		stop(0, 0) + stop(1, 1.1),
		stop(0, 0) + stop(1, 1, 'Themed'),
		stop(0, 0) + stop(1, 1, 'not-a-color'),
		stop(0, 0) + stop(1, 1, '#0000ff', 1.1),
		stop(0, 0) + row(1, '', cell('GradientStopColor', '#0000ff') + cell('GradientStopPosition', 1)),
	])('rejects incomplete or malformed stop rows: %s', async (stops) => {
		const { style, diagnostics } = await parse('', stops);
		expect(style.fillGradient).toBeUndefined();
		expect(diagnostics.map((d) => d.code)).toContain('unsupported-saved-fill-gradient');
	});
	it.each([cell('FillPattern', 0), cell('FillGradientEnabled', 0)])(
		'respects disabled fills/gradients: %s',
		async (overrides) => {
			const { style } = await parse(overrides);
			expect(style.fillGradient).toBeUndefined();
		},
	);
});
