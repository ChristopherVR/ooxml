import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures.js';

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
describe('saved horizontal fill gradients', () => {
	it('normalizes complete saved stop caches into local inches', async () => {
		const { style, diagnostics } = await parse();
		expect(style.fillGradient).toEqual({
			type: 'linear',
			start: [0, 1],
			end: [4, 1],
			stops: [
				{ offset: 0, color: '#ff0000', opacity: 1 },
				{ offset: 1, color: '#0000ff', opacity: 1 },
			],
		});
		expect(diagnostics.some((d) => d.code.includes('gradient'))).toBe(false);
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
		const { style } = await parse('', stop(0, 0.2) + stop(1, 0.2) + stop(2, 0.8));
		expect(style.fillGradient?.stops.map((s) => s.offset)).toEqual([0.2, 0.2, 0.8]);
	});
	it('ignores rows beyond the first ten and deleted rows', async () => {
		const rows = Array.from({ length: 10 }, (_, i) => stop(i, i / 10)).join('');
		const { style } = await parse('', '<Row IX="50" Del="1"/>' + rows + stop(10, 'bad'));
		expect(style.fillGradient?.stops).toHaveLength(10);
	});
	it.each([
		cell('FillGradientDir', 3),
		cell('FillGradientDir', 'Themed'),
		cell('FillGradientAngle', Math.PI / 2),
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
