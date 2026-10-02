import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { linearGradientEndpoints } from './theme-gradient.js';
import { generatedTheme, themeFixture } from './theme-fixtures.js';
import { cell, row, section } from './test-fixtures.js';

const stop = (position: number, color = '<a:schemeClr val="phClr"/>') =>
	`<a:gs pos="${position}">${color}</a:gs>`;
const gradient = (
	options: {
		angle?: string;
		extra?: string;
		attributes?: string;
		stops?: string;
		linear?: string;
	} = {},
) =>
	`<a:gradFill rotWithShape="1" ${options.attributes ?? ''}><a:gsLst>${options.stops ?? stop(0) + stop(100000, '<a:srgbClr val="FFFFFF"/>')}</a:gsLst>${options.linear ?? `<a:lin ang="${options.angle ?? '5400000'}" scaled="0"/>`}${options.extra ?? ''}</a:gradFill>`;
async function parse(fill = gradient(), contents = '') {
	const document = await parseVsdx(
		await themeFixture({
			theme: generatedTheme({ fill }),
			contents: cell('Width', 4) + cell('Height', 2) + contents,
		}),
	);
	return { shape: document.pages[0]!.shapes[0]!, diagnostics: document.diagnostics };
}
describe('normalized Visio theme gradients', () => {
	it('normalizes clockwise vertical gradients into y-up inches', async () => {
		const result = await parse();
		expect(result.shape.style.fill).toBe('#123456');
		expect(result.shape.style.fillGradient).toEqual({
			type: 'linear',
			start: [2, 2],
			end: [2, 0],
			stops: [
				{ offset: 0, color: '#123456', opacity: 1 },
				{ offset: 1, color: '#ffffff', opacity: 1 },
			],
		});
		expect(result.diagnostics.some((item) => item.code.includes('gradient'))).toBe(false);
	});
	it('projects nonsquare bounds instead of stretching the direction', () => {
		const endpoints = linearGradientEndpoints(4, 2, 45 * 60_000);
		expect(endpoints.start[0]).toBeCloseTo(0.5);
		expect(endpoints.start[1]).toBeCloseTo(2.5);
		expect(endpoints.end[0]).toBeCloseTo(3.5);
		expect(endpoints.end[1]).toBeCloseTo(-0.5);
	});
	it.each([
		[0, [0, 1], [4, 1]],
		[90, [2, 2], [2, 0]],
		[180, [4, 1], [0, 1]],
		[270, [2, 0], [2, 2]],
	])('normalizes cardinal angle %i', (angle, start, end) => {
		expect(linearGradientEndpoints(4, 2, Number(angle) * 60_000)).toEqual({ start, end });
	});
	it('resolves placeholder, scheme, shade/tint and alpha stop properties', async () => {
		const result = await parse(
			gradient({
				stops:
					stop(
						0,
						'<a:srgbClr val="FFFFFF"><a:shade val="50000"/><a:alpha val="50000"/></a:srgbClr>',
					) +
					stop(24000, '<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>') +
					stop(54000, '<a:schemeClr val="lt1"/>'),
			}),
		);
		expect(result.shape.style.fillGradient?.stops).toEqual([
			{ offset: 0, color: '#bcbcbc', opacity: 0.5 },
			{ offset: 0.24, color: '#bcbcbc', opacity: 1 },
			{ offset: 0.54, color: '#f0e0d0', opacity: 1 },
		]);
	});
	it('ignores scaled, flip and tileRect as required specifically by MS-VSDX', async () => {
		const baseline = await parse();
		const ignored = await parse(
			gradient({
				linear: '<a:lin ang="5400000" scaled="1"/>',
				attributes: 'flip="xy"',
				extra: '<a:tileRect l="40000" r="-20000"/>',
			}),
		);
		expect(ignored.shape.style.fillGradient).toEqual(baseline.shape.style.fillGradient);
	});
	it('preserves coincident stops and the theme last-stop position', async () => {
		const result = await parse(gradient({ stops: stop(0) + stop(0) + stop(82000) }));
		expect(result.shape.style.fillGradient?.stops.map((item) => item.offset)).toEqual([0, 0, 0.82]);
	});
	it.each([
		gradient({ angle: '21600000' }),
		gradient({ angle: '5bad' }),
		gradient({ linear: '<a:path path="circle"/>' }),
		gradient().replace('rotWithShape="1"', 'rotWithShape="0"'),
		gradient({ stops: stop(0) }),
		gradient({ stops: stop(80000) + stop(20000) }),
		gradient({ stops: stop(0) + stop(100001) }),
		gradient({ stops: stop(0).repeat(33) }),
		gradient({
			stops:
				stop(0) + stop(100000, '<a:schemeClr val="phClr"><a:satMod val="50000"/></a:schemeClr>'),
		}),
	])('leaves unsupported gradients as diagnosed solid fallbacks', async (fill) => {
		const result = await parse(fill);
		expect(result.shape.style.fillGradient).toBeUndefined();
		expect(result.shape.style.fill).toBe('#123456');
		expect(result.diagnostics.some((item) => item.code === 'unsupported-theme-gradient')).toBe(
			true,
		);
	});
	it.each([
		cell('FillGradientAngle', 0),
		cell('Width', 0),
		section('FillGradient', row(0, '', cell('GradientStopColor', '#ff0000'))),
	])('does not overwrite saved local gradient overrides', async (contents) => {
		const result = await parse(gradient(), contents);
		expect(result.shape.style.fillGradient).toBeUndefined();
		expect(result.diagnostics.some((item) => item.code === 'unsupported-theme-gradient')).toBe(
			true,
		);
	});
	it('honors cached disabled gradients and solid literal overrides', async () => {
		expect(
			(await parse(gradient(), cell('FillGradientEnabled', 0))).shape.style.fillGradient,
		).toBeUndefined();
		expect(
			(await parse(gradient(), cell('FillForegnd', '#112233'))).shape.style.fillGradient,
		).toBeUndefined();
	});
});
