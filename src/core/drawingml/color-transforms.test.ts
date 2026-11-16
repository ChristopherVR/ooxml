import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/index';
import {
	applyDrawingColorTransformList,
	declaredDrawingColorAlpha,
	drawingColorAlpha,
} from './color-transforms';
import { parseDrawingColorIn, resolveDrawingColor } from './drawing-color';
import { DRAWING_PRESET_COLORS, DRAWING_SYSTEM_COLORS } from './preset-colors';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const colorOf = (xml: string) =>
	parseDrawingColorIn(
		parseXml(`<a:solidFill xmlns:a="${A}">${xml}</a:solidFill>`).documentElement,
	)!;
const hex = (xml: string) => resolveDrawingColor(colorOf(xml))?.hex;

describe('shared colour transforms', () => {
	it('applies the RGB channel transforms after the HSL round trip', () => {
		expect(hex('<a:srgbClr val="804020"><a:red val="100000"/></a:srgbClr>')).toBe('#FF4020');
		expect(hex('<a:srgbClr val="804020"><a:greenMod val="200000"/></a:srgbClr>')).toBe('#808020');
		expect(hex('<a:srgbClr val="804020"><a:blueOff val="-50000"/></a:srgbClr>')).toBe('#804000');
		// lumMod runs first, then the red channel is replaced.
		expect(hex('<a:srgbClr val="804020"><a:lumMod val="50000"/><a:red val="0"/></a:srgbClr>')).toBe(
			'#002010',
		);
	});

	it('gamma-encodes and decodes before the other transforms', () => {
		expect(hex('<a:srgbClr val="808080"><a:gamma/></a:srgbClr>')).toBe('#BCBCBC');
		expect(hex('<a:srgbClr val="808080"><a:invGamma/></a:srgbClr>')).toBe('#373737');
	});

	it('reads each transform once, the first of its name', () => {
		const once = applyDrawingColorTransformList({ r: 68, g: 114, b: 196 }, [
			{ name: 'tint', value: '60000' },
		]);
		const twice = applyDrawingColorTransformList({ r: 68, g: 114, b: 196 }, [
			{ name: 'tint', value: '60000' },
			{ name: 'tint', value: '50000' },
		]);
		expect(twice).toEqual(once);
	});

	it('reports only unknown transforms as unapplied', () => {
		const resolved = resolveDrawingColor(
			colorOf(
				'<a:srgbClr val="000000"><a:redMod val="1"/><a:gamma/><a:fooOff val="1"/></a:srgbClr>',
			),
		);
		expect(resolved?.unapplied).toEqual(['fooOff']);
	});
});

describe('preset and system colours', () => {
	it('resolves every preset name, case-insensitively', () => {
		expect(hex('<a:prstClr val="cornflowerBlue"/>')).toBe('#6495ED');
		expect(hex('<a:prstClr val="dkSlateGray"/>')).toBe(DRAWING_PRESET_COLORS.darkslategray);
		expect(hex('<a:prstClr val="nope"/>')).toBeUndefined();
	});

	it('falls back to the system table when a sysClr has no usable lastClr', () => {
		expect(hex('<a:sysClr val="windowText" lastClr="123456"/>')).toBe('#123456');
		expect(hex('<a:sysClr val="highlight"/>')).toBe(DRAWING_SYSTEM_COLORS.highlight);
		expect(hex('<a:sysClr val="window" lastClr="zz"/>')).toBe('#FFFFFF');
	});
});

describe('colour alpha', () => {
	const t = (...pairs: [string, string][]) => pairs.map(([name, value]) => ({ name, value }));

	it('combines alpha, alphaMod and alphaOff, clamping after each step', () => {
		expect(drawingColorAlpha(t(['alpha', '40000'], ['alphaMod', '200000']))).toBeCloseTo(0.8);
		expect(drawingColorAlpha(t(['alpha', '90000'], ['alphaMod', '200000']))).toBe(1);
		expect(drawingColorAlpha(t(['alpha', '50000'], ['alphaOff', '-25000']))).toBeCloseTo(0.25);
		expect(drawingColorAlpha([])).toBe(1);
	});

	it('reads the first of a repeated transform, as the resolver does', () => {
		expect(drawingColorAlpha(t(['alpha', '30000'], ['alpha', '90000']))).toBeCloseTo(0.3);
	});

	it('reports no declared alpha when no readable alpha transform is present', () => {
		expect(declaredDrawingColorAlpha(t(['lumMod', '50000']))).toBeUndefined();
		expect(declaredDrawingColorAlpha(t(['alpha', 'x']))).toBeUndefined();
		expect(declaredDrawingColorAlpha(t(['alphaOff', '-10000']))).toBeCloseTo(0.9);
	});

	it('is the alpha the resolver returns', () => {
		const color = colorOf('<a:schemeClr val="accent1"><a:alpha val="35000"/></a:schemeClr>');
		expect(resolveDrawingColor(color, { scheme: () => '#112233' })?.alpha).toBe(
			declaredDrawingColorAlpha(color.transforms),
		);
	});
});
