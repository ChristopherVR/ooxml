import { describe, expect, it } from 'vitest';

import { buildDashArray } from './connector-dash';
import { capCompensatedDashArray, dashPatternUnits, dashPenWidth } from './dash-pattern';
import { getSvgStrokeDasharray } from './element-style-transform';
import { svgLineCap } from './connector-style';

// Lengths measured from PowerPoint 1280x720 exports (flat caps, 6 px and 12 px pens).
describe('PowerPoint dash patterns', () => {
	const presets: Array<[string, string]> = [
		['dash', '24 18'],
		['dot', '6 18'],
		['lgDash', '48 18'],
		['dashDot', '24 18 6 18'],
		['lgDashDot', '48 18 6 18'],
		['lgDashDotDot', '48 18 6 18 6 18'],
		['sysDash', '18 6'],
		['sysDot', '6 6'],
		['sysDashDot', '18 6 6 6'],
		['sysDashDotDot', '18 6 6 6 6 6'],
	];
	it.each(presets)('%s at a 6 px pen', (name, expected) => {
		expect(getSvgStrokeDasharray(name as never, 6)).toBe(expected);
		expect(buildDashArray(name, 6)).toBe(expected);
	});

	it('scales with the pen width rounded to whole pixels', () => {
		expect(dashPenWidth(2.67)).toBe(3);
		expect(dashPenWidth(1.33)).toBe(1);
		expect(dashPenWidth(4.5)).toBe(4);
		expect(dashPenWidth(0.2)).toBe(1);
		expect(getSvgStrokeDasharray('dash', 2.67)).toBe('12 9');
	});

	it('reads custDash in 1000ths of a percent of the pen', () => {
		expect(
			buildDashArray('custom', 6, [
				{ dash: 100000, space: 300000 },
				{ dash: 300000, space: 100000 },
			]),
		).toBe('6 18 18 6');
		expect(dashPatternUnits('custom', [{ dash: 500000, space: 200000 }])).toEqual([5, 2]);
	});

	it('keeps the period but folds round caps into the dash', () => {
		expect(getSvgStrokeDasharray('dash', 6, undefined, 'rnd')).toBe('18 24');
		expect(getSvgStrokeDasharray('dot', 6, undefined, 'rnd')).toBe('0 24');
		expect(getSvgStrokeDasharray('sysDot', 6, undefined, 'rnd')).toBe('0 12');
		expect(capCompensatedDashArray([4, 3], 6, 'sq')).toBe('24 18');
	});

	it('draws interior dashes of dashed square-cap lines flat', () => {
		expect(svgLineCap('sq', true)).toBe('butt');
		expect(svgLineCap('sq')).toBe('square');
		expect(svgLineCap('rnd', true)).toBe('round');
	});

	it('is solid for solid and absent dashes', () => {
		expect(getSvgStrokeDasharray('solid', 6)).toBeUndefined();
		expect(getSvgStrokeDasharray(undefined, 6)).toBeUndefined();
	});
});
