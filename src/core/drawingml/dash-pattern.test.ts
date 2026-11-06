import { describe, expect, it } from 'vitest';

import {
	capCompensatedDashArray,
	dashPatternUnits,
	dashPenWidth,
	isDashedStroke,
} from './dash-pattern';

describe('drawingml dash pattern', () => {
	it('rounds the pen to whole pixels, ties down, at least one', () => {
		expect(dashPenWidth(0.2)).toBe(1);
		expect(dashPenWidth(2.67)).toBe(3);
		expect(dashPenWidth(4.5)).toBe(4);
	});

	it('uses the measured preset multiples', () => {
		expect(capCompensatedDashArray(dashPatternUnits('dash')!, 6)).toBe('24 18');
		expect(capCompensatedDashArray(dashPatternUnits('sysDot')!, 6)).toBe('6 6');
		expect(dashPatternUnits('solid')).toBeUndefined();
	});

	it('reads custDash in 1000ths of a percent', () => {
		const pattern = dashPatternUnits('custom', [{ dash: 300000, space: 100000 }]);
		expect(capCompensatedDashArray(pattern!, 2)).toBe('6 2');
	});

	it('gives a round dash up to its caps', () => {
		expect(capCompensatedDashArray(dashPatternUnits('dot')!, 6, 'rnd')).toBe('0 24');
	});

	it('knows which tokens are dashed', () => {
		expect(isDashedStroke('dot')).toBe(true);
		expect(isDashedStroke('solid')).toBe(false);
		expect(isDashedStroke(undefined)).toBe(false);
	});
});
