import { describe, expect, it } from 'vitest';

import { namedFontVerticalMetrics } from './vertical-metrics';

describe('namedFontVerticalMetrics', () => {
	it.each([
		'MS Mincho',
		'MS PMincho',
		'MS Gothic',
		'MS PGothic',
		'MS UI Gothic',
		'SimSun',
		'NSimSun',
	])('knows %s', (name) => {
		const metrics = namedFontVerticalMetrics(` ${name.toUpperCase()} `);
		expect(metrics?.ascent).toBeCloseTo(0.859, 3);
		expect(metrics?.descent).toBeCloseTo(0.141, 3);
		expect((metrics?.ascent ?? 0) + (metrics?.descent ?? 0)).toBeCloseTo(1, 10);
	});
	it('knows no other font', () => {
		expect(namedFontVerticalMetrics('Calibri')).toBeUndefined();
		expect(namedFontVerticalMetrics(undefined)).toBeUndefined();
	});
});
