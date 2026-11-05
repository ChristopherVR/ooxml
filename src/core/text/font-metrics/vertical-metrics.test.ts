import { describe, expect, it } from 'vitest';

import { namedFontVerticalMetrics, officeFontVerticalMetrics } from './vertical-metrics';

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

describe('officeFontVerticalMetrics', () => {
	it('knows the Windows and Office fonts the scan measured, ignoring case', () => {
		const ocr = officeFontVerticalMetrics(' ocr a extended ');
		expect(ocr?.ascent).toBeCloseTo(0.858, 3);
		expect(ocr?.descent).toBeCloseTo(0.177, 3);
		expect(ocr?.lineGap).toBeUndefined();
	});
	it('uses the typographic metrics and line gap of a font that asks for them', () => {
		expect(officeFontVerticalMetrics('Nirmala Text Semilight')).toStrictEqual({
			ascent: 0.7759,
			descent: 0.2241,
			lineGap: 0.8623,
		});
		expect(officeFontVerticalMetrics('Liberation Sans Narrow')?.lineGap).toBeCloseTo(0.131, 3);
	});
	it('knows no other font', () => {
		expect(officeFontVerticalMetrics('Foo Sans')).toBeUndefined();
		expect(officeFontVerticalMetrics(undefined)).toBeUndefined();
	});
});
