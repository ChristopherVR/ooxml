import { describe, expect, it } from 'vitest';

import { getEffectKeyframes } from './animation-keyframes';
import {
	PIXELATE_FILTER_DEFS_MARKUP,
	PIXELATE_FILTER_VALUES,
	PIXELATE_IN_KEYFRAMES,
	PIXELATE_LEVELS,
	PIXELATE_OUT_KEYFRAMES,
} from './animation-pixelate-filter';

/** Pull `{ pct: number; opacity: number; filter: string }` stops out of a `@keyframes` block. */
function parseStops(css: string): Array<{ pct: number; opacity: number; filter: string }> {
	const stopRegex = /([\d.]+)%\s*\{\s*opacity:\s*([\d.]+);\s*filter:\s*([^;]+);\s*\}/g;
	const stops: Array<{ pct: number; opacity: number; filter: string }> = [];
	let match: RegExpExecArray | null;
	while ((match = stopRegex.exec(css)) !== null) {
		stops.push({ pct: Number(match[1]), opacity: Number(match[2]), filter: match[3] });
	}
	return stops;
}

describe('pixelate levels', () => {
	it('is a monotonically decreasing (coarse -> fine) fraction sequence', () => {
		expect(PIXELATE_LEVELS.length).toBeGreaterThanOrEqual(5);
		for (let i = 1; i < PIXELATE_LEVELS.length; i++) {
			expect(PIXELATE_LEVELS[i]).toBeLessThan(PIXELATE_LEVELS[i - 1]);
		}
		for (const level of PIXELATE_LEVELS) {
			expect(level).toBeGreaterThan(0);
			expect(level).toBeLessThanOrEqual(1);
		}
	});
});

describe('pixelate filter values', () => {
	it('emits one same-document filter reference per level', () => {
		expect(PIXELATE_FILTER_VALUES).toHaveLength(PIXELATE_LEVELS.length);
		PIXELATE_FILTER_VALUES.forEach((value, index) => {
			expect(value).toBe(`url(#pptx-pixelate-${index})`);
		});
	});

	it('defines every referenced filter with the mosaic primitives', () => {
		for (const primitive of [
			'primitiveUnits="objectBoundingBox"',
			'feFlood',
			'feTile',
			'SourceGraphic',
		]) {
			expect(PIXELATE_FILTER_DEFS_MARKUP).toContain(primitive);
		}
		PIXELATE_FILTER_VALUES.forEach((_value, index) => {
			expect(PIXELATE_FILTER_DEFS_MARKUP).toContain(`<filter id="pptx-pixelate-${index}"`);
		});
	});

	it('has no two levels sharing the same filter id', () => {
		const ids = PIXELATE_FILTER_VALUES.map((v) => /#(pptx-pixelate-\d+)\)$/.exec(v)?.[1]);
		expect(new Set(ids).size).toBe(ids.length);
	});
});

describe('pixelateIn keyframes (reveal: coarse -> fine)', () => {
	const stops = parseStops(PIXELATE_IN_KEYFRAMES);

	it('names the pptx-pixelateIn keyframes block', () => {
		expect(PIXELATE_IN_KEYFRAMES.startsWith('@keyframes pptx-pixelateIn {')).toBeTruthy();
	});

	it('has one stop per mosaic level plus the final fully-resolved stop', () => {
		expect(stops).toHaveLength(PIXELATE_LEVELS.length + 1);
	});

	it('starts hidden with the coarsest filter and ends visible with no filter', () => {
		expect(stops[0].pct).toBe(0);
		expect(stops[0].opacity).toBe(0);
		expect(stops[0].filter).toBe(PIXELATE_FILTER_VALUES[0]);

		const last = stops[stops.length - 1];
		expect(last.pct).toBe(100);
		expect(last.opacity).toBe(1);
		expect(last.filter).toBe('none');
	});

	it('steps through every level in coarse -> fine order', () => {
		for (let i = 0; i < PIXELATE_FILTER_VALUES.length; i++) {
			expect(stops[i].filter).toBe(PIXELATE_FILTER_VALUES[i]);
		}
	});

	it('ramps opacity and percentage monotonically', () => {
		for (let i = 1; i < stops.length; i++) {
			expect(stops[i].pct).toBeGreaterThan(stops[i - 1].pct);
			expect(stops[i].opacity).toBeGreaterThanOrEqual(stops[i - 1].opacity);
		}
	});
});

describe('pixelateOut keyframes (conceal: fine -> coarse)', () => {
	const stops = parseStops(PIXELATE_OUT_KEYFRAMES);

	it('names the pptx-pixelateOut keyframes block', () => {
		expect(PIXELATE_OUT_KEYFRAMES.startsWith('@keyframes pptx-pixelateOut {')).toBeTruthy();
	});

	it('starts visible with no filter and ends hidden with the coarsest filter', () => {
		expect(stops[0].pct).toBe(0);
		expect(stops[0].opacity).toBe(1);
		expect(stops[0].filter).toBe('none');

		const last = stops[stops.length - 1];
		expect(last.pct).toBe(100);
		expect(last.opacity).toBe(0);
		expect(last.filter).toBe(PIXELATE_FILTER_VALUES[0]);
	});

	it('steps through every level in fine -> coarse order', () => {
		const reversed = [...PIXELATE_FILTER_VALUES].reverse();
		for (let i = 0; i < reversed.length; i++) {
			expect(stops[i + 1].filter).toBe(reversed[i]);
		}
	});

	it('ramps opacity down and percentage up monotonically', () => {
		for (let i = 1; i < stops.length; i++) {
			expect(stops[i].pct).toBeGreaterThan(stops[i - 1].pct);
			expect(stops[i].opacity).toBeLessThanOrEqual(stops[i - 1].opacity);
		}
	});
});

describe('getEffectKeyframes integration', () => {
	it('resolves pixelateIn/pixelateOut through the composed effect lookup', () => {
		expect(getEffectKeyframes('pixelateIn')).toBe(PIXELATE_IN_KEYFRAMES);
		expect(getEffectKeyframes('pixelateOut')).toBe(PIXELATE_OUT_KEYFRAMES);
	});
});
