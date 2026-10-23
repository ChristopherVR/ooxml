import { describe, it, expect, test, expectTypeOf } from 'vitest';

import {
	FILL_MODE_OPTIONS,
	PATTERN_PRESET_OPTIONS,
	GRADIENT_TYPE_OPTIONS,
	IMAGE_MODE_OPTIONS,
} from './fill-stroke-options';

// ---------------------------------------------------------------------------
// Helper: assert every item has non-empty value + label and no duplicate values
// ---------------------------------------------------------------------------

function assertOptionArray(options: ReadonlyArray<{ value: string; label: string }>, name: string) {
	test(`${name}: every item has a non-empty value`, () => {
		for (const opt of options) {
			expect(opt.value).toBeTruthy();
			expectTypeOf(opt.value).toBeString();
		}
	});

	test(`${name}: every item has a non-empty label`, () => {
		for (const opt of options) {
			expect(opt.label).toBeTruthy();
			expectTypeOf(opt.label).toBeString();
		}
	});

	test(`${name}: no duplicate values`, () => {
		const values = options.map((o) => o.value);
		expect(new Set(values).size).toBe(values.length);
	});
}

// ---------------------------------------------------------------------------
// FILL_MODE_OPTIONS
// ---------------------------------------------------------------------------

describe('fILL_MODE_OPTIONS', () => {
	it('has exactly 5 items', () => {
		expect(FILL_MODE_OPTIONS).toHaveLength(5);
	});

	it('contains solid, gradient, pattern, image, none', () => {
		const values = FILL_MODE_OPTIONS.map((o) => o.value);
		expect(values).toStrictEqual(['solid', 'gradient', 'pattern', 'image', 'none']);
	});

	assertOptionArray(FILL_MODE_OPTIONS, 'FILL_MODE_OPTIONS');
});

// ---------------------------------------------------------------------------
// PATTERN_PRESET_OPTIONS
// ---------------------------------------------------------------------------

// `PATTERN_PRESET_OPTIONS` is now the shared `{ value, labelKey }` catalogue
// (render/fill-pattern-label-keys.ts): no English `label`, so it does not fit
// `assertOptionArray` (used by every other list on this page). This pins the
// repoint's shape + wiring through the binding instead.
describe('pATTERN_PRESET_OPTIONS', () => {
	it('has 50 or more items', () => {
		expect(PATTERN_PRESET_OPTIONS.length).toBeGreaterThanOrEqual(50);
	});

	it('contains well-known pattern values', () => {
		const values = PATTERN_PRESET_OPTIONS.map((o) => o.value);
		expect(values).toContain('pct5');
		expect(values).toContain('horz');
		expect(values).toContain('vert');
		expect(values).toContain('cross');
		expect(values).toContain('diagCross');
		expect(values).toContain('zigZag');
	});

	it('every item has a non-empty value and labelKey, no duplicate values', () => {
		for (const opt of PATTERN_PRESET_OPTIONS) {
			expect(opt.value).toBeTruthy();
			expect(opt.labelKey).toBeTruthy();
			expectTypeOf(opt.labelKey).toBeString();
		}
		const values = PATTERN_PRESET_OPTIONS.map((o) => o.value);
		expect(new Set(values).size).toBe(values.length);
	});
});

// ---------------------------------------------------------------------------
// GRADIENT_TYPE_OPTIONS
// ---------------------------------------------------------------------------

describe('gRADIENT_TYPE_OPTIONS', () => {
	it('has exactly 2 items', () => {
		expect(GRADIENT_TYPE_OPTIONS).toHaveLength(2);
	});

	it('contains linear and radial', () => {
		const values = GRADIENT_TYPE_OPTIONS.map((o) => o.value);
		expect(values).toStrictEqual(['linear', 'radial']);
	});

	assertOptionArray(GRADIENT_TYPE_OPTIONS, 'GRADIENT_TYPE_OPTIONS');
});

// ---------------------------------------------------------------------------
// IMAGE_MODE_OPTIONS
// ---------------------------------------------------------------------------

describe('iMAGE_MODE_OPTIONS', () => {
	it('has exactly 2 items', () => {
		expect(IMAGE_MODE_OPTIONS).toHaveLength(2);
	});

	it('contains stretch and tile', () => {
		const values = IMAGE_MODE_OPTIONS.map((o) => o.value);
		expect(values).toStrictEqual(['stretch', 'tile']);
	});

	assertOptionArray(IMAGE_MODE_OPTIONS, 'IMAGE_MODE_OPTIONS');
});
