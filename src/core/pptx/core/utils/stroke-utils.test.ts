import { describe, it, expect } from 'vitest';

import {
	normalizeStrokeDashType,
	getCssBorderDashStyle,
	getSvgStrokeDasharray,
} from './stroke-utils';

// ---------------------------------------------------------------------------
// normalizeStrokeDashType
// ---------------------------------------------------------------------------

describe('normalizeStrokeDashType', () => {
	it('returns undefined for undefined input', () => {
		expect(normalizeStrokeDashType(undefined)).toBeUndefined();
	});

	it('returns undefined for empty string', () => {
		expect(normalizeStrokeDashType('')).toBeUndefined();
	});

	it('returns undefined for whitespace-only', () => {
		expect(normalizeStrokeDashType('   ')).toBeUndefined();
	});

	it('normalizes "solid" to "solid"', () => {
		expect(normalizeStrokeDashType('solid')).toBe('solid');
	});

	it('normalizes "dot" to "dot"', () => {
		expect(normalizeStrokeDashType('dot')).toBe('dot');
	});

	it('normalizes "dash" to "dash"', () => {
		expect(normalizeStrokeDashType('dash')).toBe('dash');
	});

	it('normalizes "lgDash" (case-insensitive) to "lgDash"', () => {
		expect(normalizeStrokeDashType('lgDash')).toBe('lgDash');
		expect(normalizeStrokeDashType('LGDASH')).toBe('lgDash');
	});

	it('normalizes "dashDot" to "dashDot"', () => {
		expect(normalizeStrokeDashType('dashDot')).toBe('dashDot');
		expect(normalizeStrokeDashType('DASHDOT')).toBe('dashDot');
	});

	it('normalizes "lgDashDot" to "lgDashDot"', () => {
		expect(normalizeStrokeDashType('lgDashDot')).toBe('lgDashDot');
	});

	it('normalizes "lgDashDotDot" to "lgDashDotDot"', () => {
		expect(normalizeStrokeDashType('lgDashDotDot')).toBe('lgDashDotDot');
	});

	it('normalizes system dash types (sysDot, sysDash, etc.)', () => {
		expect(normalizeStrokeDashType('sysDot')).toBe('sysDot');
		expect(normalizeStrokeDashType('sysDash')).toBe('sysDash');
		expect(normalizeStrokeDashType('sysDashDot')).toBe('sysDashDot');
		expect(normalizeStrokeDashType('sysDashDotDot')).toBe('sysDashDotDot');
	});

	it('normalizes "custom" to "custom"', () => {
		expect(normalizeStrokeDashType('custom')).toBe('custom');
	});

	it('returns undefined for unknown dash types', () => {
		expect(normalizeStrokeDashType('zigzag')).toBeUndefined();
		expect(normalizeStrokeDashType('unknown')).toBeUndefined();
	});
});

// ---------------------------------------------------------------------------
// getCssBorderDashStyle
// ---------------------------------------------------------------------------

describe('getCssBorderDashStyle', () => {
	it('returns "solid" for undefined', () => {
		expect(getCssBorderDashStyle(undefined)).toBe('solid');
	});

	it('returns "solid" for solid', () => {
		expect(getCssBorderDashStyle('solid')).toBe('solid');
	});

	it('returns "dotted" for dot', () => {
		expect(getCssBorderDashStyle('dot')).toBe('dotted');
	});

	it('returns "dotted" for sysDot', () => {
		expect(getCssBorderDashStyle('sysDot')).toBe('dotted');
	});

	it('returns "dashed" for dash', () => {
		expect(getCssBorderDashStyle('dash')).toBe('dashed');
	});

	it('returns "dashed" for lgDash', () => {
		expect(getCssBorderDashStyle('lgDash')).toBe('dashed');
	});

	it('returns "dashed" for dashDot', () => {
		expect(getCssBorderDashStyle('dashDot')).toBe('dashed');
	});

	it('returns "dashed" for lgDashDotDot', () => {
		expect(getCssBorderDashStyle('lgDashDotDot')).toBe('dashed');
	});
});

// ---------------------------------------------------------------------------
// getSvgStrokeDasharray
// ---------------------------------------------------------------------------

describe('getSvgStrokeDasharray', () => {
	it('returns undefined for undefined and solid', () => {
		expect(getSvgStrokeDasharray(undefined, 2)).toBeUndefined();
		expect(getSvgStrokeDasharray('solid', 2)).toBeUndefined();
	});

	// Lengths measured from PowerPoint exports, in pen widths.
	it.each([
		['dot', 1, '1 3'],
		['dash', 2, '8 6'],
		['lgDash', 1, '8 3'],
		['dashDot', 1, '4 3 1 3'],
		['lgDashDot', 1, '8 3 1 3'],
		['lgDashDotDot', 1, '8 3 1 3 1 3'],
		['sysDot', 6, '6 6'],
		['sysDash', 6, '18 6'],
		['sysDashDot', 6, '18 6 6 6'],
	] as const)('%s at %d px', (type, width, expected) => {
		expect(getSvgStrokeDasharray(type, width)).toBe(expected);
	});

	it('rounds the pen to whole pixels, at least one', () => {
		expect(getSvgStrokeDasharray('dot', 0)).toBe('1 3');
		expect(getSvgStrokeDasharray('dot', 2.67)).toBe('3 9');
	});

	it('builds custom dashes from a:ds units (100000 = one pen)', () => {
		const segments = [
			{ dash: 400000, space: 200000 },
			{ dash: 100000, space: 100000 },
		];
		expect(getSvgStrokeDasharray('custom', 2, segments)).toBe('8 4 2 2');
	});

	it('shortens round dashes by the pen width', () => {
		expect(getSvgStrokeDasharray('dot', 6, undefined, 'rnd')).toBe('0 24');
	});
});
