import { describe, expect, it } from 'vitest';

import {
	exactPitchPoints,
	lineHeightToPx,
	proportionalLineHeight,
	resolveLineHeight,
} from './text-line-height';

describe('resolveLineHeight', () => {
	it('returns an exact pt string when lineSpacingExactPt is set', () => {
		expect(resolveLineHeight({ lineSpacingExactPt: 18 }, false)).toBe('18pt');
	});

	it('ignores a non-positive exact pt and uses the multiplier', () => {
		// 1.5 stacks on the 1.2 single-spacing base (spcPct multiplies the pitch).
		expect(resolveLineHeight({ lineSpacingExactPt: 0, lineSpacing: 1.5 }, false)).toBeCloseTo(
			1.8,
			10,
		);
	});

	it('uses the proportional multiplier when set', () => {
		// 200% spacing lays out at 2.4x the font size in PowerPoint (COM-measured
		// on the issue #132 deck), not 2.0x: spcPct stacks on the 1.2 base.
		expect(resolveLineHeight({ lineSpacing: 2 }, false)).toBeCloseTo(2.4, 10);
	});

	it('defaults to PowerPoint single spacing (1.2x), italic or not', () => {
		// Measured against PowerPoint (COM TextRange2.BoundHeight, issue #131
		// deck): single-spaced lines are exactly 1.2x the font point size.
		expect(resolveLineHeight(undefined, false)).toBe(1.2);
		expect(resolveLineHeight(undefined, true)).toBe(1.2);
	});

	it('compatLnSpc stacks on the 1.2 pitch like every other body (PowerPoint 16)', () => {
		expect(resolveLineHeight({ lineSpacing: 1.5, compatibleLineSpacing: true }, false)).toBeCloseTo(
			1.8,
			10,
		);
		expect(resolveLineHeight({ compatibleLineSpacing: true }, false)).toBe(1.2);
	});

	it('exact pt spacing is unaffected by compatLnSpc', () => {
		expect(resolveLineHeight({ lineSpacingExactPt: 18, compatibleLineSpacing: true }, false)).toBe(
			'18pt',
		);
	});
});

describe('proportionalLineHeight', () => {
	it('stacks the multiplier on the 1.2 single-spacing pitch by default', () => {
		expect(proportionalLineHeight(1.5)).toBeCloseTo(1.8, 10);
		expect(proportionalLineHeight(undefined)).toBe(1.2);
	});

	it('compatLineSpacing does not change the pitch', () => {
		expect(proportionalLineHeight(1.5, true)).toBeCloseTo(1.8, 10);
		expect(proportionalLineHeight(undefined, true)).toBe(1.2);
		expect(proportionalLineHeight(0.9, true)).toBeCloseTo(1.08, 10);
	});
});

describe('lineHeightToPx', () => {
	it('multiplies a unitless multiplier by the font size', () => {
		expect(lineHeightToPx(20, 1.2)).toBe(24);
	});

	it('converts a pt string to px', () => {
		expect(lineHeightToPx(20, '18pt')).toBeCloseTo(24, 5);
	});

	it('reads a px string verbatim', () => {
		expect(lineHeightToPx(20, '30px')).toBe(30);
	});

	it('falls back to the single-spacing pitch for an unrecognised/absent value', () => {
		expect(lineHeightToPx(20, undefined)).toBeCloseTo(24, 5);
		expect(lineHeightToPx(20, 'normal')).toBeCloseTo(24, 5);
	});
});

describe('exact line pitch', () => {
	// PowerPoint 16, 8-line paragraphs, spcPts 10.2 to 15.5 in 0.1 and 0.5pt
	// steps: the fitted pitch was 13.32, 14.68, 16.0, 17.32, 18.68, 20.0, 21.32px.
	it('lays a fractional pitch out in whole points, half up', () => {
		expect(exactPitchPoints(10.2)).toBe(10);
		expect(exactPitchPoints(10.4)).toBe(10);
		expect(exactPitchPoints(10.5)).toBe(11);
		expect(exactPitchPoints(13.4)).toBe(13);
		expect(exactPitchPoints(13.5)).toBe(14);
		expect(exactPitchPoints(14.3)).toBe(14);
		expect(exactPitchPoints(14.5)).toBe(15);
	});

	it('leaves a pitch under half a point alone rather than collapsing it to zero', () => {
		expect(exactPitchPoints(0.3)).toBe(0.3);
	});

	it('resolves a 14.3pt pitch as 14pt', () => {
		expect(resolveLineHeight({ lineSpacingExactPt: 14.3 }, false)).toBe('14pt');
	});
});
