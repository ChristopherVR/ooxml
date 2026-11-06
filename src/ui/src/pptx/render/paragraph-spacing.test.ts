import type { TextStyle } from 'ooxml-core/pptx';
import { describe, it, expect } from 'vitest';

import { resolveParagraphSpacing } from './paragraph-spacing';

const base = {
	isFirst: false,
	isLast: false,
	spaceFirstLast: true,
};

describe('resolveParagraphSpacing', () => {
	it('uses inherited level spacing instead of the first paragraph direct formatting (#74)', () => {
		const bodyStyle = {
			lineSpacingExactPt: 12.86,
			paragraphSpacingAfter: (3.8 * 96) / 72,
			inheritedParagraphSpacing: {
				0: { lineSpacing: 1, paragraphSpacingAfter: 0 },
				1: { lineSpacingExactPt: 18, paragraphSpacingAfter: 8 },
			},
		};
		expect(resolveParagraphSpacing({ paraProps: undefined, bodyStyle })).toEqual({
			lineHeight: 1.2,
		});
		expect(resolveParagraphSpacing({ paraProps: undefined, bodyStyle, paragraphLevel: 1 })).toEqual(
			{ lineHeight: '24px', runLineHeight: 0, spaceAfterPx: 8 },
		);
		expect(
			resolveParagraphSpacing({
				paraProps: { lineSpacingExactPt: 12.86, paragraphSpacingAfter: 4 },
				bodyStyle,
			}),
			// 12.86pt lays out at 13pt in PowerPoint (whole points, half up).
		).toEqual({ lineHeight: `${(13 * 96) / 72}px`, runLineHeight: 0, spaceAfterPx: 4 });
	});

	it('continues applying body-wide line spacing edits after loading (#74)', () => {
		expect(
			resolveParagraphSpacing({
				paraProps: undefined,
				bodyStyle: {
					lineSpacing: 2,
					resolvedParagraphGeometry: { lineSpacingExactPt: 12.86 },
					inheritedParagraphSpacing: { 0: { lineSpacing: 1 } },
				},
			}),
		).toEqual({ lineHeight: 2.4 });
	});
	it('returns nothing when neither the paragraph nor the body sets spacing', () => {
		expect(resolveParagraphSpacing({ ...base, paraProps: undefined })).toStrictEqual({});
		expect(
			resolveParagraphSpacing({ ...base, paraProps: undefined, bodyStyle: undefined }),
		).toStrictEqual({});
	});

	it('expresses exact spcPts line spacing in px', () => {
		const out = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacingExactPt: 18 } as TextStyle,
		});
		// 18pt at 96dpi. Equivalent to the "18pt" some bindings emitted before,
		// but one unit for all five.
		expect(out.lineHeight).toBe('24px');
	});

	// Issue #35: PowerPoint keeps an exact pitch whatever run sizes a line
	// holds, so the runs must not grow the line box around their own metrics.
	it('gives the runs of an exactly spaced paragraph a zero line-height', () => {
		const exact = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacingExactPt: 11 } as TextStyle,
		});
		expect(exact.runLineHeight).toBe(0);
		const proportional = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacing: 1 } as TextStyle,
		});
		expect(proportional.runLineHeight).toBeUndefined();
	});

	it('maps a paragraphs own before/after spacing to its top and bottom margins', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(result.spaceBeforePx).toBe(12);
		expect(result.spaceAfterPx).toBe(8);
	});

	it('falls back to body-level spacing when the paragraph has none', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: undefined,
			bodyStyle: { paragraphSpacingBefore: 5, paragraphSpacingAfter: 7 } as TextStyle,
		});
		expect(result.spaceBeforePx).toBe(5);
		expect(result.spaceAfterPx).toBe(7);
	});

	it('prefers paragraph spacing over the body fallback', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { paragraphSpacingBefore: 20 } as TextStyle,
			bodyStyle: { paragraphSpacingBefore: 5 } as TextStyle,
		});
		expect(result.spaceBeforePx).toBe(20);
		expect(result.spaceAfterPx).toBeUndefined();
	});

	it('maps a line-spacing multiplier to a unitless line-height', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacing: 1.5 } as TextStyle,
			bodyStyle: undefined,
		});
		// 1.5 stacks on the 1.2 single-spacing base (spcPct multiplies the pitch).
		expect(result.lineHeight).toBeCloseTo(1.8, 10);
	});

	it('maps exact-pt line spacing to a px line-height', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacingExactPt: 18 } as TextStyle,
			bodyStyle: undefined,
		});
		// 18pt * 96/72 = 24px
		expect(result.lineHeight).toBe('24px');
	});

	it('uses the paragraph line-spacing unit without mixing body values', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacing: 2 } as TextStyle,
			bodyStyle: { lineSpacingExactPt: 40 } as TextStyle,
		});
		expect(result.lineHeight).toBeCloseTo(2.4, 10);
	});

	it('suppresses the first paragraphs own before-spacing when spaceFirstLast is false', () => {
		const result = resolveParagraphSpacing({
			...base,
			isFirst: true,
			spaceFirstLast: false,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(result.spaceBeforePx).toBeUndefined();
		expect(result.spaceAfterPx).toBe(8);
	});

	it('suppresses the last paragraphs own after-spacing when spaceFirstLast is false', () => {
		const result = resolveParagraphSpacing({
			...base,
			isLast: true,
			spaceFirstLast: false,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(result.spaceBeforePx).toBe(12);
		expect(result.spaceAfterPx).toBeUndefined();
	});

	it('suppresses first/last edge spacing by default when spaceFirstLast is omitted', () => {
		// ECMA-376's default for `a:bodyPr/@spcFirstLastPara` is false, confirmed
		// on PowerPoint slide exports: an omitted attribute behaves as "0".
		const firstResult = resolveParagraphSpacing({
			isFirst: true,
			isLast: false,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(firstResult).toStrictEqual({ spaceAfterPx: 8 });

		const lastResult = resolveParagraphSpacing({
			isFirst: false,
			isLast: true,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(lastResult).toStrictEqual({ spaceBeforePx: 12 });
	});

	it('applies first/last edge spacing when spaceFirstLast is true', () => {
		const result = resolveParagraphSpacing({
			...base,
			isFirst: true,
			isLast: true,
			spaceFirstLast: true,
			paraProps: { paragraphSpacingBefore: 12, paragraphSpacingAfter: 8 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(result.spaceBeforePx).toBe(12);
		expect(result.spaceAfterPx).toBe(8);
	});

	it('leaves an interior paragraph fully unaffected by spaceFirstLast either way', () => {
		const interior = {
			isFirst: false,
			isLast: false,
			paraProps: { paragraphSpacingBefore: 40, paragraphSpacingAfter: 30 } as TextStyle,
			bodyStyle: undefined,
		};
		const withFlagOff = resolveParagraphSpacing({ ...interior, spaceFirstLast: false });
		const withFlagOn = resolveParagraphSpacing({ ...interior, spaceFirstLast: true });
		expect(withFlagOff).toStrictEqual({ spaceBeforePx: 40, spaceAfterPx: 30 });
		expect(withFlagOn).toStrictEqual({ spaceBeforePx: 40, spaceAfterPx: 30 });
	});

	it('gates only the first paragraph before-spacing and last paragraph after-spacing independently', () => {
		// p1 in a body of >= 3 paragraphs: own before gated, own after always on
		// (it is not the last paragraph).
		const p1 = {
			isFirst: true,
			isLast: false,
			paraProps: { paragraphSpacingBefore: 40, paragraphSpacingAfter: 30 } as TextStyle,
			bodyStyle: undefined,
		};
		expect(resolveParagraphSpacing({ ...p1, spaceFirstLast: false })).toStrictEqual({
			spaceAfterPx: 30,
		});
		expect(resolveParagraphSpacing({ ...p1, spaceFirstLast: true })).toStrictEqual({
			spaceBeforePx: 40,
			spaceAfterPx: 30,
		});

		// pN in a body of >= 3 paragraphs: own after gated, own before always on
		// (it is not the first paragraph).
		const pN = {
			isFirst: false,
			isLast: true,
			paraProps: { paragraphSpacingBefore: 40, paragraphSpacingAfter: 30 } as TextStyle,
			bodyStyle: undefined,
		};
		expect(resolveParagraphSpacing({ ...pN, spaceFirstLast: false })).toStrictEqual({
			spaceBeforePx: 40,
		});
		expect(resolveParagraphSpacing({ ...pN, spaceFirstLast: true })).toStrictEqual({
			spaceBeforePx: 40,
			spaceAfterPx: 30,
		});
	});

	it('reduces a proportional line-height by the autofit lnSpcReduction fraction', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacing: 1.5 } as TextStyle,
			bodyStyle: undefined,
			lineSpacingReduction: 0.2,
		});
		// 1.5 * 1.2 pitch = 1.8, reduced by 20% -> 1.44.
		expect(result.lineHeight).toBeCloseTo(1.44, 10);
	});

	it('does not apply lnSpcReduction to an exact-pt line-height', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacingExactPt: 18 } as TextStyle,
			bodyStyle: undefined,
			lineSpacingReduction: 0.2,
		});
		expect(result.lineHeight).toBe('24px');
	});

	it('ignores a zero or absent lnSpcReduction', () => {
		const noReduction = resolveParagraphSpacing({
			...base,
			paraProps: { lineSpacing: 1.5 } as TextStyle,
			bodyStyle: undefined,
			lineSpacingReduction: 0,
		});
		expect(noReduction.lineHeight).toBeCloseTo(1.8, 10);
	});

	it('ignores zero / negative spacing values', () => {
		const result = resolveParagraphSpacing({
			...base,
			paraProps: { paragraphSpacingBefore: 0, paragraphSpacingAfter: -4 } as TextStyle,
			bodyStyle: undefined,
		});
		expect(result.spaceBeforePx).toBeUndefined();
		expect(result.spaceAfterPx).toBeUndefined();
	});
});

// Measured on PowerPoint 16 slide exports (`Slide.Export` PNG at 1280x720,
// Calibri 28pt, zero insets, single spacing: a 33.6pt pitch). The line tops
// a measured gap puts between consecutive paragraphs is the pitch plus the
// first one's spcAft PLUS the next one's spcBef. Units are used as px.
describe('resolveParagraphSpacing - measured paragraph gaps', () => {
	const PITCH = 33.6;
	function gaps(paras: TextStyle[], spaceFirstLast = false): number[] {
		const resolved = paras.map((paraProps, index) =>
			resolveParagraphSpacing({
				paraProps,
				isFirst: index === 0,
				isLast: index === paras.length - 1,
				spaceFirstLast,
			}),
		);
		return resolved
			.slice(1)
			.map(
				(next, index) => PITCH + (resolved[index]!.spaceAfterPx ?? 0) + (next.spaceBeforePx ?? 0),
			);
	}

	it('adds the space after one paragraph to the space before the next (51.75pt measured)', () => {
		const result = gaps([
			{ paragraphSpacingAfter: 6 },
			{ paragraphSpacingBefore: 12 },
			{},
		] as TextStyle[]);
		expect(result[0]).toBeCloseTo(51.6, 5);
		expect(result[1]).toBeCloseTo(PITCH, 5);
	});

	it('opens a middle paragraphs space before above it, not below (63.75pt, then one pitch)', () => {
		const result = gaps([{}, { paragraphSpacingBefore: 30 }, {}] as TextStyle[]);
		expect(result).toEqual([PITCH + 30, PITCH]);
	});

	it('adds before and after on every boundary when all paragraphs set both (51.75pt twice)', () => {
		const both = { paragraphSpacingBefore: 12, paragraphSpacingAfter: 6 } as TextStyle;
		expect(gaps([both, both, both])).toEqual([PITCH + 18, PITCH + 18]);
	});

	it('moves the first line down by its own space before only with spcFirstLastPara', () => {
		const first = { paragraphSpacingBefore: 24 } as TextStyle;
		const off = resolveParagraphSpacing({ paraProps: first, isFirst: true, isLast: false });
		const on = resolveParagraphSpacing({
			paraProps: first,
			isFirst: true,
			isLast: false,
			spaceFirstLast: true,
		});
		expect(off.spaceBeforePx).toBeUndefined();
		expect(on.spaceBeforePx).toBe(24);
	});
});
