/**
 * Per-paragraph line-height and vertical margins (`a:lnSpc` / `a:spcBef` /
 * `a:spcAft`), resolved against the text body's own spacing.
 *
 * Split out of `text-paragraphs` to keep the builder focused; every binding
 * reaches it through the same barrel, so the import path is unchanged.
 */
import type { TextStyle } from 'ooxml-core/pptx';

import { proportionalLineHeight } from './text-line-height';

/** Points to CSS px. */
const PT_TO_PX = 96 / 72;

/** Resolved per-paragraph line-height + space-before/after. */
export interface ParagraphSpacing {
	/** Unitless multiplier (`a:spcPct`) or a `"<n>pt"` / `"<n>px"` string (`a:spcPts`). */
	lineHeight?: number | string;
	/**
	 * `margin-top` in px: this paragraph's own `a:spcBef`. Suppressed on the
	 * first paragraph unless `spaceFirstLast` (see
	 * {@link ParagraphSpacingInput.spaceFirstLast}).
	 */
	spaceBeforePx?: number;
	/**
	 * `margin-bottom` in px: this paragraph's own `a:spcAft`. Suppressed on the
	 * last paragraph unless `spaceFirstLast`. The text body lays its paragraphs
	 * out in a flex column, so the two margins never collapse: the gap between
	 * two paragraphs is the first one's `spcAft` PLUS the next one's `spcBef`,
	 * as PowerPoint adds them.
	 */
	spaceAfterPx?: number;
	/**
	 * `line-height` for every run of the paragraph, set only with exact line
	 * spacing (`a:spcPts`): `0`, so a run larger than the paragraph's strut
	 * cannot grow the line box. See {@link FIXED_SPACING_RUN_LINE_HEIGHT}.
	 */
	runLineHeight?: number;
}

/**
 * The `line-height` a run gets inside a paragraph with exact line spacing.
 *
 * PowerPoint keeps an exact pitch whatever sizes the runs on a line have; a
 * larger run may overflow its line but never moves the next one. CSS does not:
 * a run inherits the paragraph's absolute `line-height` and places it around
 * its OWN font metrics, so a 9pt run among 7pt text sits higher and the line
 * box grows to hold both (issue #35: 11pt spacing measured 11.75pt in
 * Chromium). With a `line-height` of `0` a run's inline box shrinks to a point
 * inside the paragraph's strut, which alone sizes the line. A run whose glyph
 * midline rises above the strut's top (roughly one more than the line pitch
 * plus the strut size, in points) can still push the line down, but by far
 * less than before.
 */
export const FIXED_SPACING_RUN_LINE_HEIGHT = 0;

/** Input for {@link resolveParagraphSpacing}. */
export interface ParagraphSpacingInput {
	/** This paragraph's own `a:pPr` geometry (from its first segment). */
	paraProps: TextStyle | undefined;
	/** The text body's style, used as the inheritance fallback. */
	bodyStyle?: TextStyle | undefined;
	/** OOXML outline level (zero when omitted). */
	paragraphLevel?: number | undefined;
	/** True for the first paragraph in the body. */
	isFirst?: boolean;
	/** True for the last paragraph in the body. */
	isLast?: boolean;
	/**
	 * `a:bodyPr/@spcFirstLastPara`. Gates two things, each scoped to ONE edge
	 * paragraph only: the very first paragraph's own `spcBef`, and the very
	 * last paragraph's own `spcAft`. Off (the default, including omitted)
	 * suppresses both; every other paragraph's own spacing is unaffected by
	 * this flag either way.
	 *
	 * Measured on PowerPoint 16 slide exports (`Slide.Export` PNG, Calibri 28pt,
	 * zero insets): a first paragraph's `spcBef` of 24pt moves its first line
	 * down by exactly 24pt only with the flag on, and a bottom-anchored body's
	 * last paragraph `spcAft` of 24pt lifts the block by exactly 24pt only with
	 * the flag on.
	 */
	spaceFirstLast?: boolean;
	/**
	 * `a:normAutofit/@lnSpcReduction`: the shrink-to-fit line-spacing reduction,
	 * in 0..1. Applied to a resolved proportional (`a:spcPct`) line-height the
	 * same way `computeAutoFitTextStyle` applies it at the block level - it has
	 * to be applied here too because a paragraph or body line-spacing value
	 * produces an explicit per-paragraph `line-height`, which always outranks
	 * the block's own CSS `line-height` in the cascade. Not applied to an exact
	 * (`a:spcPts`) line-height: PowerPoint's reduction only affects the
	 * proportional form.
	 */
	lineSpacingReduction?: number;
}

/**
 * Resolve a paragraph's line-height and vertical margin.
 *
 * OOXML puts line spacing (`a:lnSpc`), space-before (`a:spcBef`) and
 * space-after (`a:spcAft`) on the paragraph, so collapsing them into one
 * body-level padding gives every paragraph the same gap and loses the authored
 * rhythm. Values the paragraph does not set fall back to the text body's, and
 * an exact measure (`a:spcPts`) beats a proportional one (`a:spcPct`) taken
 * from the same level; a paragraph's own multiplier is never mixed with an
 * exact value inherited from the body.
 *
 * A paragraph's own `spcBef` is space ABOVE it and its own `spcAft` space
 * BELOW it, and PowerPoint adds the two across a paragraph boundary. Measured
 * on PowerPoint 16 slide exports (Calibri 28pt, single spacing, a 33.6pt
 * pitch): `spcAft` 6pt on the first paragraph and `spcBef` 12pt on the second
 * put the second line 51.75pt below the first (33.6 + 6 + 12), and a lone
 * `spcBef` of 30pt on a middle paragraph opens the gap ABOVE it (63.75pt)
 * while the gap below it stays one pitch. Percentage spacing (`spcPct`) is a
 * fraction of the paragraph's 1.2 x font-size line (core resolves it to px):
 * `spcAft` 100% plus the next paragraph's `spcBef` 50% at 28pt measured
 * 84pt between the lines (33.6 + 33.6 + 16.8). Autofit's `lnSpcReduction`
 * shrinks the line pitch only, never this spacing (with a 20% reduction a
 * 6pt `spcAft` and a 12pt `spcBef` still added 18pt to the gap).
 *
 * (`TextRange2.Paragraphs(n).BoundTop` reads cannot tell these apart: a
 * paragraph's bound box already includes its own `spcBef`, which once led
 * this function to fold `spcBef` below the paragraph instead.)
 *
 * `paragraphSpacingBefore` / `paragraphSpacingAfter` are already px from core.
 */
export function resolveParagraphSpacing(input: ParagraphSpacingInput): ParagraphSpacing {
	const {
		paraProps,
		bodyStyle,
		isFirst = false,
		isLast = false,
		spaceFirstLast = false,
		lineSpacingReduction,
	} = input;
	const out: ParagraphSpacing = {};
	const inherited = bodyStyle?.inheritedParagraphSpacing?.[input.paragraphLevel ?? 0];
	const fallback = inherited ? { ...inherited } : bodyStyle;
	// Element-wide edits still override the parsed baseline for all paragraphs.
	if (inherited && fallback && bodyStyle?.resolvedParagraphGeometry) {
		for (const key of ['paragraphSpacingBefore', 'paragraphSpacingAfter'] as const) {
			if (bodyStyle[key] !== bodyStyle.resolvedParagraphGeometry[key]) {
				fallback[key] = bodyStyle[key];
			}
		}
		if (
			bodyStyle.lineSpacing !== bodyStyle.resolvedParagraphGeometry.lineSpacing ||
			bodyStyle.lineSpacingExactPt !== bodyStyle.resolvedParagraphGeometry.lineSpacingExactPt
		) {
			fallback.lineSpacing = bodyStyle.lineSpacing;
			fallback.lineSpacingExactPt = bodyStyle.lineSpacingExactPt;
		}
	}

	const before = paraProps?.paragraphSpacingBefore ?? fallback?.paragraphSpacingBefore;
	const effectiveBefore =
		typeof before === 'number' && before > 0 && (!isFirst || spaceFirstLast) ? before : 0;

	const after = paraProps?.paragraphSpacingAfter ?? fallback?.paragraphSpacingAfter;
	const effectiveAfter =
		typeof after === 'number' && after > 0 && (!isLast || spaceFirstLast) ? after : 0;

	if (effectiveBefore > 0) {
		out.spaceBeforePx = effectiveBefore;
	}
	if (effectiveAfter > 0) {
		out.spaceAfterPx = effectiveAfter;
	}

	const hasOwnLineSpacing =
		paraProps?.lineSpacing !== undefined || paraProps?.lineSpacingExactPt !== undefined;
	const lineSource = hasOwnLineSpacing ? paraProps : fallback;
	const exactPt = lineSource?.lineSpacingExactPt;
	const multiplier = lineSource?.lineSpacing;
	if (typeof exactPt === 'number' && exactPt > 0) {
		out.lineHeight = `${exactPt * PT_TO_PX}px`;
		out.runLineHeight = FIXED_SPACING_RUN_LINE_HEIGHT;
	} else if (typeof multiplier === 'number' && multiplier > 0) {
		// `a:spcPct` stacks on PowerPoint's 1.2 single-spacing pitch; see
		// `proportionalLineHeight` for the COM measurement behind it.
		// `compatLnSpc` (`a:bodyPr` only, never per-paragraph) opts out of that
		// stacking.
		const proportional = proportionalLineHeight(multiplier, bodyStyle?.compatibleLineSpacing);
		out.lineHeight =
			typeof lineSpacingReduction === 'number' && lineSpacingReduction > 0
				? proportional * (1 - lineSpacingReduction)
				: proportional;
	}

	return out;
}
