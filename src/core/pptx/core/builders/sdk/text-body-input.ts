/**
 * Text body options of a new text box or shape (`a:bodyPr`): insets, autofit
 * and wrapping, and the character spacing of its runs.
 *
 * The pptx save path writes `a:bodyPr` from the `TextStyle` body fields, so
 * the options map onto those fields; the shared DrawingML body writer is not
 * used by the pptx writer yet.
 *
 * @module sdk/text-body-input
 */

import type { TextStyle } from '../../types/text';
import type { TextBodyInput } from './types';

/** EMU per CSS pixel at 96 dpi. */
const EMU_PER_PX = 9525;
/** `ST_Coordinate32` bounds an inset in EMU. */
const MAX_INSET_EMU = 51206400;
/** `ST_TextPoint` bounds `a:rPr/@spc` in hundredths of a point. */
const MAX_SPC = 400000;

type InsetSide = 'left' | 'top' | 'right' | 'bottom';

const INSET_FIELDS = {
	left: 'bodyInsetLeft',
	top: 'bodyInsetTop',
	right: 'bodyInsetRight',
	bottom: 'bodyInsetBottom',
} as const satisfies Record<InsetSide, keyof TextStyle>;

const AUTO_FIT = {
	// The model's `autoFitMode` names read the other way round: `'normal'` is
	// `a:normAutofit` (PowerPoint shrinks the text on overflow) and `'shrink'`
	// is `a:spAutoFit` (the shape grows to fit the text).
	shrinkText: { autoFit: true, autoFitMode: 'normal' },
	resizeShape: { autoFit: true, autoFitMode: 'shrink' },
	none: { autoFit: false, autoFitMode: 'none' },
} as const satisfies Record<NonNullable<TextBodyInput['autoFit']>, Partial<TextStyle>>;

/**
 * The `TextStyle` body fields for {@link TextBodyInput}. Only the options that
 * are given are set, so an element without them saves exactly as before.
 *
 * @throws RangeError for an inset that is not a finite number within
 *   `ST_Coordinate32`, or an `autoFit` value that is not one of the three names.
 */
export function mapTextBodyInput(input: TextBodyInput | undefined): Partial<TextStyle> {
	const style: Partial<TextStyle> = {};
	if (!input) {
		return style;
	}
	const { insets, autoFit, wrap } = input;
	for (const side of Object.keys(INSET_FIELDS) as InsetSide[]) {
		const value = insets?.[side];
		if (value === undefined) {
			continue;
		}
		if (!Number.isFinite(value) || Math.abs(value * EMU_PER_PX) > MAX_INSET_EMU) {
			throw new RangeError(`insets.${side} must be a finite px length, got ${value}`);
		}
		style[INSET_FIELDS[side]] = value;
	}
	if (autoFit !== undefined) {
		const mapped: Partial<TextStyle> | undefined = AUTO_FIT[autoFit];
		if (!mapped) {
			throw new RangeError(
				`autoFit must be 'none', 'shrinkText' or 'resizeShape', got ${String(autoFit)}`,
			);
		}
		Object.assign(style, mapped);
	}
	if (wrap !== undefined) {
		style.textWrap = wrap ? 'square' : 'none';
	}
	return style;
}

/**
 * `a:rPr/@spc` (hundredths of a point) for a character spacing in px.
 *
 * @throws RangeError for a value that is not finite or lies outside `ST_TextPoint`.
 */
export function characterSpacingToSpc(px: number): number {
	const spc = Math.round(px * 75);
	if (!Number.isFinite(px) || Math.abs(spc) > MAX_SPC) {
		throw new RangeError(`characterSpacing must be a finite px length, got ${px}`);
	}
	return spc;
}
