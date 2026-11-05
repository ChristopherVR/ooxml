/**
 * The edge profiles PowerPoint paints for blurred outer shadows and glows,
 * and the CSS lengths that reproduce them.
 *
 * Measured on PowerPoint 16 (`Slide.Export`, 1280 x 720, a 200 x 120 pt
 * rectangle on white; edge profile outside the right edge, fitted by
 * `A * erfc((x - c) / (s * sqrt 2)) / 2`, residual under 0.001):
 *
 * - `outerShdw blurRad = B`: a Gaussian-blurred copy of the silhouette,
 *   `c = 0`, `s = 0.326 B` (B = 10, 20, 30, 50 pt: s = 0.325, 0.327, 0.328,
 *   0.328 of B in px), reaching about `0.75 B` out. A CSS blur length `L` has
 *   deviation `L / 2`, so the length is `0.654 B`; halving `B` (the previous
 *   mapping, deviation `0.25 B`) left every shadow 24 % too narrow.
 * - `glow rad = R`: the silhouette grown by `c = 0.417 R` and then blurred
 *   with `s = R / 6` (R = 5..40 pt; c / R = 0.41-0.42, s / R = 0.164-0.167).
 *   One CSS `box-shadow: 0 0 (2 s) (c)` is exactly that profile.
 *
 * @module render/effect-profile
 */

/** CSS blur length per DrawingML `blurRad` unit: twice PowerPoint's deviation of 0.326. */
export const SHADOW_BLUR_TO_CSS = 0.654;

/** A glow's dilation, as a fraction of `a:glow@rad`. */
export const GLOW_GROW_FRACTION = 0.417;

/** A glow's Gaussian deviation, as a fraction of `a:glow@rad`. */
export const GLOW_SIGMA_FRACTION = 1 / 6;

/** The CSS blur length (`box-shadow` / `drop-shadow`) of a DrawingML shadow blur radius. */
export function shadowBlurToCss(blurRadius: number): number {
	return Math.max(0, blurRadius) * SHADOW_BLUR_TO_CSS;
}

/** A glow as CSS lengths: the dilation (`spread`) and the blur length (`2 s`). */
export function glowCssLengths(radius: number): { spread: number; blur: number } {
	const r = Math.max(0, radius);
	return { spread: r * GLOW_GROW_FRACTION, blur: r * GLOW_SIGMA_FRACTION * 2 };
}

/** Layers of the `drop-shadow()` chain that stands in for a glow where no `box-shadow` spread can. */
export const GLOW_FILTER_LAYERS = 3;

/** Each chained layer's blur length, as a fraction of `a:glow@rad`. */
export const GLOW_FILTER_BLUR_FRACTION = 0.38;

/**
 * The blur lengths of a glow painted as chained `drop-shadow()`s (text runs,
 * pictures, groups: `drop-shadow` cannot grow the silhouette, but each one
 * shadows the previous one's output, so a few equal layers widen the halo the
 * way the dilation does). Fitted against PowerPoint's own export of an 80 pt
 * bold run with a 20 pt glow: mean abs diff outside the glyphs 0.36 against
 * 0.77 for the earlier 0.33/0.66/1.0 stack.
 */
export function glowFilterLengths(radius: number): number[] {
	const blur = Math.max(0, radius) * GLOW_FILTER_BLUR_FRACTION;
	return Array.from({ length: GLOW_FILTER_LAYERS }, () => Math.round(blur * 100) / 100);
}

/**
 * A text run's shadow blur as a CSS `text-shadow` length per `blurRad` unit. A
 * glyph's shadow reads wider than a rectangle's: against PowerPoint's export of
 * an 80 pt bold run (blurRad 12 pt) the best length is 11.5-12 px for 16 px,
 * 0.72, not 0.654.
 */
export const TEXT_SHADOW_BLUR_TO_CSS = 0.72;
