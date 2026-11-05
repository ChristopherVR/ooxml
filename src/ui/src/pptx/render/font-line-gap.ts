/**
 * The typographic line gap PowerPoint adds above the ascent of fonts that ask
 * for their typo metrics, which a browser does not expose.
 *
 * Measured on PowerPoint 16 exports of an `H` at 72pt, single spaced, in every
 * font in the Windows font folder (235 families). For a font without the
 * `USE_TYPO_METRICS` bit (`OS/2.fsSelection` bit 7) the baseline is the Windows
 * ascent scaled to the 1.2 em pitch, whatever its line gap (Calibri, Arial and
 * Consolas have large gaps and are within 1px of that). A font with the bit
 * uses `sTypoAscender` and `sTypoDescender`, and `sTypoLineGap` is added above
 * the ascent before the box is scaled to the pitch: Gabriola (gap 0.70 em)
 * puts its baseline at 0.977 em, where ascent and descent alone say 0.82 em.
 * Canvas `fontBoundingBox*` metrics already follow the same typo ascent and
 * descent but leave the gap out, so the gaps of the flagged fonts that have one
 * are listed here (em of `sTypoLineGap`, read from the font files).
 *
 * @module render/font-line-gap
 */

/** `sTypoLineGap` (em) of the installed and Office fonts with `USE_TYPO_METRICS` and a gap. */
const TYPO_LINE_GAP: ReadonlyMap<string, number> = new Map([
	['bahnschrift', 0.2],
	['cambria math', 0.172],
	['dejavu math tex gyre', 0.2],
	['gabriola', 0.7],
	['nirmala text', 0.862],
	['nirmala text semilight', 0.862],
	['sitka text', 0.25],
	['sitka heading', 0.25],
	['sitka subheading', 0.25],
	['sitka small', 0.25],
	['sitka banner', 0.25],
	['sitka display', 0.25],
]);

/** The line gap (em) PowerPoint adds above the ascent of `family`, or 0. */
export function typoLineGapEm(family: string | undefined): number {
	return (family && TYPO_LINE_GAP.get(family.trim().toLowerCase())) || 0;
}
