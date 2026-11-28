/**
 * Vertical metrics (em) of legacy East Asian Office fonts, which PowerPoint
 * lays a line out on even when the font is missing from the machine.
 *
 * Measured on PowerPoint 16: MS PMincho (not installed), MS PGothic and
 * SimSun all put the baseline at the same place, which is the one their own
 * Windows metrics give (`winAscent` 220 and `winDescent` 36 of 256 units, read
 * with fontTools for MS Gothic, MS PGothic, MS UI Gothic, SimSun and NSimSun;
 * the Mincho faces share the grid). A browser without the font measures its
 * stand-in instead, so a renderer needs these to place the line.
 *
 * @module text/font-metrics/vertical-metrics
 */

import { OFFICE_FONT_METRICS } from './office-font-metrics.generated';

/** A font's ascent and descent as fractions of the em. */
export interface VerticalMetrics {
	ascent: number;
	descent: number;
	/** The line gap PowerPoint adds above the ascent (em), when the font has one. */
	lineGap?: number;
}

const EAST_ASIAN_OFFICE_METRICS: VerticalMetrics = { ascent: 220 / 256, descent: 36 / 256 };

const NAMED_METRICS: ReadonlyMap<string, VerticalMetrics> = new Map(
	['ms mincho', 'ms pmincho', 'ms gothic', 'ms pgothic', 'ms ui gothic', 'simsun', 'nsimsun'].map(
		(name) => [name, EAST_ASIAN_OFFICE_METRICS],
	),
);

/**
 * The vertical metrics of `family` when it is a legacy East Asian Office font
 * PowerPoint knows by name, else `undefined`. Matching ignores case.
 */
export function namedFontVerticalMetrics(family: string | undefined): VerticalMetrics | undefined {
	return family ? NAMED_METRICS.get(family.trim().toLowerCase()) : undefined;
}

/**
 * The vertical metrics PowerPoint lays `family` out on, when it is one of the
 * fonts that ship with Windows and Office (see `office-font-metrics.generated`),
 * else `undefined`. Matching ignores case. Unlike {@link namedFontVerticalMetrics}
 * this is the table for every such font, installed or not: a renderer uses it
 * for the line gap of the ones whose browser box leaves it out, and for the
 * ascent and descent of the ones the browser does not have.
 */
export function officeFontVerticalMetrics(family: string | undefined): VerticalMetrics | undefined {
	const known = family ? OFFICE_FONT_METRICS.get(family.trim().toLowerCase()) : undefined;
	if (!known) {
		return undefined;
	}
	const [ascent, descent, lineGap] = known;
	return lineGap > 0 ? { ascent, descent, lineGap } : { ascent, descent };
}

/** `sTypoLineGap` (em) of fonts with `USE_TYPO_METRICS` and a gap that the Windows table lacks. */
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

/**
 * The typographic line gap (em of `sTypoLineGap`) of `family` when it sets `USE_TYPO_METRICS`
 * (`OS/2.fsSelection` bit 7) and has a gap, else 0. A browser exposes neither the bit nor the gap:
 * canvas `fontBoundingBox*` follows the typo ascent and descent and leaves the gap out. It is the
 * only case where the font's line gap is the typographic one; any other gap a browser reports is
 * the `hhea` leading the Windows metrics leave out, which Word places above the ascent (#382).
 * Covers the Windows and Office table (`officeFontVerticalMetrics`) and a few flagged fonts the
 * Windows font folder does not hold. Matching ignores case.
 */
export function typoLineGapEm(family: string | undefined): number {
	const own = family ? TYPO_LINE_GAP.get(family.trim().toLowerCase()) : undefined;
	return own ?? officeFontVerticalMetrics(family)?.lineGap ?? 0;
}
