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
