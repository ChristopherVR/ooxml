/**
 * Pure connector line-style helpers shared across bindings.
 *
 * Compound (double/triple) parallel-stroke geometry and connector-family
 * classification. No framework imports. The actual SVG `<marker>` / `<path>`
 * emission stays in each binding's view layer.
 */

/** OOXML compound line token type (`a:ln/@cmpd`). */
export type CompoundLineType = 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';

/**
 * The bands a compound line is cut into across its width, first (left of the
 * direction of travel, so above a line drawn left to right) to last: strands
 * and the gaps between them, as fractions of the whole width.
 *
 * Measured on PowerPoint 16 (`Slide.Export`, 1280 x 720) with 12 pt (16 px)
 * lines: `dbl` paints 5.3 / 5.3 / 5.3 px (thirds), `thickThin` 9.6 / 3.2 / 3.2
 * (3 : 1 : 1), `thinThick` 3.2 / 3.2 / 9.6 and `tri` 2.7 / 2.7 / 5.3 / 2.7 /
 * 2.7 (1 : 1 : 2 : 1 : 1), always filling exactly the line's width.
 */
const COMPOUND_BANDS: Readonly<Record<string, readonly number[]>> = {
	dbl: [1 / 3, 1 / 3, 1 / 3],
	thickThin: [0.6, 0.2, 0.2],
	thinThick: [0.2, 0.2, 0.6],
	tri: [1 / 6, 1 / 6, 1 / 3, 1 / 6, 1 / 6],
};

/** Each strand's centre offset and width, as fractions of the line width. */
function compoundStrands(compoundLine: string | undefined): Array<[number, number]> | undefined {
	const bands = compoundLine ? COMPOUND_BANDS[compoundLine] : undefined;
	if (!bands) {
		return undefined;
	}
	const strands: Array<[number, number]> = [];
	let edge = -0.5;
	bands.forEach((band, index) => {
		if (index % 2 === 0) {
			strands.push([edge + band / 2, band]);
		}
		edge += band;
	});
	return strands;
}

/**
 * Compute perpendicular offsets for compound (double/triple) line styles.
 * Returns an array of offset distances from the centre line in px, negative
 * to the left of the direction of travel. A single line style returns `[0]`.
 *
 * @param compoundLine - OOXML `a:ln/@cmpd` value (e.g. `"dbl"`, `"tri"`).
 * @param strokeWidth  - The resolved stroke width in pixels.
 */
export function getCompoundLineOffsets(
	compoundLine: string | undefined,
	strokeWidth: number,
): number[] {
	const base = Math.max(strokeWidth, 1);
	// Rounded so a centred strand is exactly 0 and thirds come out whole.
	const exact = (value: number) => Math.round(value * 1e6) / 1e6 + 0;
	return compoundStrands(compoundLine)?.map(([offset]) => exact(offset * base)) ?? [0];
}

/**
 * Compute individual stroke widths for each parallel path in a compound line.
 * The array length matches the one returned by {@link getCompoundLineOffsets}.
 *
 * @param compoundLine - OOXML `a:ln/@cmpd` value.
 * @param strokeWidth  - The resolved stroke width in pixels.
 */
export function getCompoundLineWidths(
	compoundLine: string | undefined,
	strokeWidth: number,
): number[] {
	const base = Math.max(strokeWidth, 1);
	return compoundStrands(compoundLine)?.map(([, width]) => width * base) ?? [base];
}

/**
 * Map an OOXML line-cap token (`a:ln/@cap`) to its SVG `stroke-linecap` value.
 *
 * `flat` -> `butt`, `sq` -> `square`, `rnd` -> `round`. Anything absent or
 * unrecognised falls back to `round`, matching the viewer's historical default.
 */
export function svgLineCap(
	lineCap: 'flat' | 'rnd' | 'sq' | undefined,
): 'butt' | 'round' | 'square' {
	switch (lineCap) {
		case 'flat':
			return 'butt';
		case 'sq':
			return 'square';
		default:
			return 'round';
	}
}

/** Connector routing family, derived from the OOXML preset shape type. */
export type ConnectorKind = 'straight' | 'bent' | 'curved';

/**
 * Classify a connector by its OOXML preset shape type (case-insensitive):
 * `bentConnector*` → `"bent"`, `curvedConnector*` → `"curved"`, everything
 * else (incl. `straightConnector1`) → `"straight"`.
 */
export function connectorKind(shapeType: string | undefined): ConnectorKind {
	const t = (shapeType ?? '').toLowerCase();
	if (t.includes('bentconnector')) {
		return 'bent';
	}
	if (t.includes('curvedconnector')) {
		return 'curved';
	}
	return 'straight';
}

/**
 * Determine whether a connector `shapeType` requires multi-segment (path)
 * rendering rather than a simple `<line>` element.
 */
export function connectorNeedsPath(shapeType: string | undefined): boolean {
	return connectorKind(shapeType) !== 'straight';
}
