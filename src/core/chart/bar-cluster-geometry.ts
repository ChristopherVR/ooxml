/**
 * chart-bar-cluster-geometry.ts: clustered bar width and spacing from
 * `c:gapWidth` and `c:overlap`. Split out of `chart-cartesian-bars.ts` so the
 * bar lane of a bar + line combo (`chart-combo-classify.ts`) sizes its bars
 * the same way a plain bar chart does.
 *
 * @module chart-bar-cluster-geometry
 */
/**
 * ECMA-376 default of `c:gapWidth` when a bar chart omits it: the gap between
 * categories is 150% of one bar's width. Every bar renderer (2D clustered,
 * stacked and percentStacked, and the 3D layouts) falls back to this one value.
 */
export const DEFAULT_BAR_GAP_WIDTH = 150;

export interface BarClusterOptions {
	barGapWidth?: number;
	barOverlap?: number;
}

/** Validate authored OOXML integers without silently clamping a requested edit. */
export function assertBarClusterOptions(options: BarClusterOptions): void {
	for (const [field, min, max] of [
		['barGapWidth', 0, 500],
		['barOverlap', -100, 100],
	] as const) {
		const value = options[field];
		if (value !== undefined && (!Number.isInteger(value) || value < min || value > max))
			throw new RangeError(`${field} must be an integer from ${min} to ${max}`);
	}
}

export interface ClusteredBarGeometry {
	/** One bar's width. */
	singleBarWidth: number;
	/** Distance between the left edges of adjacent bars in a cluster. */
	step: number;
	/** Width of one category's whole (possibly overlapped) cluster. */
	clusterWidth: number;
}

/** Bar width, bar-to-bar step and total cluster width for one category slot. */
export function clusteredBarGeometry(
	barGroupWidth: number,
	seriesCount: number,
	chartData: BarClusterOptions,
): ClusteredBarGeometry {
	// Honour c:overlap (% overlap between adjacent series). overlap=0 reproduces
	// the original side-by-side layout exactly. Read before singleBarWidth: a
	// gapWidth-based bar width must size itself so the OVERLAPPED cluster (not
	// seriesCount side-by-side bars) fills the gap-reduced group width; see the
	// comment below.
	const overlap = chartData.barOverlap ?? 0;
	// Honour c:gapWidth (gap between clusters, % of a bar width); an absent
	// c:gapWidth takes the spec default, DEFAULT_BAR_GAP_WIDTH (150%).
	//
	// COM-verified ground truth (PowerPoint Object 16, single category, six
	// series, gapWidth=5, overlap=23): dividing by seriesCount alone (the
	// pre-existing formula) sizes every bar as if the cluster were laid out
	// SIDE BY SIDE, then shrinks it further by overlap when computing `step`
	// below - so at high overlap the bars render far too NARROW (in the
	// limit, overlap=100 should make every bar in a cluster the same width
	// as a single-series bar, independent of seriesCount, since they fully
	// coincide; the old formula kept shrinking by 1/seriesCount regardless).
	// `overlapSpan` is how many bar-widths wide the OVERLAPPED cluster spans;
	// `clusterWidth` below re-derives the same relationship from `step`.
	const overlapSpan = 1 + (seriesCount - 1) * (1 - overlap / 100);
	// COM-verified ground truth (PowerPoint's own Office-default 3-series
	// clustered column chart, gapWidth=219, overlap=-27, four categories):
	// the rendered bar is 17.6% of the category pitch. ECMA-376's own wording
	// for c:gapWidth is "the amount of space between bar or column clusters,
	// AS A PERCENTAGE OF THE BAR OR COLUMN WIDTH" - i.e. the gap between
	// clusters is `gapWidth% * singleBarWidth`, not a percentage of the pitch
	// or of the cluster width. So `pitch = clusterWidth + gap = barWidth *
	// overlapSpan + barWidth * gapWidth / 100 = barWidth * (overlapSpan +
	// gapWidth / 100)`. The previous formula divided the gap-shrunk group
	// width by `overlapSpan` (`pitch / ((1 + gapWidth / 100) * overlapSpan)`),
	// which treats the gap as a percentage of the PITCH and then shrinks the
	// cluster a second time, rendering every bar roughly half its correct
	// width whenever gapWidth and overlap are both non-trivial (this fixture:
	// 8.9% computed vs. 17.6% measured). Single-series charts are unaffected:
	// `overlapSpan` is 1 regardless of `overlap`, so both formulas agree.
	const gapWidth = chartData.barGapWidth ?? DEFAULT_BAR_GAP_WIDTH;
	const singleBarWidth = barGroupWidth / (overlapSpan + Math.max(gapWidth, 0) / 100);
	const step = singleBarWidth * (1 - overlap / 100);
	return { singleBarWidth, step, clusterWidth: singleBarWidth + step * (seriesCount - 1) };
}
