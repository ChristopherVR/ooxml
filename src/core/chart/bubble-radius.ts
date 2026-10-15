/**
 * Bubble sizing shared by every format: how a `c:bubbleSize` value becomes a radius, given the
 * chart's `c:bubbleScale`, `c:sizeRepresents` and `c:showNegBubbles`. Moved from the PowerPoint
 * renderer (`chart-view-model-points.ts`) so the shared SVG painter sizes bubbles identically.
 */

/** `c:bubbleScale` / `c:sizeRepresents` inputs to {@link computeBubbleRadius}. */
export interface BubbleRadiusOptions {
	/** `c:bubbleScale` (0-300%). Office's own default, 100, leaves the envelope unscaled. */
	bubbleScale?: number | undefined;
	/**
	 * `c:sizeRepresents`. `'area'` (ECMA-376's own default) sizes bubbles so the rendered AREA is
	 * proportional to the value (radius scales with the square root of it); `'w'` sizes so the
	 * DIAMETER is proportional (radius scales linearly). Defaults to `'w'` here so the raw
	 * helper's three-argument callers stay unchanged; a chart-wide painter resolves the `'area'`
	 * default itself before calling in.
	 */
	sizeRepresents?: 'area' | 'w' | undefined;
}

/**
 * Radius of a bubble given its size value, the max size in the chart, and a median radius derived
 * from the plot area. With no size value the bubble uses the median radius (scaled by
 * `bubbleScale`); otherwise it scales from 0.5x to 2x the median, linearly (`'w'`) or by square
 * root (`'area'`, so the bubble's rendered AREA, not its diameter, tracks the value).
 */
export function computeBubbleRadius(
	sizeVal: number | undefined,
	maxBubble: number,
	medianRadius: number,
	options?: BubbleRadiusOptions,
): number {
	const scale = (options?.bubbleScale ?? 100) / 100;
	if (sizeVal === undefined) return medianRadius * scale;
	const denom = maxBubble > 0 ? maxBubble : 1;
	const magnitude = Math.abs(sizeVal);
	const ratio =
		(options?.sizeRepresents ?? 'w') === 'area'
			? Math.sqrt(magnitude) / Math.sqrt(denom)
			: magnitude / denom;
	return medianRadius * scale * 0.5 + ratio * medianRadius * scale * 1.5;
}

/** The size every bubble of the chart is measured against: the largest magnitude, at least 1. */
export function maxBubbleSize(sizes: readonly number[]): number {
	return sizes.length > 0 ? Math.max(1, ...sizes.map(Math.abs)) : 1;
}

/** The median bubble radius for a plot area (4% of its shorter side). */
export function medianBubbleRadius(plotWidth: number, plotHeight: number): number {
	return Math.min(plotWidth, plotHeight) * 0.04;
}
