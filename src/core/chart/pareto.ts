// Pareto arithmetic shared by every painter: Office draws a pareto chart's bars in descending order
// of value and its `paretoLine` series as the running share of the total, on a 0 to 100% axis
// (the third `cx:axis`, `cx:units unit="percentage"`). Negative values count as zero.

/** The display order of pareto bars: descending value, ties in source order. */
export function paretoOrder(values: readonly number[]): number[] {
	return values
		.map((value, index) => ({ value: Math.max(0, Number.isFinite(value) ? value : 0), index }))
		.sort((a, b) => b.value - a.value || a.index - b.index)
		.map((entry) => entry.index);
}

/** The cumulative share (0 to 1) of each value in the order given; all zeros when nothing counts. */
export function paretoCumulative(values: readonly number[]): number[] {
	const counted = values.map((value) => Math.max(0, Number.isFinite(value) ? value : 0));
	const total = counted.reduce((sum, value) => sum + value, 0);
	if (total <= 0) return counted.map(() => 0);
	let running = 0;
	return counted.map((value, index) => {
		running += value;
		return index === counted.length - 1 ? 1 : running / total;
	});
}
