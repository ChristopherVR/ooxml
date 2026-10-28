/** Flat, centred bars in source order, matching the existing PowerPoint funnel geometry. */
export function computeFunnelGeometry(
	values: readonly number[],
	plot: { x: number; y: number; width: number; height: number },
	gapWidth = 33,
) {
	const max = Math.max(...values.map((v) => Math.abs(v)), 1);
	const band = plot.height / Math.max(1, values.length);
	// Office ChartGroups.GapWidth defaults to 33 for a funnel. Keep half a gap at each edge.
	const gap = Number.isFinite(gapWidth) && gapWidth >= 0 ? gapWidth : 33;
	const height = band / (1 + gap / 100);
	const inset = (band - height) / 2;
	const center = plot.x + plot.width / 2;
	return values.map((value, index) => {
		const width = (Math.abs(value) / max) * plot.width;
		return {
			index,
			x: center - width / 2,
			y: plot.y + index * band + inset,
			width,
			height,
			center,
		};
	});
}
