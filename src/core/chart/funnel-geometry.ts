/** Flat, centred bars in source order, matching the existing PowerPoint funnel geometry. */
export function computeFunnelGeometry(
	values: readonly number[],
	plot: { x: number; y: number; width: number; height: number },
) {
	const max = Math.max(...values.map((v) => Math.abs(v)), 1);
	const height = plot.height / Math.max(1, values.length);
	const center = plot.x + plot.width / 2;
	return values.map((value, index) => {
		const width = (Math.abs(value) / max) * plot.width;
		return { index, x: center - width / 2, y: plot.y + index * height, width, height, center };
	});
}
