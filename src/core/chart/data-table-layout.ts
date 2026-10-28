/** Plot and data-table painters share one set of measured row/key dimensions. */
export function chartDataTableMetrics(fontSize: number, names: readonly string[], showKeys = true) {
	const rowH = Math.max(14, Math.round(fontSize * 1.45));
	const longestName = names.reduce((m, name) => Math.max(m, name.length), 0);
	const keyW = showKeys ? Math.min(160, Math.max(40, longestName * fontSize * 0.55 + 24)) : 0;
	return { fontSize, rowH, keyW, height: rowH * (1 + names.length) };
}
