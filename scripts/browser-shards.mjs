/** Longest-first bin packing. History changes scheduling only, never test selection. */
export function balanceFiles(files, totals, shards) {
	if (!Number.isInteger(shards) || shards < 1) throw new Error('Invalid shard count');
	const known = Object.values(totals)
		.filter((n) => Number.isFinite(n) && n > 0)
		.sort((a, b) => a - b);
	const fallback = known[Math.floor(known.length / 2)] ?? 30_000;
	const weight = (file) =>
		Number.isFinite(totals[file]) && totals[file] > 0 ? totals[file] : fallback;
	const bins = Array.from({ length: shards }, () => ({ files: [], duration: 0 }));
	for (const file of [...new Set(files)].sort(
		(a, b) => weight(b) - weight(a) || a.localeCompare(b, 'en'),
	)) {
		const target = bins.reduce((best, bin) => (bin.duration < best.duration ? bin : best));
		target.files.push(file);
		target.duration += weight(file);
	}
	return bins;
}

export function fileDurations(history, project) {
	const totals = {};
	if (history?.version !== 1 || !history.tests || typeof history.tests !== 'object') return totals;
	for (const [key, ms] of Object.entries(history.tests)) {
		try {
			const [binding, file] = JSON.parse(key);
			if (binding === project && typeof file === 'string' && Number.isFinite(ms) && ms > 0)
				totals[file] = (totals[file] ?? 0) + ms;
		} catch {
			/* Old or malformed timing data is only a cache miss. */
		}
	}
	return totals;
}
