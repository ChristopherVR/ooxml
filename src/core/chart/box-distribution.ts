import { computeBoxStats } from './box-stats';
/** Shared quartiles, whisker fences and mean; the host retains point and series identities. */
export function boxWhiskerDistribution(
	values: readonly number[],
	method: 'inclusive' | 'exclusive' = 'exclusive',
	trimOutliers = false,
) {
	const stats = computeBoxStats(values, method);
	if (!stats) return undefined;
	const iqr = stats.q3 - stats.q1;
	const lowerFence = stats.q1 - 1.5 * iqr;
	const upperFence = stats.q3 + 1.5 * iqr;
	const inliers = values.filter((v) => !trimOutliers || (v >= lowerFence && v <= upperFence));
	return {
		stats,
		whiskerMin: inliers.length ? Math.min(...inliers) : stats.min,
		whiskerMax: inliers.length ? Math.max(...inliers) : stats.max,
		mean: values.reduce((a, b) => a + b, 0) / values.length,
	};
}
