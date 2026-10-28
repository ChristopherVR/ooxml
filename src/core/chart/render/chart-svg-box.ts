import { boxWhiskerDistribution, boxMeanMarkerPath } from '../box-distribution';
import { groupRowsByCategory } from '../box-stats';
import type { ChartViewModel } from './chart-view';
import type { Frame } from './chart-svg-cartesian';
import { line, rect, n } from './chart-svg-util';

/** One box per series/category distribution; observations retain their original row identity. */
export function boxWhiskerSvg(model: ChartViewModel, frame: Frame, out: string[]): void {
	const visible = model.series.filter((s) => !s.hidden);
	const slot = frame.bandSize / Math.max(1, visible.length);
	const width = slot * 0.7;
	visible.forEach((series, si) => {
		if (series.hidden) return;
		const groups = groupRowsByCategory(series.boxCategories ?? []);
		model.categories.forEach((category, ci) => {
			const observations = (groups.rowIndexesByCategory.get(category) ?? []).flatMap((row) => {
				const value = series.values[row];
				return value === null || value === undefined ? [] : [{ row, value }];
			});
			const options = series.boxWhisker;
			const rawMethod = options?.statistics?.quartileMethod;
			const method = rawMethod === 'inclusive' ? 'inclusive' : 'exclusive';
			const visibility = options?.visibility;
			const trim =
				rawMethod === 'inclusive' ||
				rawMethod === 'exclusive' ||
				(!!visibility && Object.values(visibility).some((v) => v !== undefined));
			const distribution = boxWhiskerDistribution(
				observations.map((p) => p.value),
				method,
				trim,
			);
			if (!distribution) return;
			const { stats, whiskerMin, whiskerMax, mean } = distribution;
			const x = frame.band(ci) - frame.bandSize / 2 + si * slot + (slot - width) / 2;
			const mid = x + width / 2;
			const y = frame.value;
			out.push(
				`<g data-chart-part="box" data-chart-series="${series.sourceIndex ?? si}" data-chart-point="${ci}">`,
			);
			out.push(line(mid, y(whiskerMax), mid, y(stats.q3), '#64748b', 1));
			out.push(line(mid, y(stats.q1), mid, y(whiskerMin), '#64748b', 1));
			out.push(
				line(x + width * 0.25, y(whiskerMin), x + width * 0.75, y(whiskerMin), '#64748b', 1),
			);
			out.push(
				line(x + width * 0.25, y(whiskerMax), x + width * 0.75, y(whiskerMax), '#64748b', 1),
			);
			out.push(
				rect(
					x,
					Math.min(y(stats.q1), y(stats.q3)),
					width,
					Math.abs(y(stats.q1) - y(stats.q3)),
					series.pointColors?.[ci] ?? series.color,
				),
			);
			out.push(line(x, y(stats.median), x + width, y(stats.median), '#1e293b', 2));
			if (visibility?.meanLine) out.push(line(x, y(mean), x + width, y(mean), '#0f766e', 1.5));
			if (visibility?.meanMarker !== false)
				out.push(
					`<path d="${boxMeanMarkerPath(mid, y(mean))}" fill="none" stroke="#0f766e" stroke-width="1.5" data-chart-part="mean"/>`,
				);
			observations.forEach((point, index) => {
				const outlier = point.value < whiskerMin || point.value > whiskerMax;
				if (outlier ? !visibility?.outliers : !visibility?.nonoutliers) return;
				const px = x + width * (0.2 + (0.6 * (index + 1)) / (observations.length + 1));
				out.push(
					`<circle cx="${n(px)}" cy="${n(y(point.value))}" r="2.25" fill="${outlier ? '#dc2626' : '#1e293b'}" data-chart-series="${series.sourceIndex ?? si}" data-chart-point="${point.row}" data-chart-part="${outlier ? 'outlier' : 'observation'}"/>`,
				);
			});
			out.push('</g>');
		});
	});
}
