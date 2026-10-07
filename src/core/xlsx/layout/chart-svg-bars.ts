import type { ChartViewModel } from './chart-view';
import type { Frame } from './chart-svg-cartesian';
import { rect } from './chart-svg-util';
import { clusteredBarGeometry } from '../../chart/bar-cluster-geometry';
import { chartBarSpacing } from './chart-spacing';

export function bars(
	model: ChartViewModel,
	frame: Frame,
	out: string[],
	values: (number | null)[][],
): void {
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	const count = Math.max(1, model.series.length);
	const geometry = clusteredBarGeometry(frame.bandSize, count, chartBarSpacing(model));
	const barSize = geometry.singleBarWidth;
	const base = frame.value(0);
	const pos = new Array<number>(model.categories.length).fill(0);
	const neg = new Array<number>(model.categories.length).fill(0);
	model.series.forEach((s, si) => {
		values[si]?.forEach((v, i) => {
			if (v === null || i >= model.categories.length) return;
			let from = 0;
			let to = v;
			if (stacked) {
				const acc = v >= 0 ? pos : neg;
				from = acc[i] ?? 0;
				to = from + v;
				acc[i] = to;
			}
			const a = stacked ? frame.value(from) : base;
			const b = frame.value(to);
			const offset = model.horizontal
				? geometry.clusterWidth / 2 - barSize - si * geometry.step
				: -geometry.clusterWidth / 2 + si * geometry.step;
			const center = frame.band(i);
			out.push(`<g data-chart-series="${si}" data-chart-point="${i}">`);
			if (model.horizontal)
				out.push(
					rect(
						Math.min(a, b),
						center + offset,
						Math.abs(b - a),
						barSize,
						s.pointColors?.[i] ?? s.color,
						s.shadowFilter,
					),
				);
			else
				out.push(
					rect(
						center + offset,
						Math.min(a, b),
						barSize,
						Math.abs(b - a),
						s.pointColors?.[i] ?? s.color,
						s.shadowFilter,
					),
				);
			out.push('</g>');
		});
	});
}
