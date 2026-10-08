import type { ChartViewModel } from './chart-view';
import type { Frame } from './chart-svg-cartesian';
import { rect, n } from './chart-svg-util';
import { clusteredBarGeometry } from '../bar-cluster-geometry';
import { chartBarSpacing } from './chart-spacing';
import { buildChartGradientDef } from '../gradient-definition';
import { chartGradientMarkup } from '../gradient-markup';

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
			const width = model.horizontal ? Math.abs(b - a) : barSize;
			const height = model.horizontal ? barSize : Math.abs(b - a);
			let color = s.pointColors?.[i] ?? s.color;
			const gradient =
				s.pointGradients?.[i] ?? (s.pointColors?.[i] === undefined ? s.gradient : undefined);
			if (gradient && (gradient.path === 'circle' || gradient.path === 'shape')) {
				const baseId = color.match(/^url\(#(.+)\)$/)?.[1];
				if (baseId) {
					const id = baseId.replace(/(-s\d+(?:-p\d+)?)$/, `-mark${i}$1`);
					const def = buildChartGradientDef(id, gradient, {
						width: Number(n(width)),
						height: Number(n(height)),
						shape: 'rect',
					});
					out.push(`<defs>${chartGradientMarkup(def)}</defs>`);
					color = `url(#${id})`;
				}
			}
			out.push(`<g data-chart-series="${si}" data-chart-point="${i}">`);
			if (model.horizontal)
				out.push(
					rect(Math.min(a, b), center + offset, Math.abs(b - a), barSize, color, s.shadowFilter),
				);
			else
				out.push(
					rect(center + offset, Math.min(a, b), barSize, Math.abs(b - a), color, s.shadowFilter),
				);
			out.push('</g>');
		});
	});
}
