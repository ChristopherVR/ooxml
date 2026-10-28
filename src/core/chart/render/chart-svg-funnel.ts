import { computeFunnelGeometry } from '../funnel-geometry';
import { chartTextAttributes, chartAreaRect } from './chart-svg-appearance';
import { rect, text, fit, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

export function funnelSvg(model: ChartViewModel, area: Rect): string {
	const series = model.series[0];
	if (!series || series.hidden) return '';
	const attrs = chartTextAttributes(model, 'categoryAxis', 10);
	const width = Math.min(
		area.w * 0.3,
		Math.max(32, ...model.categories.map((s) => s.length * (attrs.size ?? 10) * 0.55 + 8)),
	);
	const plot = { x: area.x + width, y: area.y, width: Math.max(1, area.w - width), height: area.h };
	const out = [
		chartAreaRect(model, 'plotArea', { x: plot.x, y: plot.y, w: plot.width, h: plot.height }),
	];
	for (const bar of computeFunnelGeometry(
		series.values.map((v) => v ?? 0),
		plot,
	)) {
		const value = series.values[bar.index];
		if (value === null || value === undefined) continue;
		out.push(`<g data-chart-series="${series.sourceIndex ?? 0}" data-chart-point="${bar.index}">`);
		out.push(
			rect(bar.x, bar.y, bar.width, bar.height, series.pointColors?.[bar.index] ?? series.color),
		);
		const size = Math.min(10, bar.height * 0.4);
		const label = series.labels?.[bar.index];
		if (label)
			out.push(
				text(bar.center, bar.y + bar.height / 2 + size * 0.35, label.text, {
					...chartTextAttributes(model, 'dataLabel', size),
					size,
					anchor: 'middle',
				}),
			);
		out.push(
			text(
				plot.x - 8,
				bar.y + bar.height / 2 + (attrs.size ?? 10) * 0.35,
				fit(model.categories[bar.index] ?? '', width - 8, attrs.size ?? 10),
				{ ...attrs, anchor: 'end' },
			),
		);
		out.push('</g>');
	}
	return out.join('');
}
