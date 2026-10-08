import { cartesianSvg } from './chart-svg-cartesian';
import { pieSvg, radarSvg } from './chart-svg-radial';
import { esc, n, text, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes, chartGradientPaint } from './chart-svg-appearance';
import { chartLegendSvg } from './chart-svg-legend';
import { chartRichTitleSvg } from './chart-svg-title-text';
import {
	CHART_FONT_FAMILY,
	chartTextWidth,
	chartFontMetrics,
	fitChartText,
	type ChartSvgOptions,
} from './chart-svg-text-metrics';
export type { ChartSvgOptions } from './chart-svg-text-metrics';

/**
 * A standalone SVG document for a chart view model: title, legend, axes and series. Every piece
 * of text is escaped. Unsupported chart types (bubble, stock, surface) render a labelled frame.
 */
export function renderChartSvg(
	model: ChartViewModel,
	width: number,
	height: number,
	options: ChartSvgOptions = {},
): string {
	const paint = chartGradientPaint(model, Math.max(1, width), Math.max(1, height));
	model = paint.model;
	const w = Math.max(1, width);
	const h = Math.max(1, height);
	const out: string[] = [];
	let foregroundTitle = '';
	out.push(
		`<svg xmlns="http://www.w3.org/2000/svg" width="${n(w)}" height="${n(h)}" viewBox="0 0 ${n(w)} ${n(h)}" font-family="${esc(CHART_FONT_FAMILY)}" role="img">`,
	);
	if (model.title) out.push(`<title>${esc(model.title)}</title>`);
	if (paint.defs) out.push(paint.defs);
	out.push(chartAreaRect(model, 'chartArea', { x: 0, y: 0, w, h }));
	let area: Rect = { x: 8, y: 8, w: w - 16, h: h - 16 };
	const richTitle = chartRichTitleSvg(model, w, options, h);
	if (richTitle) {
		if (model.titleLayout || model.titleOverlay) foregroundTitle = richTitle.markup;
		else out.push(richTitle.markup);
		if (!model.titleOverlay)
			area = { ...area, y: area.y + richTitle.height, h: area.h - richTitle.height };
	} else if (model.title) {
		const attrs = chartTextAttributes(model, 'title', 14)!;
		const size = attrs.size ?? 14;
		const label = fitChartText(model.title, w - 16, attrs, options);
		const titleWidth = Math.min(w - 16, chartTextWidth(label, attrs, options) + 8);
		const font = chartFontMetrics(attrs, options);
		const titleHeight = Math.max(24, font ? font.ascent + font.descent + 4 : size * 1.4 + 4);
		out.push(
			chartAreaRect(model, 'title', {
				x: (w - titleWidth) / 2,
				y: 4,
				w: Math.max(0, titleWidth),
				h: font ? font.ascent + font.descent + 4 : size * 1.4 + 4,
			}),
		);
		out.push(
			text(w / 2, 8 + size, label, {
				anchor: 'middle',
				fill: '#404040',
				...attrs,
			}),
		);
		if (!model.titleOverlay)
			area = { x: area.x, y: area.y + titleHeight, w: area.w, h: area.h - titleHeight };
	}
	if (!model.supported) {
		out.push(
			text(w / 2, area.y + area.h / 2, `${model.type} charts are not drawn`, { anchor: 'middle' }),
		);
		if (foregroundTitle) out.push(foregroundTitle);
		out.push('</svg>');
		return out.join('');
	}
	const legend = chartLegendSvg(model, area, { width: w, height: h }, options);
	area = legend.plot;
	if (!legend.foreground) out.push(legend.markup);
	if (area.w > 4 && area.h > 4) {
		if (model.type === 'pie' || model.type === 'doughnut') out.push(pieSvg(model, area));
		else if (model.type === 'radar') out.push(radarSvg(model, area));
		else out.push(cartesianSvg(model, area));
	}
	if (legend.foreground) out.push(legend.markup);
	if (foregroundTitle) out.push(foregroundTitle);
	out.push('</svg>');
	return out.join('');
}
