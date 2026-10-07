import { cartesianSvg } from './chart-svg-cartesian';
import { pieSvg, radarSvg } from './chart-svg-radial';
import { esc, n, rect, text, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes, chartGradientPaint } from './chart-svg-appearance';
import { chartRichTitleSvg } from './chart-svg-title-text';
import {
	CHART_FONT_FAMILY,
	chartTextWidth,
	chartFontMetrics,
	fitChartText,
	type ChartSvgOptions,
} from './chart-svg-text-metrics';
export type { ChartSvgOptions } from './chart-svg-text-metrics';

interface LegendEntry {
	label: string;
	color: string;
}

function legendEntries(model: ChartViewModel): LegendEntry[] {
	if (model.type === 'pie' || model.type === 'doughnut') {
		const first = model.series[0];
		return model.categories.map((label, i) => ({
			label,
			color: first?.pointColors?.[i] ?? first?.color ?? '#4472C4',
		}));
	}
	return model.series.map((s) => ({ label: s.name, color: s.color }));
}

/** Draws the legend and returns the area left for the plot. */
function legend(model: ChartViewModel, area: Rect, out: string[], options: ChartSvgOptions): Rect {
	const entries = legendEntries(model);
	if (!model.showLegend || entries.length === 0) return area;
	const pos = model.legendPosition;
	const attrs = chartTextAttributes(model, 'legend')!;
	const size = attrs.size ?? 10;
	const rowH = Math.max(16, size * 1.4);
	if (pos === 't' || pos === 'b') {
		const widths = entries.map((e) =>
			Math.min(chartTextWidth(e.label, attrs, options) + 20, area.w / 2),
		);
		const total = widths.reduce((a, b) => a + b, 0);
		let x = area.x + Math.max(0, (area.w - total) / 2);
		const y = pos === 't' ? area.y + 4 : area.y + area.h - rowH + 4;
		out.push(
			chartAreaRect(model, 'legend', {
				x: x - 4,
				y: y - 4,
				w: Math.min(total, area.w) + 8,
				h: rowH + 4,
			}),
		);
		entries.forEach((e, i) => {
			const w = widths[i] ?? 0;
			out.push(rect(x, y, 8, 8, e.color));
			out.push(text(x + 12, y + 8, fitChartText(e.label, w - 16, attrs, options), attrs));
			x += w;
		});
		return pos === 't'
			? { x: area.x, y: area.y + rowH + 4, w: area.w, h: area.h - rowH - 4 }
			: { x: area.x, y: area.y, w: area.w, h: area.h - rowH - 4 };
	}
	const width = Math.min(
		area.w * 0.35,
		Math.max(...entries.map((e) => chartTextWidth(e.label, attrs, options))) + 22,
	);
	const height = entries.length * rowH;
	const x = pos === 'l' ? area.x + 4 : area.x + area.w - width;
	const y0 = pos === 'tr' ? area.y + 4 : area.y + Math.max(0, (area.h - height) / 2);
	out.push(
		chartAreaRect(model, 'legend', {
			x: x - 4,
			y: y0,
			w: width + 4,
			h: Math.min(height, area.y + area.h - y0),
		}),
	);
	entries.forEach((e, i) => {
		const y = y0 + i * rowH;
		if (y + rowH > area.y + area.h + 2) return;
		out.push(rect(x, y + 3, 8, 8, e.color));
		out.push(text(x + 12, y + 11, fitChartText(e.label, width - 16, attrs, options), attrs));
	});
	return pos === 'l'
		? { x: area.x + width + 8, y: area.y, w: area.w - width - 8, h: area.h }
		: { x: area.x, y: area.y, w: area.w - width - 8, h: area.h };
}

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
	out.push(
		`<svg xmlns="http://www.w3.org/2000/svg" width="${n(w)}" height="${n(h)}" viewBox="0 0 ${n(w)} ${n(h)}" font-family="${esc(CHART_FONT_FAMILY)}" role="img">`,
	);
	if (model.title) out.push(`<title>${esc(model.title)}</title>`);
	if (paint.defs) out.push(paint.defs);
	out.push(chartAreaRect(model, 'chartArea', { x: 0, y: 0, w, h }));
	let area: Rect = { x: 8, y: 8, w: w - 16, h: h - 16 };
	const richTitle = chartRichTitleSvg(model, w, options);
	if (richTitle) {
		out.push(richTitle.markup);
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
		area = { x: area.x, y: area.y + titleHeight, w: area.w, h: area.h - titleHeight };
	}
	if (!model.supported) {
		out.push(
			text(w / 2, area.y + area.h / 2, `${model.type} charts are not drawn`, { anchor: 'middle' }),
		);
		out.push('</svg>');
		return out.join('');
	}
	area = legend(model, area, out, options);
	if (area.w > 4 && area.h > 4) {
		if (model.type === 'pie' || model.type === 'doughnut') out.push(pieSvg(model, area));
		else if (model.type === 'radar') out.push(radarSvg(model, area));
		else out.push(cartesianSvg(model, area));
	}
	out.push('</svg>');
	return out.join('');
}
