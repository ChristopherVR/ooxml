import { rect, text, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
import {
	chartTextWidth,
	chartFontMetrics,
	fitChartText,
	type ChartSvgOptions,
} from './chart-svg-text-metrics';
import { resolveManualLayoutRect } from '../manual-layout';

/** Keep the automatic reservation when a legend is moved; overlay releases it. */
export function chartLegendSvg(
	model: ChartViewModel,
	area: Rect,
	frame: { width: number; height: number },
	options: ChartSvgOptions,
): { plot: Rect; markup: string; foreground: boolean } {
	const automatic: string[] = [];
	const bounds: { box?: Rect } = {};
	const reserved = automaticLegend(model, area, automatic, options, bounds);
	const plot = model.legendOverlay ? area : reserved;
	const box = bounds.box;
	if (!box) return { plot, markup: '', foreground: false };
	const manual = resolveManualLayoutRect(model.legendLayout, frame, {
		x: box.x,
		y: box.y,
		width: box.w,
		height: box.h,
	});
	if (!manual) return { plot, markup: automatic.join(''), foreground: !!model.legendOverlay };
	const entries = legendEntries(model);
	const attrs = chartTextAttributes(model, 'legend');
	const size = attrs.size ?? 10;
	const font = chartFontMetrics(attrs, options);
	const keySize = size / 2;
	const keyGap = size / 4;
	const widths = entries.map(
		(entry) => chartTextWidth(entry.label, attrs, options) + keySize + keyGap,
	);
	const horizontal =
		manual.width > manual.height &&
		widths.reduce((sum, width) => sum + width, 0) + size <= manual.width;
	const rowWidth = horizontal ? manual.width / entries.length : manual.width;
	const rowHeight = horizontal ? manual.height : manual.height / entries.length;
	const horizontalGap = Math.max(
		0,
		(manual.width - widths.reduce((sum, width) => sum + width, 0)) / (entries.length + 1),
	);
	let horizontalX = manual.x + horizontalGap;
	const markup = [
		chartAreaRect(model, 'legend', { x: manual.x, y: manual.y, w: manual.width, h: manual.height }),
	];
	const baselineOffset = font ? (font.ascent - font.descent) / 2 : size * 0.35;
	entries.forEach((entry, i) => {
		const available = horizontal
			? chartTextWidth(entry.label, attrs, options)
			: Math.max(0, rowWidth - size / 2 - keySize - keyGap);
		const label = fitChartText(entry.label, available, attrs, options);
		const contentWidth = Math.min(
			horizontal ? (widths[i] ?? 0) : Math.max(...widths),
			rowWidth - size / 2,
		);
		const x = horizontal
			? horizontalX
			: manual.x + Math.max(size / 4, (rowWidth - contentWidth) / 2);
		const centerY = manual.y + (horizontal ? 0 : i * rowHeight) + rowHeight / 2;
		markup.push(rect(x, centerY - keySize / 2, keySize, keySize, entry.color));
		markup.push(text(x + keySize + keyGap, centerY + baselineOffset, label, attrs));
		if (horizontal) horizontalX += (widths[i] ?? 0) + horizontalGap;
	});
	return {
		plot,
		markup: `<g data-chart-legend-layout="true">${markup.join('')}</g>`,
		foreground: true,
	};
}

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
function automaticLegend(
	model: ChartViewModel,
	area: Rect,
	out: string[],
	options: ChartSvgOptions,
	bounds: { box?: Rect },
): Rect {
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
		bounds.box = { x: x - 4, y: y - 4, w: Math.min(total, area.w) + 8, h: rowH + 4 };
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
	bounds.box = { x: x - 4, y: y0, w: width + 4, h: Math.min(height, area.y + area.h - y0) };
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
