import { cartesianSvg } from './chart-svg-cartesian.js';
import { pieSvg, radarSvg } from './chart-svg-radial.js';
import { esc, fit, n, rect, text, textWidth, type Rect } from './chart-svg-util.js';
import type { ChartViewModel } from './chart-view.js';

const FONT_FAMILY = 'Calibri, Carlito, Arial, sans-serif';

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
function legend(model: ChartViewModel, area: Rect, out: string[]): Rect {
	const entries = legendEntries(model);
	if (!model.showLegend || entries.length === 0) return area;
	const pos = model.legendPosition;
	const rowH = 16;
	if (pos === 't' || pos === 'b') {
		const widths = entries.map((e) => Math.min(textWidth(e.label) + 20, area.w / 2));
		const total = widths.reduce((a, b) => a + b, 0);
		let x = area.x + Math.max(0, (area.w - total) / 2);
		const y = pos === 't' ? area.y + 4 : area.y + area.h - rowH + 4;
		entries.forEach((e, i) => {
			const w = widths[i] ?? 0;
			out.push(rect(x, y, 8, 8, e.color));
			out.push(text(x + 12, y + 8, fit(e.label, w - 16)));
			x += w;
		});
		return pos === 't'
			? { x: area.x, y: area.y + rowH + 4, w: area.w, h: area.h - rowH - 4 }
			: { x: area.x, y: area.y, w: area.w, h: area.h - rowH - 4 };
	}
	const width = Math.min(area.w * 0.35, Math.max(...entries.map((e) => textWidth(e.label))) + 22);
	const height = entries.length * rowH;
	const x = pos === 'l' ? area.x + 4 : area.x + area.w - width;
	const y0 = pos === 'tr' ? area.y + 4 : area.y + Math.max(0, (area.h - height) / 2);
	entries.forEach((e, i) => {
		const y = y0 + i * rowH;
		if (y + rowH > area.y + area.h + 2) return;
		out.push(rect(x, y + 3, 8, 8, e.color));
		out.push(text(x + 12, y + 11, fit(e.label, width - 16)));
	});
	return pos === 'l'
		? { x: area.x + width + 8, y: area.y, w: area.w - width - 8, h: area.h }
		: { x: area.x, y: area.y, w: area.w - width - 8, h: area.h };
}

/**
 * A standalone SVG document for a chart view model: title, legend, axes and series. Every piece
 * of text is escaped. Unsupported chart types (bubble, stock, surface) render a labelled frame.
 */
export function renderChartSvg(model: ChartViewModel, width: number, height: number): string {
	const w = Math.max(1, width);
	const h = Math.max(1, height);
	const out: string[] = [];
	out.push(
		`<svg xmlns="http://www.w3.org/2000/svg" width="${n(w)}" height="${n(h)}" viewBox="0 0 ${n(w)} ${n(h)}" font-family="${esc(FONT_FAMILY)}" role="img">`,
	);
	if (model.title) out.push(`<title>${esc(model.title)}</title>`);
	out.push(rect(0, 0, w, h, '#FFFFFF'));
	let area: Rect = { x: 8, y: 8, w: w - 16, h: h - 16 };
	if (model.title) {
		out.push(
			text(w / 2, 22, fit(model.title, w - 16, 14), {
				anchor: 'middle',
				size: 14,
				fill: '#404040',
			}),
		);
		area = { x: area.x, y: area.y + 24, w: area.w, h: area.h - 24 };
	}
	if (!model.supported) {
		out.push(
			text(w / 2, area.y + area.h / 2, `${model.type} charts are not drawn`, { anchor: 'middle' }),
		);
		out.push('</svg>');
		return out.join('');
	}
	area = legend(model, area, out);
	if (area.w > 4 && area.h > 4) {
		if (model.type === 'pie' || model.type === 'doughnut') out.push(pieSvg(model, area));
		else if (model.type === 'radar') out.push(radarSvg(model, area));
		else out.push(cartesianSvg(model, area));
	}
	out.push('</svg>');
	return out.join('');
}
