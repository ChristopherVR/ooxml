import {
	regionMapGeometry,
	regionBounds,
	regionViewBounds,
	scalePathD,
	resolveRegionCode,
	buildRegionMapEntries,
	normalizeValue,
	sequentialColorScale,
	resolveValueColorStops,
	buildValueColorScale,
	shouldRenderRegionLabel,
	formatRegionMapValue,
} from '../region-map';
import { text, rect, esc, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
export function regionMapSvg(model: ChartViewModel, area: Rect): string {
	const series = model.series[0];
	if (!series || series.hidden) return '';
	const options = model.regionMap;
	const geometry = regionMapGeometry(options);
	const resolve = (label: string) =>
		geometry.regions.find(
			(r) =>
				r.entityId === label || r.code === label || r.name.toLowerCase() === label.toLowerCase(),
		)?.code ?? resolveRegionCode(label);
	const entries = buildRegionMapEntries(model.categories, series.values, options, resolve);
	const min = entries.length ? Math.min(...entries.map((e) => e.value)) : 0,
		max = entries.length ? Math.max(...entries.map((e) => e.value)) : 1;
	const stops = resolveValueColorStops(
		options?.valueColors,
		options?.valueColorPositions,
		min,
		max,
	);
	const colors = stops ? buildValueColorScale(stops) : sequentialColorScale;
	const matched = new Map(entries.filter((e) => e.code).map((e) => [e.code!, e]));
	const missing = entries.filter(
		(e) => !e.code || !geometry.regions.some((r) => r.code === e.code),
	);
	const bounds = regionViewBounds(options?.viewedRegionType, matched, geometry.regions);
	const width = bounds.maxX - bounds.minX || 1,
		height = bounds.maxY - bounds.minY || 1;
	const reserve =
		(model.showLegend ? 46 : 12) +
		Math.min(missing.length, 5) * 14 +
		(geometry.approximate ? 14 : 0);
	const scale = Math.max(
		0.00001,
		Math.min(Math.max(1, area.w - 8) / width, Math.max(1, area.h - reserve) / height),
	);
	const dx = area.x + (area.w - width * scale) / 2 - bounds.minX * scale,
		dy = area.y - bounds.minY * scale;
	const out = [
		`<g data-chart-part="regionMap" data-map-geometry="${geometry.approximate ? 'approximate' : 'cached'}"${geometry.issues.length ? ' data-map-partial="true"' : ''}>`,
	];
	out.push(
		`<svg x="${area.x}" y="${area.y}" width="${area.w}" height="${Math.max(1, area.h - reserve)}" viewBox="${area.x} ${area.y} ${area.w} ${Math.max(1, area.h - reserve)}" overflow="hidden">`,
	);
	for (const region of geometry.regions) {
		const entry = matched.get(region.code);
		out.push(
			`<path d="${scalePathD(region.path, scale, dx, dy)}" fill="${esc(options?.fixedColor ?? (entry ? colors(normalizeValue(entry.value, min, max)) : '#e2e8f0'))}" stroke="#94a3b8" stroke-width="0.4" fill-rule="evenodd"${entry ? ` data-chart-series="0" data-chart-point="${entry.sourceIndex}"` : ''}><title>${esc(region.name)}${entry ? ': ' + esc(formatRegionMapValue(entry.value, options?.cultureLanguage)) : ''}</title></path>`,
		);
		const b = regionBounds(region);
		if (
			entry &&
			shouldRenderRegionLabel(
				options?.regionLabelLayout,
				(b.maxX - b.minX) * scale,
				(b.maxY - b.minY) * scale,
			)
		)
			out.push(
				text(
					region.labelXY[0] * scale + dx,
					region.labelXY[1] * scale + dy,
					formatRegionMapValue(entry.value, options?.cultureLanguage),
					{ size: 8, anchor: 'middle' },
				),
			);
	}
	out.push('</svg>');
	const y = area.y + height * scale + 5,
		bar = Math.min(160, area.w * 0.4),
		left = area.x + (area.w - bar) / 2;
	if (model.showLegend) {
		out.push(text(left - 8, y + 6, series.name, { size: 10, anchor: 'end' }));
		for (let i = 0; i < 32; i++)
			out.push(rect(left + (i * bar) / 32, y, bar / 32 + 0.1, 7, colors((i + 0.5) / 32)));
		out.push(
			text(left, y + 18, String(min), { size: 8 }),
			text(left + bar, y + 18, String(max), { size: 8, anchor: 'end' }),
		);
	}
	let row = y + (model.showLegend ? 32 : 0);
	if (geometry.approximate) {
		out.push(text(area.x, row, 'Approximate country outlines', { size: 8 }));
		row += 14;
	}
	for (const entry of missing.slice(0, 5)) {
		out.push(
			text(
				area.x,
				row,
				`${entry.label}: ${formatRegionMapValue(entry.value, options?.cultureLanguage)}`,
				{ size: 9 },
			),
		);
		row += 14;
	}
	const credits = [options?.attribution, ...geometry.copyright].filter(Boolean).join('; ');
	if (credits)
		out.push(
			`<title>${esc(credits)}</title>`,
			text(
				area.x,
				area.y + area.h - 1,
				credits.length > 180 ? credits.slice(0, 177) + '...' : credits,
				{ size: 6 },
			),
		);
	return out.join('') + '</g>';
}
