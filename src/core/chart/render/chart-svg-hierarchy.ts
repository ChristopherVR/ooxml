// Treemap and sunburst charts: the squarified cells and the concentric arc rings of the shared
// hierarchy geometry (`../hierarchy`, `../sunburst-arcs`), coloured by top-level branch.
import { buildHierarchyRoots, layoutTreemap } from '../hierarchy';
import { computeHierarchicalSunburstArcs, computeSunburstArcs } from '../sunburst-arcs';
import { chartAreaRect } from './chart-svg-appearance';
import { esc, fit, n, rect, text, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const colorAt = (model: ChartViewModel, index: number): string => {
	const first = model.series[0];
	const colors = first?.pointColors;
	return (
		(colors && colors.length > 0 ? colors[index % colors.length] : undefined) ??
		first?.color ??
		'#4472C4'
	);
};

/** Leaf values as the hierarchy weights them (magnitudes; missing values count as zero). */
const leafValues = (model: ChartViewModel): number[] =>
	(model.series[0]?.values ?? []).map((v) => v ?? 0);

export function treemapSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [chartAreaRect(model, 'plotArea', area)];
	const hierarchy = model.hierarchy;
	const series = model.series.map((s) => ({
		name: s.name,
		values: s.values.map((v) => v ?? 0),
		parentLabelLayout: hierarchy?.parentLabelLayout,
	}));
	const roots = buildHierarchyRoots(
		series,
		model.categories,
		hierarchy && hierarchy.levels.length > 1 ? hierarchy.levels : undefined,
	);
	const cells = layoutTreemap(roots, { x: area.x, y: area.y, w: area.w, h: area.h });
	for (const cell of cells) {
		if (cell.kind === 'parent') {
			out.push(text(cell.x, cell.y, cell.label, { size: 10, weight: 'bold', fill: '#303030' }));
			continue;
		}
		out.push(`<g data-chart-series="${cell.seriesIndex}" data-chart-point="${cell.pointIndex}">`);
		out.push(
			rect(cell.x, cell.y, cell.w, cell.h, colorAt(model, cell.colorIndex)).replace(
				'/>',
				' rx="2" fill-opacity="0.85"/>',
			),
		);
		if (cell.labelFits)
			out.push(
				text(
					cell.x + cell.w / 2,
					cell.y + cell.h / 2,
					fit(cell.label, cell.w - 4, cell.labelSize),
					{
						anchor: 'middle',
						size: cell.labelSize,
						weight: 'bold',
						fill: '#FFFFFF',
						baseline: 'central',
					},
				),
			);
		out.push('</g>');
	}
	return out.join('');
}

export function sunburstSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [chartAreaRect(model, 'plotArea', area)];
	const cx = area.x + area.w / 2;
	const cy = area.y + area.h / 2;
	const maxR = Math.min(area.w, area.h) / 2 - 4;
	if (maxR <= 0) return out.join('');
	const levels = model.hierarchy?.levels ?? [];
	const arcs =
		levels.length > 1
			? computeHierarchicalSunburstArcs(levels, leafValues(model), cx, cy, maxR)
			: computeSunburstArcs(
					model.series.map((s) => ({ values: s.values.map((v) => v ?? 0) })),
					cx,
					cy,
					maxR,
				);
	for (const arc of arcs)
		out.push(
			`<path d="${arc.d}" fill="${colorAt(model, arc.colorIndex)}" fill-opacity="${n(arc.opacity)}" stroke="#FFFFFF" stroke-width="1"${arc.pointIndex === undefined ? '' : ` data-chart-series="0" data-chart-point="${arc.pointIndex}"`}/>`,
		);
	// A flat sunburst labels its wedges with the categories of its only ring.
	const flatLabels = levels.length > 1 || model.series.length !== 1 ? undefined : model.categories;
	for (const arc of arcs) {
		const label =
			arc.label ?? (arc.pointIndex === undefined ? undefined : flatLabels?.[arc.pointIndex]);
		if (!label) continue;
		out.push(
			`<text x="${n(arc.labelX)}" y="${n(arc.labelY)}" font-size="10" fill="#FFFFFF" text-anchor="middle" dominant-baseline="central" transform="rotate(${n(arc.labelRotation)}, ${n(arc.labelX)}, ${n(arc.labelY)})">${esc(label)}</text>`,
		);
	}
	return out.join('');
}
