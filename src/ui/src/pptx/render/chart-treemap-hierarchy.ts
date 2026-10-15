import { buildHierarchyRoots, hierarchyBranchLabels, layoutTreemap } from 'ooxml-core/chart';
import type { TreemapBox } from 'ooxml-core/chart';
import type { PptxChartData } from 'ooxml-core/pptx';

import { dataLabelFontOverride, resolveDataLabelTextStyle } from './chart-data-label-text';
import type { SvgRect, SvgText } from './chart-view-model';
import { paletteColor } from './chart-view-model';

export type TreemapPrimitive = SvgRect | SvgText;

/**
 * The top-level branch labels a treemap's legend should show, in the same order colour indexes
 * are assigned to its roots (the rule lives in `ooxml-core/chart`, `hierarchy`).
 */
export function treemapBranchLabels(
	chartData: PptxChartData,
	categoryLabels: ReadonlyArray<string>,
): string[] {
	return hierarchyBranchLabels(chartData.series, categoryLabels, chartData.categoryLevels);
}

/**
 * Build nested treemap rectangles from ChartEx leaf-first category levels. The tree and the
 * squarified layout are the shared core's; this maps its cells onto SVG primitives.
 */
export function buildHierarchicalTreemapPrimitives(
	chartData: PptxChartData,
	categoryLabels: ReadonlyArray<string>,
	box: TreemapBox,
): TreemapPrimitive[] {
	const roots = buildHierarchyRoots(
		chartData.series.map((series) => ({
			name: series.name,
			values: series.values,
			parentLabelLayout: series.treemapOptions?.parentLabelLayout,
		})),
		categoryLabels,
		chartData.categoryLevels,
	);
	const primitives: TreemapPrimitive[] = [];
	for (const cell of layoutTreemap(roots, box)) {
		if (cell.kind === 'parent') {
			primitives.push({
				kind: 'text',
				x: cell.x,
				y: cell.y,
				text: cell.label,
				fontSize: 10,
				fill: '#303030',
				textAnchor: 'start',
				fontWeight: 'bold',
			});
			continue;
		}
		primitives.push({
			kind: 'rect',
			x: cell.x,
			y: cell.y,
			w: cell.w,
			h: cell.h,
			fill: paletteColor(cell.colorIndex, chartData.colorPalette),
			rx: 2,
			opacity: 0.85,
			part: { role: 'dataPoint', seriesIndex: cell.seriesIndex, pointIndex: cell.pointIndex },
		});
		if (cell.labelFits) {
			const leafSeries = chartData.series[cell.seriesIndex];
			primitives.push({
				kind: 'text',
				x: cell.x + cell.w / 2,
				y: cell.y + cell.h / 2,
				text: cell.label,
				fontSize: cell.labelSize,
				fill: '#ffffff',
				textAnchor: 'middle',
				fontWeight: 'bold',
				dominantBaseline: 'central',
				...(leafSeries
					? dataLabelFontOverride(resolveDataLabelTextStyle(chartData, leafSeries, cell.pointIndex))
					: {}),
			});
		}
	}
	return primitives;
}
