/**
 * The pptx chart adapter (docs/agnostic-core-plan.md, steps 5 and 6): pptx builds its chart model
 * (`PptxChartData`) from the neutral `c:chartSpace` model (`chart/parseChartSpace`) for classic
 * chart parts. This module holds the chart-level pieces: chart type, the first group's settings,
 * chrome flags, 3-D view, the style overlay and the pairing of each neutral chart group with its
 * object-tree node. Series, axes and caches are in `chart-from-neutral-series.ts`,
 * `chart-from-neutral-axes.ts` and `chart-from-neutral-data.ts`.
 *
 * What the neutral model does not carry (resolved colours and fonts, gradients, trendlines, error
 * bars, picture fills, c15/c16 extensions, layouts, the colour-style, chart-style, workbook and
 * user-shape parts, `c:clrMapOvr`, pivot source and formats, print settings, protection) is still
 * read from the part's `fast-xml-parser` tree, handed to the existing object-tree readers.
 *
 * @module chart-from-neutral
 */
import { readDisplayNaAsBlank } from '../../../chart/index';
import type { ChartPlotGroup, ChartSpace } from '../../../chart/index';
import type {
	PptxChartChrome,
	PptxChartData,
	PptxChartScatterStyle,
	PptxChartStyle,
	PptxChartType,
	PptxChartView3D,
	XmlObject,
} from '../types';
import { chartContainerLocalNameToType } from './chart-container-type-map';
import { labelOptionsFromNeutral, setOrRemove } from './chart-from-neutral-labels';
import { parseBar3DShapeVal, parseRadarStyleVal } from './chart-subtype-values';

export { axesFromNeutral } from './chart-from-neutral-axes';
export { seriesFromNeutral, type NeutralSeriesPlace } from './chart-from-neutral-series';
export {
	neutralCategories,
	neutralDateCategories,
	neutralTitleText,
} from './chart-from-neutral-data';
export type { ChartTreeReader } from './chart-series-tree-fields';

/** The pptx chart type of a group element (`barChart` -> `bar`, `surface3DChart` -> `surface`). */
export const groupChartType = (group: ChartPlotGroup): PptxChartType | undefined =>
	chartContainerLocalNameToType(group.element);

/** The chart type: the group's own, `combo` for two or more group elements, else `unknown`. */
export function chartTypeFromNeutral(space: ChartSpace): PptxChartType {
	const elements = new Set(space.plotArea.groups.map((group) => group.element));
	const first = space.plotArea.groups[0];
	if (elements.size >= 2) return 'combo';
	return (first && groupChartType(first)) ?? 'unknown';
}

const SCATTER_STYLES = new Set<string>([
	'none',
	'line',
	'lineMarker',
	'marker',
	'smooth',
	'smoothMarker',
]);

type GroupOptions = Pick<
	PptxChartData,
	| 'grouping'
	| 'groupingStandard'
	| 'varyColors'
	| 'firstSliceAngle'
	| 'doughnutHoleSize'
	| 'barGapWidth'
	| 'barOverlap'
	| 'gapDepth'
	| 'scatterStyle'
	| 'barDirection'
	| 'barShape'
	| 'radarStyle'
	| 'wireframe'
	| 'surfaceTopView'
>;

/**
 * The chart-level group settings, read from the first chart group (a combo chart's other groups
 * keep theirs in the part). Grouping is `stacked`, `percentStacked` or else `clustered`, with
 * `groupingStandard` marking a written `standard`; radar style, 3-D bar shape, wireframe and the
 * surface projection are read only for their own chart type.
 */
export function groupOptionsFromNeutral(
	group: ChartPlotGroup | undefined,
	chartType: PptxChartType,
): GroupOptions {
	const out: GroupOptions = {};
	if (!group) return out;
	const grouping = group.grouping?.trim();
	if (group.grouping) {
		if (grouping === 'stacked' || grouping === 'percentStacked') out.grouping = grouping;
		else {
			out.grouping = 'clustered';
			if (grouping === 'standard') out.groupingStandard = true;
		}
	}
	setOrRemove(out, 'varyColors', group.varyColors);
	setOrRemove(out, 'firstSliceAngle', group.firstSliceAngle);
	setOrRemove(out, 'doughnutHoleSize', group.holeSize);
	setOrRemove(out, 'barGapWidth', group.gapWidth);
	setOrRemove(out, 'barOverlap', group.overlap);
	setOrRemove(out, 'gapDepth', group.gapDepth);
	const scatter = group.scatterStyle?.trim();
	if (scatter && SCATTER_STYLES.has(scatter)) out.scatterStyle = scatter as PptxChartScatterStyle;
	setOrRemove(out, 'barDirection', group.barDirection);
	if (chartType === 'bar3D')
		setOrRemove(out, 'barShape', parseBar3DShapeVal(group.shape?.trim() ?? ''));
	if (chartType === 'radar')
		setOrRemove(out, 'radarStyle', parseRadarStyleVal(group.radarStyle?.trim() ?? ''));
	if (chartType === 'surface') {
		setOrRemove(out, 'wireframe', group.wireframe);
		out.surfaceTopView = group.element === 'surfaceChart';
	}
	return out;
}

/** `c:autoTitleDeleted`, `c:dispBlanksAs`, `c:showDLblsOverMax` and the c16r3 #N/A flag. */
export function chartChromeFromNeutral(space: ChartSpace): PptxChartChrome | undefined {
	const chrome: PptxChartChrome = {};
	setOrRemove(chrome, 'autoTitleDeleted', space.autoTitleDeleted);
	const blanks = space.displayBlanksAs;
	if (blanks === 'gap' || blanks === 'zero' || blanks === 'span') chrome.dispBlanksAs = blanks;
	setOrRemove(chrome, 'showDLblsOverMax', space.showDataLabelsOverMax);
	setOrRemove(chrome, 'dispNaAsBlank', readDisplayNaAsBlank(space.chartExtLst));
	return Object.keys(chrome).length > 0 ? chrome : undefined;
}

/** `c:view3D`; undefined when it sets nothing, so no empty element round-trips. */
export function view3DFromNeutral(space: ChartSpace): PptxChartView3D | undefined {
	const view = space.view3D;
	if (!view) return undefined;
	const out: PptxChartView3D = {};
	setOrRemove(out, 'rotX', view.rotX);
	setOrRemove(out, 'rotY', view.rotY);
	setOrRemove(out, 'depthPercent', view.depthPercent);
	setOrRemove(out, 'perspective', view.perspective);
	setOrRemove(out, 'hPercent', view.heightPercent);
	setOrRemove(out, 'rAngAx', view.rightAngleAxes);
	return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Fills the structural chart-style fields from the neutral model onto the object-tree style
 * (which keeps the resolved fills, borders, fonts and legend entry styles): `c:style`, legend
 * presence, position and overlay (a classic legend without `c:overlay` overlays), and the
 * switches of the first group's data labels.
 */
export function chartStyleFromNeutral(
	space: ChartSpace,
	styled: PptxChartStyle | undefined,
): PptxChartStyle | undefined {
	const style: PptxChartStyle = styled ?? {};
	setOrRemove(style, 'styleId', space.style);
	const legend = space.legend;
	setOrRemove(style, 'hasLegend', legend ? true : undefined);
	setOrRemove(style, 'legendPosition', legend?.position || undefined);
	setOrRemove(style, 'legendOverlay', legend ? (legend.overlay ?? true) : undefined);
	if (style.dataLabels) {
		const labels = space.plotArea.groups.find((group) => group.dataLabels)?.dataLabels;
		labelOptionsFromNeutral(labels, style.dataLabels);
	}
	return Object.keys(style).length > 0 ? style : undefined;
}

/** The raw children named `local` of an object-tree node, in order (bare elements as `{}`). */
export function rawChildren(
	parent: XmlObject | undefined,
	local: string,
	localName: (key: string) => string,
): XmlObject[] {
	if (!parent || typeof parent !== 'object') return [];
	const key = Object.keys(parent).find((candidate) => localName(candidate) === local);
	if (key === undefined) return [];
	const value = parent[key];
	const list = Array.isArray(value) ? value : [value];
	return list.map((entry) =>
		entry && typeof entry === 'object' && !Array.isArray(entry) ? (entry as XmlObject) : {},
	);
}

/** Each neutral chart group with its object-tree node (the same element at the same position). */
export function pairGroupNodes(
	groups: ChartPlotGroup[],
	plotArea: XmlObject,
	localName: (key: string) => string,
): { group: ChartPlotGroup; node: XmlObject }[] {
	const seen = new Map<string, number>();
	return groups.map((group) => {
		const at = seen.get(group.element) ?? 0;
		seen.set(group.element, at + 1);
		return { group, node: rawChildren(plotArea, group.element, localName)[at] ?? {} };
	});
}
