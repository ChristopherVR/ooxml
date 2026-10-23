/**
 * insert-chart-preview.ts: the previews of the PowerPoint Insert Chart dialog. Each one is the
 * chart `createDefaultChartElement` would insert for the dialog's choices, drawn by the same view
 * model and SVG projector the slide uses (`buildChartViewModel`, `renderChartViewModelSvg`), and
 * handed to the shared `office-ui-chart-picker` as markup.
 *
 * @module insert-chart-preview
 */
import type { InsertChartGrouping, InsertChartKind } from './insert-chart';
import { createDefaultChartElement } from './insert-chart';
import { withResolvedPalette } from './chart-style-palette';
import { buildChartViewModel } from './chart-view-model-build';
import { renderChartViewModelSvg } from './chart-view-model-dom';
import { resolveChartKind } from './chart-view-model-kinds';

/** The choices a preview draws (the shared picker's `OfficeChartValue`). */
export interface InsertChartPreviewValue {
	type: string;
	grouping: string;
	title: string;
}

/**
 * SVG markup of the chart the dialog would insert, at `size`. A `tile` leaves out the title and
 * legend (the gallery scales it down). Empty for a chart the projector cannot draw.
 */
export function insertChartPreviewSvg(
	doc: Document,
	value: InsertChartPreviewValue,
	size: { width: number; height: number },
	tile: boolean,
	translate?: (key: string) => string,
): string {
	const element = createDefaultChartElement(
		value.type as InsertChartKind,
		{ x: 0, y: 0, width: size.width, height: size.height },
		{
			title: tile ? '' : value.title,
			hasLegend: !tile,
			grouping: value.grouping as InsertChartGrouping,
		},
	);
	const data = element.chartData;
	if (!data || resolveChartKind(data.chartType) === 'unsupported') return '';
	const kind = resolveChartKind(data.chartType);
	const square = kind === 'pie' || kind === 'doughnut' || kind === 'radar' || kind === 'regionMap';
	const svg = renderChartViewModelSvg(
		doc,
		buildChartViewModel({ ...element, chartData: withResolvedPalette(data) }),
		square ? 'xMidYMid meet' : 'none',
		translate,
	);
	svg.setAttribute('width', String(size.width));
	svg.setAttribute('height', String(size.height));
	const Serializer = doc.defaultView?.XMLSerializer ?? globalThis.XMLSerializer;
	return typeof Serializer === 'function' ? new Serializer().serializeToString(svg) : '';
}
