import type { ChartSummary, ChartThemePalette } from './summary';
import type { ChartDataTable } from '../model-chrome';
import { chartAppearance, type ChartAppearanceEntry } from './chart-appearance';
import { chartRunStyle } from './chart-title-text';
import { strokeOf, type ChartStrokeView } from './chart-view-analysis';
import type { DrawingColor } from '../../drawingml/types';

export interface ChartDataTableView {
	source: ChartDataTable;
	stroke: ChartStrokeView;
	fill?: string;
	text?: ChartAppearanceEntry;
}
export function chartDataTableView(
	chart: ChartSummary,
	theme: ChartThemePalette,
	color: (c: DrawingColor | undefined) => string | undefined,
): ChartDataTableView | undefined {
	const source = chart.dataTable;
	if (!source) return undefined;
	const props = source.txPr?.paragraphs[0];
	const direct = {
		sourceXml: '',
		...chartRunStyle(props?.defaultProperties ?? {}),
		...chartRunStyle(props?.runs[0] ?? {}),
	};
	const formatting = chart.formatting;
	const appearance = chartAppearance(
		{
			...chart,
			formatting: {
				...formatting,
				sourceXml: formatting?.sourceXml ?? '',
				entries: { ...formatting?.entries, categoryAxis: direct },
			},
		},
		theme,
	);
	const fill = source.spPr?.fill;
	return {
		source,
		stroke: strokeOf(source.spPr?.line, { color: '#D9D9D9', width: 1 }, color) ?? {
			color: 'none',
			width: 0,
		},
		...(fill?.kind === 'solid' && color(fill.color) ? { fill: color(fill.color)! } : {}),
		...(appearance?.categoryAxis ? { text: appearance.categoryAxis } : {}),
	};
}
