import * as neutral from '../../../chart/index';
import type { ChartSpace } from '../../../chart/index';
import * as pptx from '../../core/builders/sdk/chart-operations';
import type { ChartPptxElement, PptxChartData } from '../../index';

/**
 * The edits the chart save harness (`chart-save-parity.test.ts`) makes: each is one PowerPoint
 * operation on the loaded model and the equivalent `chart/edit` operation on the parsed part,
 * grouped by the feature family whose writers it exercises.
 */
export type SaveFamily = 'none' | 'title' | 'legend' | 'axes' | 'series' | 'labels' | 'fill';

export interface SaveEdit {
	name: string;
	family: SaveFamily;
	/** Whether the edit applies to a part (default: every classic part). */
	applies?: (space: ChartSpace, chart: PptxChartData) => boolean;
	pptx: (element: ChartPptxElement) => void;
	/** The equivalent neutral edit; `chart` is the part as PowerPoint loaded it, before the edit. */
	neutral: (space: ChartSpace, chart: PptxChartData) => ChartSpace;
}

const kinds = (space: ChartSpace) => space.plotArea.groups.map((group) => group.kind);
const hasSeries = (space: ChartSpace) => space.plotArea.groups.some((g) => g.series.length > 0);
const hasValueAxis = (space: ChartSpace) => space.plotArea.axes.some((axis) => axis.kind === 'val');
/** Charts whose categories are text (scatter and bubble take numbers as x values). */
const textCategories = (space: ChartSpace) =>
	hasSeries(space) && kinds(space).every((kind) => kind !== 'scatter' && kind !== 'bubble');
const red = { kind: 'srgb' as const, value: 'FF0000', transforms: [] };

/** The first series' values, doubled plus one, so the point count stays as it was. */
const editedValues = (chart: PptxChartData): number[] =>
	(chart.series[0]?.values ?? []).map((value) => value * 2 + 1);

export const SAVE_EDITS: SaveEdit[] = [
	{
		name: 'no edit (slide dirty)',
		family: 'none',
		pptx: () => {},
		neutral: (space) => space,
	},
	{
		name: 'title text',
		family: 'title',
		pptx: (element) => pptx.setChartTitle(element, 'Edited title'),
		neutral: (space) => neutral.setChartTitle(space, 'Edited title'),
	},
	{
		name: 'title hidden',
		family: 'title',
		pptx: (element) => {
			element.chartData!.style = { ...element.chartData!.style, hasTitle: false };
		},
		neutral: (space) => neutral.setChartTitleVisible(space, false),
	},
	{
		name: 'legend position',
		family: 'legend',
		pptx: (element) => pptx.setChartLegend(element, { position: 't' }),
		neutral: (space) => neutral.setChartLegend(space, { position: 't' }),
	},
	{
		name: 'legend hidden',
		family: 'legend',
		pptx: (element) => pptx.setChartLegend(element, { show: false }),
		neutral: (space) => neutral.setChartLegend(space, { visible: false }),
	},
	{
		name: 'value axis scaling',
		family: 'axes',
		applies: hasValueAxis,
		pptx: (element) => pptx.setChartAxis(element, 'valAx', { min: 0, max: 990, majorUnit: 110 }),
		neutral: (space) => neutral.setChartAxis(space, 'val', { min: 0, max: 990, majorUnit: 110 }),
	},
	{
		name: 'value axis hidden',
		family: 'axes',
		applies: hasValueAxis,
		pptx: (element) => pptx.setChartAxis(element, 'valAx', { visible: false }),
		neutral: (space) => neutral.setChartAxis(space, 'val', { visible: false }),
	},
	{
		name: 'value axis number format',
		family: 'axes',
		applies: hasValueAxis,
		pptx: (element) => pptx.setChartAxis(element, 'valAx', { numberFormat: '0.0%' }),
		neutral: (space) => neutral.setChartAxis(space, 'val', { numberFormat: '0.0%' }),
	},
	{
		name: 'series values',
		family: 'series',
		applies: hasSeries,
		pptx: (element) => pptx.updateChartSeriesValues(element, 0, editedValues(element.chartData!)),
		neutral: (space, chart) => neutral.setChartSeriesValues(space, 0, editedValues(chart)),
	},
	{
		name: 'categories',
		family: 'series',
		applies: textCategories,
		pptx: (element) =>
			pptx.setChartCategories(
				element,
				element.chartData!.categories.map((_, i) => `Cat ${String(i + 1)}`),
			),
		neutral: (space, chart) =>
			neutral.setChartCategories(
				space,
				chart.categories.map((_, i) => `Cat ${String(i + 1)}`),
			),
	},
	{
		name: 'series name',
		family: 'series',
		applies: hasSeries,
		pptx: (element) => {
			element.chartData!.series[0]!.name = 'Renamed';
		},
		neutral: (space) => neutral.renameChartSeries(space, 0, 'Renamed'),
	},
	{
		name: 'data labels',
		family: 'labels',
		applies: (space) => hasSeries(space) && !kinds(space).includes('surface'),
		pptx: (element) => pptx.setChartDataLabels(element, { show: true }),
		neutral: (space) => neutral.setChartDataLabels(space, { show: true }),
	},
	{
		name: 'series colour',
		family: 'fill',
		applies: hasSeries,
		pptx: (element) => pptx.setChartSeriesColor(element, 0, '#FF0000'),
		neutral: (space) => neutral.setChartSeriesColor(space, 0, red),
	},
];
