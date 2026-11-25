// Excel's source for the shared chart Format pane (`office-ui-chart-format-pane`): the chart's
// `ChartSpace` read through the neutral `chartSpaceFormatSource`, so the pane offers every element
// and option the `chart/edit` intents write (fill, border, series, axis, gridline, legend and data
// label options), as PowerPoint's pane does. Excel keeps three choices of its own: the fills of the
// chart area, plot area, title, legend and series are drawn by Excel's fill controls in the pane's
// `fill` slot (preset gradients, stops, directions), the spacing boxes have sliders, and an
// untitled chart (an empty title) lists no title.
import {
	chartBarSpacing,
	chartColorScheme,
	type ChartObject,
	type Workbook,
} from 'ooxml-core/xlsx';
import type { ChartFormatTarget, ChartRangeNumbers, ChartSpace } from 'ooxml-core/chart';
import type { DrawingColorTheme } from 'ooxml-core/drawingml';
import { chartSpaceFormatSource, type ChartFormatSource } from '../chart/format-pane-api';

/** The slot Excel's fill controls go into. */
export const XLSX_CHART_FILL_SLOT = 'fill';

/** The workbook theme as the pane resolves scheme colours (`accent1`, `tx1`...). */
export function workbookChartTheme(workbook: Workbook): DrawingColorTheme {
	const scheme = chartColorScheme(workbook.theme) as unknown as Readonly<Record<string, string>>;
	return { scheme: (name) => scheme[name] };
}

const FILL_PARTS: readonly string[] = ['chartArea', 'plotArea', 'title', 'legend', 'series'];

/** Whether Excel's own fill controls format `target`. */
export const xlsxFillsTarget = (target: ChartFormatTarget): boolean =>
	FILL_PARTS.includes(target.kind);

/** Pattern and picture fills keep the shared pane's own Fill section, with their detail fields. */
export const xlsxSlotsFill = (kind: string): boolean => kind !== 'pattern' && kind !== 'picture';

/**
 * The Format pane source of `chart` (whose model is `space`) with `target` selected: Excel's fill
 * controls take the Fill section of the elements they edit, and `resolveRange` reads the cells a
 * custom error-bar source names from the workbook.
 */
export function xlsxChartFormatSource(
	chart: ChartObject,
	space: ChartSpace,
	theme: DrawingColorTheme,
	target: ChartFormatTarget | undefined,
	resolveRange?: ChartRangeNumbers,
): ChartFormatSource {
	// Excel saves picture fills: the workbook writer stores the pictures as media parts.
	const source = chartSpaceFormatSource(space, {
		theme,
		pictures: true,
		...(resolveRange ? { resolveRange } : {}),
	});
	const spacing = chartBarSpacing(chart);
	return {
		...source,
		targets: chart.title
			? source.targets
			: source.targets.filter((option) => option.target.kind !== 'title'),
		seriesOptions: (seriesIndex) => {
			const view = source.seriesOptions?.(seriesIndex);
			if (!view || view.chartType !== 'bar') return view;
			return {
				...view,
				gapWidth: view.gapWidth ?? spacing.barGapWidth,
				overlap: view.overlap ?? spacing.barOverlap,
				sliders: true,
			};
		},
		...(target && xlsxFillsTarget(target) && xlsxSlotsFill(source.element(target).fill.kind)
			? { slots: { fill: XLSX_CHART_FILL_SLOT } }
			: {}),
		flat: true,
		picker: { key: 'chartFormat.chartElement', fallback: 'Chart element' },
	};
}
