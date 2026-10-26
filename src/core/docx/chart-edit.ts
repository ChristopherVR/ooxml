import { chartIssuesLostOnWrite } from '../chart/keep-source';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Editing a Word chart: the shared chart edit operations (`chart/edit`) and data sheet applied to
// the `DocxChart` a drawing carries, and a new chart for Insert > Chart. Each edit returns a new
// chart whose title and direct formatting are read again from the edited part, so the painter
// shows exactly what will be saved; `chart-save.ts` writes edited parts and their workbooks.
// Extended charts (`cx:chartSpace`: waterfall, histogram...) and charts whose part did not load
// stay read-only.
import {
	applyChartSpaceDataSheetEdits,
	chartSpaceDataSheet,
	newChartSpace,
	readChartFormatting,
	writeChartSpace,
	type ChartDataSheet,
	type ChartDataSheetEdit,
	type ChartSpace,
	type ChartTypeTarget,
	type NewChartOptions,
} from '../chart/index';
import { parseXml } from '../xml/index';
import { CHART_NOTICE, type DocxChart } from './chart';
import type { InlineImage } from './inline-model';

/** Whether `chart` can be edited: a classic chart part that loaded. */
export function isDocxChartEditable(chart: DocxChart | undefined): chart is DocxChart & {
	chartSpace: ChartSpace;
} {
	return (
		Boolean(chart?.chartSpace) &&
		!chart?.layoutId &&
		!chart?.chartEx &&
		chartIssuesLostOnWrite(chart?.issues ?? []).length === 0
	);
}

/** `chart` showing `space`: title and direct formatting re-read from the part `space` writes. */
function withSpace(chart: DocxChart, space: ChartSpace): DocxChart {
	const { title: _title, formatting: _formatting, ...rest } = chart;
	const next: DocxChart = { ...rest, chartSpace: space };
	const title = space.autoTitleDeleted ? undefined : space.title?.text;
	if (title) next.title = title;
	if (space.plotArea.groups.length) {
		const root = parseXml(writeChartSpace(space), { label: 'Chart part' }).documentElement;
		const formatting = root ? readChartFormatting(root) : undefined;
		if (formatting) next.formatting = formatting;
	}
	return next;
}

/**
 * Applies a `chart/edit` operation (`setChartType`, `setChartTitle`, `setChartLegend`...) to the
 * chart. Throws when the chart is not editable; the operation's own errors propagate.
 */
export function editDocxChart(
	chart: DocxChart,
	operation: (space: ChartSpace) => ChartSpace,
): DocxChart {
	if (!isDocxChartEditable(chart)) throw new Error('This chart cannot be edited.');
	const space = operation(chart.chartSpace);
	return space === chart.chartSpace ? chart : withSpace(chart, space);
}

/** The data sheet Edit Data shows for the chart, or `undefined` when it is not editable. */
export function docxChartDataSheet(chart: DocxChart | undefined): ChartDataSheet | undefined {
	return isDocxChartEditable(chart) ? chartSpaceDataSheet(chart.chartSpace) : undefined;
}

/** Applies Edit Data edits; the chart itself when none applied. */
export function applyDocxChartDataSheetEdits(
	chart: DocxChart,
	edits: readonly ChartDataSheetEdit[],
): DocxChart {
	return editDocxChart(chart, (space) => applyChartSpaceDataSheetEdits(space, edits));
}

/**
 * A chart for Insert > Chart: Office's sample data (or `sheet`) as `target`. It has no part yet;
 * saving creates the chart part, its relationships and its embedded workbook.
 */
export function newDocxChart(
	target?: ChartTypeTarget,
	options: NewChartOptions & { sheet?: ChartDataSheet } = {},
): DocxChart {
	const { sheet, ...chartOptions } = options;
	return withSpace(
		{ placement: 'inline', notice: CHART_NOTICE, issues: [] },
		newChartSpace(sheet, target, chartOptions),
	);
}

/** Word's default size of a new inline chart: 6 x 3.5 inches (5486400 x 3200400 EMU). */
export const NEW_CHART_SIZE_PX = { width: 576, height: 336 } as const;

/** The inline drawing run of a new chart, `widthPx` wide (Word's 6 inches, else the line width). */
export function newChartImage(
	chart: DocxChart,
	widthPx: number = NEW_CHART_SIZE_PX.width,
): InlineImage {
	const width = Math.max(48, Math.round(Math.min(widthPx, NEW_CHART_SIZE_PX.width)));
	const height = Math.round((width * NEW_CHART_SIZE_PX.height) / NEW_CHART_SIZE_PX.width);
	return {
		relId: '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx: width,
		heightPx: height,
		unsupported: 'Chart',
		chart,
	};
}

/** Whether `chart` was inserted in this session and has no part in the package yet. */
export const isNewDocxChart = (chart: DocxChart): boolean => !chart.relId && !chart.partName;
