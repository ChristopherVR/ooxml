/**
 * The data-label and font fields of a {@link ChartInput}: series labels
 * through `withChartDataLabelOptions` (flags and position, checked against
 * the positions the series' chart type allows), their number format and font
 * on the series' label options, their angle, wrapping, anchor and box fill through
 * `withChartElementFormat`; the title and legend fonts on the chart style.
 *
 * @module sdk/chart-content-labels
 */

import type { PptxChartData } from '../../types/chart';
import type { ChartPptxElement } from '../../types/elements';
import type { ChartDataLabelsInput } from './chart-content-input';
import { chartModelLabelPositions } from '../../utils/chart-label-positions';
import type { ChartElementFormatPatch } from './chart-element-format-operations';
import { withChartElementFormat } from './chart-element-format-operations';
import { withChartDataLabelOptions } from './chart-format-options-operations';
import { chartFontInput, chartTextRotationInput } from './chart-input-values';
import { toPaint } from './chart-layout-input';
import type { ChartInput, ChartSeriesInput } from './types';

const FLAGS = [
	'showValue',
	'showCategory',
	'showSeriesName',
	'showPercent',
	'showLegendKey',
] as const;

/** The chart-wide labels with the series' own on top, field by field. */
function mergedLabels(
	chart: ChartDataLabelsInput | undefined,
	own: ChartDataLabelsInput | undefined,
): ChartDataLabelsInput | undefined {
	if (!chart && !own) return undefined;
	const font = chart?.font || own?.font ? { font: { ...chart?.font, ...own?.font } } : {};
	return { ...chart, ...own, ...font };
}

/** @throws {RangeError} A position the series' chart group (and grouping) does not offer. */
function checkPosition(data: PptxChartData, index: number, position: string): void {
	const allowed = chartModelLabelPositions(data, index);
	if (!allowed || allowed.includes(position)) return;
	const type = data.series[index]!.seriesChartType ?? data.chartType;
	const grouping = data.grouping && data.grouping !== 'clustered' ? ` ${data.grouping}` : '';
	throw new RangeError(
		`Data label position "${position}" is not available on a${grouping} ${type} series; ` +
			(allowed.length ? `use one of ${allowed.join(', ')}.` : 'it takes no position.'),
	);
}

function labelFormat(labels: ChartDataLabelsInput): ChartElementFormatPatch {
	const patch: ChartElementFormatPatch = {};
	if (labels.rotation !== undefined || labels.wrap !== undefined || labels.anchor !== undefined) {
		if (labels.anchor !== undefined && !['t', 'ctr', 'b'].includes(labels.anchor))
			throw new RangeError(`Data label anchor must be t, ctr or b, got "${labels.anchor}".`);
		patch.text = {
			...(labels.rotation !== undefined
				? { rotation: chartTextRotationInput(labels.rotation) }
				: {}),
			...(labels.wrap !== undefined ? { wrap: labels.wrap } : {}),
			...(labels.anchor !== undefined ? { anchor: labels.anchor } : {}),
		};
	}
	if (labels.fill !== undefined) {
		const paint = toPaint(labels.fill, 'Data label fill');
		patch.fill = paint === 'none' ? { kind: 'none' } : { kind: 'solid', color: paint };
	}
	return patch;
}

function applySeriesLabels(
	data: PptxChartData,
	index: number,
	labels: ChartDataLabelsInput,
): PptxChartData {
	if (labels.position !== undefined) checkPosition(data, index, labels.position);
	// Labels with no content flag show the value, as PowerPoint's Data Labels button does.
	const shown = FLAGS.some((flag) => labels[flag] !== undefined);
	let next = withChartDataLabelOptions(
		data,
		{
			...Object.fromEntries(
				FLAGS.flatMap((flag) => (labels[flag] === undefined ? [] : [[flag, labels[flag]]])),
			),
			...(shown ? {} : { showValue: true }),
			...(labels.position !== undefined ? { position: labels.position } : {}),
		},
		index,
	);
	const options = next.series[index]!.dataLabelOptions!;
	if (labels.numberFormat !== undefined) options.numberFormat = labels.numberFormat;
	if (labels.font)
		options.txPr = { ...options.txPr, ...chartFontInput(labels.font, 'Data label font') };
	const patch = labelFormat(labels);
	if (patch.text || patch.fill) {
		next = withChartElementFormat(next, { kind: 'dataLabels', seriesIndex: index }, patch);
	}
	return next;
}

/** Apply `ChartInput.dataLabels` and each series' `dataLabels`. */
export function applyChartDataLabelsInput(element: ChartPptxElement, input: ChartInput): void {
	input.series.forEach((series: ChartSeriesInput, index) => {
		const labels = mergedLabels(input.dataLabels, series.dataLabels);
		if (labels) element.chartData = applySeriesLabels(element.chartData!, index, labels);
	});
}

/** `ChartInput.titleRuns` as the model's title runs; `title` is the first run's text, as on load. */
function applyTitleRuns(data: PptxChartData, input: ChartInput): void {
	const runs = input.titleRuns;
	if (!runs) return;
	if (runs.length === 0 || runs.some((run) => typeof run.text !== 'string'))
		throw new Error('titleRuns must be one or more runs, each with its text.');
	const text = runs.map((run) => run.text).join('');
	if (input.title !== undefined && input.title !== text)
		throw new Error('title and titleRuns name different titles; give titleRuns only.');
	// A loaded chart's title is its first run's text; the runs hold the rest.
	data.title = runs[0]!.text;
	data.titleRuns = runs.map((run, index) => {
		const font = chartFontInput(
			{ size: run.size, bold: run.bold, italic: run.italic, color: run.color },
			`Title run ${index + 1} font`,
		);
		return {
			text: run.text,
			...(font.bold !== undefined ? { bold: font.bold } : {}),
			...(font.italic !== undefined ? { italic: font.italic } : {}),
			...(font.fontSize !== undefined ? { fontSize: font.fontSize } : {}),
			...(font.color !== undefined ? { color: font.color } : {}),
		};
	});
}

/** Apply `ChartInput.titleRuns`, `ChartInput.titleFont` and `ChartInput.legendFont`. */
export function applyChartFontsInput(element: ChartPptxElement, input: ChartInput): void {
	const data = element.chartData!;
	applyTitleRuns(data, input);
	const style = (data.style ??= {});
	if (input.legendFont) {
		style.legendTextStyle = {
			...style.legendTextStyle,
			...chartFontInput(input.legendFont, 'Legend font'),
		};
	}
	if (!input.titleFont) return;
	const font = chartFontInput(input.titleFont, 'Title font');
	if (font.fontFamily !== undefined) style.titleFontFamily = font.fontFamily;
	if (font.fontSize !== undefined) style.titleFontSize = font.fontSize;
	if (font.bold !== undefined) style.titleFontBold = font.bold;
	if (font.color !== undefined) style.titleFontColor = font.color;
	if (font.italic !== undefined) {
		element.chartData = withChartElementFormat(
			data,
			{ kind: 'title' },
			{ text: { italic: font.italic } },
		);
	}
}
