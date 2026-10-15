/**
 * The element-specific Options tab of a chart Format pane: Series Options,
 * Axis Options, Legend Options and Label Options.
 *
 * @module render/chart-format/chart-format-pane-options
 */
import type { PptxChartData, PptxChartFormatTarget, PptxChartType } from 'ooxml-core/pptx';

import { choice, t } from './chart-format-pane-fill-line';
import type { ChartFormatField, ChartFormatSection } from './chart-format-pane-types';
import type { ChartFormatText } from './chart-format-targets';
import { axisOptions } from './chart-format-pane-axis';

const pctField = (
	id: string,
	key: string,
	fallback: string,
	value: number,
	min: number,
	max: number,
	unit = '%',
): ChartFormatField => ({
	id,
	kind: 'number',
	label: t(key, fallback),
	value,
	min,
	max,
	step: 1,
	unit,
});

/** The chart type the series options follow, or undefined for a combo chart. */
function groupType(data: PptxChartData): PptxChartType | undefined {
	return data.chartType === 'combo' ? undefined : data.chartType;
}

/** Series Options: gap width / overlap for bars, slice angle / hole size for pies. */
function seriesOptions(data: PptxChartData): ChartFormatSection[] {
	const type = groupType(data);
	const fields: ChartFormatField[] = [];
	if (type === 'bar' || type === 'bar3D') {
		if (type === 'bar') {
			fields.push(
				pctField(
					'series.overlap',
					'seriesOverlap',
					'Series Overlap',
					data.barOverlap ?? 0,
					-100,
					100,
				),
			);
		}
		fields.push(
			pctField('series.gapWidth', 'gapWidth', 'Gap Width', data.barGapWidth ?? 150, 0, 500),
		);
	} else if (type === 'pie' || type === 'pie3D' || type === 'doughnut') {
		fields.push(
			pctField(
				'series.firstSliceAngle',
				'firstSliceAngle',
				'Angle of first slice',
				data.firstSliceAngle ?? 0,
				0,
				360,
				'°',
			),
		);
		if (type === 'doughnut') {
			fields.push(
				pctField(
					'series.holeSize',
					'holeSize',
					'Doughnut Hole Size',
					data.doughnutHoleSize ?? 50,
					10,
					90,
				),
			);
		}
	} else if (type === 'ofPie') {
		fields.push(
			pctField('series.gapWidth', 'gapWidth', 'Gap Width', data.barGapWidth ?? 100, 0, 500),
		);
	}
	const section: ChartFormatSection = {
		id: 'seriesOptions',
		title: t('section.seriesOptions', 'Series Options'),
		fields,
	};
	if (fields.length === 0) {
		section.note = t('noSeriesOptions', 'This chart type has no series spacing options');
	}
	return [section];
}

function legendOptions(data: PptxChartData): ChartFormatSection[] {
	return [
		{
			id: 'legendOptions',
			title: t('section.legendOptions', 'Legend Options'),
			fields: [
				{
					id: 'legend.position',
					kind: 'radio',
					label: t('legendPosition', 'Legend Position'),
					value: data.style?.legendPosition ?? 'r',
					choices: [
						choice('t', 'legendPos.top', 'Top'),
						choice('b', 'legendPos.bottom', 'Bottom'),
						choice('l', 'legendPos.left', 'Left'),
						choice('r', 'legendPos.right', 'Right'),
						choice('tr', 'legendPos.topRight', 'Top Right'),
					],
				},
				{
					id: 'legend.noOverlap',
					kind: 'checkbox',
					label: t('legendNoOverlap', 'Show the legend without overlapping the chart'),
					value: data.style?.legendOverlay === false,
				},
			],
		},
	];
}

/** Label positions each chart family accepts (`c:dLblPos`). */
function labelPositions(type: PptxChartType): Array<[string, string, string]> {
	if (type === 'pie' || type === 'pie3D' || type === 'ofPie') {
		return [
			['ctr', 'center', 'Center'],
			['inEnd', 'insideEnd', 'Inside End'],
			['outEnd', 'outsideEnd', 'Outside End'],
			['bestFit', 'bestFit', 'Best Fit'],
		];
	}
	if (type === 'bar') {
		return [
			['ctr', 'center', 'Center'],
			['inEnd', 'insideEnd', 'Inside End'],
			['inBase', 'insideBase', 'Inside Base'],
			['outEnd', 'outsideEnd', 'Outside End'],
		];
	}
	if (type === 'line' || type === 'scatter' || type === 'bubble') {
		return [
			['ctr', 'center', 'Center'],
			['l', 'left', 'Left'],
			['r', 'right', 'Right'],
			['t', 'above', 'Above'],
			['b', 'below', 'Below'],
		];
	}
	return [];
}

function labelOptions(data: PptxChartData): ChartFormatSection[] {
	const labels = data.style?.dataLabels ?? {};
	const check = (
		id: string,
		key: string,
		fallback: string,
		value: boolean | undefined,
	): ChartFormatField => ({
		id,
		kind: 'checkbox',
		label: t(key, fallback),
		value: Boolean(value),
	});
	const fields: ChartFormatField[] = [
		check('labels.showSeriesName', 'labelSeriesName', 'Series Name', labels.showSeriesName),
		check('labels.showCategory', 'labelCategoryName', 'Category Name', labels.showCategory),
		check('labels.showValue', 'labelValue', 'Value', labels.showValue),
	];
	if (
		data.chartType === 'pie' ||
		data.chartType === 'pie3D' ||
		data.chartType === 'doughnut' ||
		data.chartType === 'ofPie'
	) {
		fields.push(check('labels.showPercent', 'labelPercentage', 'Percentage', labels.showPercent));
	}
	fields.push(check('labels.showLegendKey', 'labelLegendKey', 'Legend key', labels.showLegendKey));
	const positions = labelPositions(data.chartType);
	if (positions.length > 0) {
		fields.push({
			id: 'labels.position',
			kind: 'radio',
			label: t('labelPosition', 'Label Position'),
			value: labels.position ?? '',
			choices: positions.map(([value, key, fallback]) =>
				choice(value, `dataLabelPos.${key}`, fallback),
			),
		});
	}
	const note: ChartFormatText = t('labelsAllSeries', 'Label options apply to every series');
	return [{ id: 'labelOptions', title: t('section.labelOptions', 'Label Options'), fields, note }];
}

/** The Options tab's sections for `target`, or `[]` when the target has no options. */
export function optionsSections(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartFormatSection[] {
	switch (target.kind) {
		case 'series':
			return seriesOptions(data);
		case 'axis':
			return axisOptions(data, target.axisIndex);
		case 'legend':
			return legendOptions(data);
		case 'dataLabels':
			return labelOptions(data);
		default:
			return [];
	}
}

/** Whether `target` has an Options tab. */
export function hasOptionsTab(target: PptxChartFormatTarget): boolean {
	return (
		target.kind === 'series' ||
		target.kind === 'axis' ||
		target.kind === 'legend' ||
		target.kind === 'dataLabels'
	);
}
