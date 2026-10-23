/**
 * The element-specific Options tab of a chart Format pane: Series Options,
 * Axis Options, Legend Options and Label Options.
 *
 * @module chart/format-pane-options
 */
import type { ChartFormatTarget } from 'ooxml-core/chart';

import { axisOptions } from './format-pane-axis';
import { choice, t } from './format-pane-common';
import type { ChartFormatField, ChartFormatSection, ChartFormatText } from './format-pane-types';
import type {
	ChartFormatChartType,
	ChartFormatLabelsView,
	ChartFormatLegendView,
	ChartFormatSeriesOptionsView,
	ChartFormatSource,
} from './format-pane-view';

const pctField = (
	id: string,
	key: string,
	fallback: string,
	value: number,
	min: number,
	max: number,
	unit = '%',
	slider = false,
): ChartFormatField => ({
	id,
	kind: 'number',
	label: t(key, fallback),
	value,
	min,
	max,
	step: 1,
	unit,
	...(slider ? { slider } : {}),
});

/**
 * Series Options: Plot Series On (primary or secondary axis), gap width /
 * overlap for bars, slice angle / hole size for pies.
 */
function seriesOptions(series: ChartFormatSeriesOptionsView): ChartFormatSection[] {
	const type = series.chartType;
	const sliders = Boolean(series.sliders);
	const fields: ChartFormatField[] = [];
	if (series.secondaryAxis !== undefined) {
		fields.push({
			id: 'series.axis',
			kind: 'radio',
			label: t('plotSeriesOn', 'Plot Series On'),
			value: series.secondaryAxis ? 'secondary' : 'primary',
			choices: [
				choice('primary', 'axis.primary', 'Primary Axis'),
				choice('secondary', 'axis.secondary', 'Secondary Axis'),
			],
		});
	}
	if (type === 'bar' || type === 'bar3D') {
		if (type === 'bar') {
			fields.push(
				pctField(
					'series.overlap',
					'seriesOverlap',
					'Series Overlap',
					series.overlap ?? 0,
					-100,
					100,
					'%',
					sliders,
				),
			);
		}
		fields.push(
			pctField(
				'series.gapWidth',
				'gapWidth',
				'Gap Width',
				series.gapWidth ?? 150,
				0,
				500,
				'%',
				sliders,
			),
		);
	} else if (type === 'pie' || type === 'pie3D' || type === 'doughnut') {
		fields.push(
			pctField(
				'series.firstSliceAngle',
				'firstSliceAngle',
				'Angle of first slice',
				series.firstSliceAngle ?? 0,
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
					series.holeSize ?? 50,
					10,
					90,
				),
			);
		}
	} else if (type === 'ofPie') {
		fields.push(
			pctField('series.gapWidth', 'gapWidth', 'Gap Width', series.gapWidth ?? 100, 0, 500),
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

function legendOptions(legend: ChartFormatLegendView): ChartFormatSection[] {
	const fields: ChartFormatField[] = [
		{
			id: 'legend.position',
			kind: 'radio',
			label: t('legendPosition', 'Legend Position'),
			value: legend.position ?? 'r',
			choices: [
				choice('t', 'legendPos.top', 'Top'),
				choice('b', 'legendPos.bottom', 'Bottom'),
				choice('l', 'legendPos.left', 'Left'),
				choice('r', 'legendPos.right', 'Right'),
				choice('tr', 'legendPos.topRight', 'Top Right'),
			],
		},
	];
	if (legend.overlaySupported !== false) {
		fields.push({
			id: 'legend.noOverlap',
			kind: 'checkbox',
			label: t('legendNoOverlap', 'Show the legend without overlapping the chart'),
			value: legend.overlay === false,
		});
	}
	return [{ id: 'legendOptions', title: t('section.legendOptions', 'Legend Options'), fields }];
}

/** Label positions each chart family accepts (`c:dLblPos`). */
function labelPositions(type: ChartFormatChartType | undefined): Array<[string, string, string]> {
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

function labelOptions(labels: ChartFormatLabelsView): ChartFormatSection[] {
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
	const type = labels.chartType;
	if (type === 'pie' || type === 'pie3D' || type === 'doughnut' || type === 'ofPie') {
		fields.push(check('labels.showPercent', 'labelPercentage', 'Percentage', labels.showPercent));
	}
	fields.push(check('labels.showLegendKey', 'labelLegendKey', 'Legend key', labels.showLegendKey));
	const positions = labelPositions(type);
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
	const note: ChartFormatText = labels.seriesScoped
		? t('labelsOneSeries', 'Label options apply to this series only')
		: t('labelsAllSeries', 'Label options apply to every series');
	return [{ id: 'labelOptions', title: t('section.labelOptions', 'Label Options'), fields, note }];
}

type OptionsSource = Pick<
	ChartFormatSource,
	'seriesOptions' | 'axisOptions' | 'legendOptions' | 'labelOptions'
>;

/** The Options tab's sections for `target`, or `undefined` when the target has no options. */
export function optionsSections(
	source: OptionsSource,
	target: ChartFormatTarget,
): ChartFormatSection[] | undefined {
	switch (target.kind) {
		case 'series': {
			const view = source.seriesOptions?.(target.seriesIndex);
			return view && seriesOptions(view);
		}
		case 'axis': {
			const view = source.axisOptions?.(target.axisIndex);
			return view && axisOptions(view);
		}
		case 'legend': {
			const view = source.legendOptions?.();
			return view && legendOptions(view);
		}
		case 'dataLabels': {
			const view = source.labelOptions?.(target.seriesIndex);
			return view && labelOptions(view);
		}
		default:
			return undefined;
	}
}
