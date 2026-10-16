/**
 * The Axis Options tab of a chart Format pane: bounds, units, crossing,
 * display units, log scale, order, tick marks, label position and number
 * format.
 *
 * @module render/chart-format/chart-format-pane-axis
 */
import type { PptxChartData } from 'ooxml-core/pptx';

import { choice, t } from './chart-format-pane-fill-line';
import type { ChartFormatField, ChartFormatSection } from './chart-format-pane-types';

const auto = (
	id: string,
	key: string,
	fallback: string,
	value: number | undefined,
	extra: { min?: number; max?: number; step?: number } = {},
): ChartFormatField => ({
	id,
	kind: 'number',
	label: t(key, fallback),
	value: value ?? '',
	auto: true,
	...extra,
});

const TICKS: Array<[string, string]> = [
	['none', 'None'],
	['in', 'Inside'],
	['out', 'Outside'],
	['cross', 'Cross'],
];
const tickChoices = () => TICKS.map(([value, label]) => choice(value, `tick.${value}`, label));

/** Axis Options, Tick Marks, Labels and Number sections for `data.axes[axisIndex]`. */
export function axisOptions(data: PptxChartData, axisIndex: number): ChartFormatSection[] {
	const axis = data.axes?.[axisIndex];
	if (!axis) {
		return [];
	}
	const value = axis.axisType === 'valAx' || axis.axisType === 'dateAx';
	const fields: ChartFormatField[] = [];
	if (value) {
		fields.push(
			auto('axis.min', 'minimum', 'Minimum', axis.min),
			auto('axis.max', 'maximum', 'Maximum', axis.max),
			auto('axis.majorUnit', 'majorUnit', 'Major unit', axis.majorUnit, { min: 0 }),
			auto('axis.minorUnit', 'minorUnit', 'Minor unit', axis.minorUnit, { min: 0 }),
		);
	}
	const crosses = axis.crossesAt !== undefined ? 'value' : (axis.crosses ?? 'autoZero');
	fields.push({
		id: 'axis.crosses',
		kind: 'select',
		label: value
			? t('crossesValue', 'Horizontal axis crosses')
			: t('crossesCategory', 'Vertical axis crosses'),
		value: crosses,
		choices: [
			choice('autoZero', 'crosses.auto', 'Automatic'),
			choice('min', 'crosses.min', 'At minimum'),
			choice('max', 'crosses.max', 'At maximum'),
			choice('value', 'crosses.value', 'At value'),
		],
	});
	if (crosses === 'value') {
		fields.push({
			id: 'axis.crossesAt',
			kind: 'number',
			label: t('crossesAt', 'Axis value'),
			value: axis.crossesAt ?? 0,
		});
	}
	if (axis.axisType === 'valAx') {
		fields.push(
			{
				id: 'axis.displayUnits',
				kind: 'select',
				label: t('displayUnits', 'Display units'),
				value: axis.displayUnits && axis.displayUnits !== 'custom' ? axis.displayUnits : 'none',
				choices: [
					choice('none', 'units.none', 'None'),
					choice('hundreds', 'units.hundreds', 'Hundreds'),
					choice('thousands', 'units.thousands', 'Thousands'),
					choice('tenThousands', 'units.tenThousands', 'Ten Thousands'),
					choice('hundredThousands', 'units.hundredThousands', 'Hundred Thousands'),
					choice('millions', 'units.millions', 'Millions'),
					choice('tenMillions', 'units.tenMillions', 'Ten Millions'),
					choice('hundredMillions', 'units.hundredMillions', 'Hundred Millions'),
					choice('billions', 'units.billions', 'Billions'),
					choice('trillions', 'units.trillions', 'Trillions'),
				],
			},
			{
				id: 'axis.logScale',
				kind: 'checkbox',
				label: t('logScale', 'Logarithmic scale'),
				value: Boolean(axis.logScale),
			},
		);
		if (axis.logScale) {
			fields.push({
				id: 'axis.logBase',
				kind: 'number',
				label: t('logBase', 'Base'),
				value: axis.logBase ?? 10,
				min: 2,
				max: 1000,
				step: 1,
			});
		}
	}
	fields.push({
		id: 'axis.reverseOrder',
		kind: 'checkbox',
		label: value
			? t('reverseValues', 'Values in reverse order')
			: t('reverseCategories', 'Categories in reverse order'),
		value: axis.orientation === 'maxMin',
	});
	return [
		{ id: 'axisOptions', title: t('section.axisOptions', 'Axis Options'), fields },
		{
			id: 'tickMarks',
			title: t('section.tickMarks', 'Tick Marks'),
			fields: [
				{
					id: 'axis.majorTickMark',
					kind: 'select',
					label: t('majorTick', 'Major type'),
					value: axis.majorTickMark ?? 'out',
					choices: tickChoices(),
				},
				{
					id: 'axis.minorTickMark',
					kind: 'select',
					label: t('minorTick', 'Minor type'),
					value: axis.minorTickMark ?? 'none',
					choices: tickChoices(),
				},
			],
		},
		{
			id: 'labels',
			title: t('section.labels', 'Labels'),
			fields: [
				{
					id: 'axis.tickLabelPosition',
					kind: 'select',
					label: t('labelPosition', 'Label Position'),
					value: axis.tickLblPos ?? 'nextTo',
					choices: [
						choice('nextTo', 'labelPos.nextTo', 'Next to Axis'),
						choice('high', 'labelPos.high', 'High'),
						choice('low', 'labelPos.low', 'Low'),
						choice('none', 'labelPos.none', 'None'),
					],
				},
			],
		},
		{
			id: 'number',
			title: t('section.number', 'Number'),
			fields: [
				{
					id: 'axis.numberFormat',
					kind: 'text',
					label: t('formatCode', 'Format Code'),
					value: axis.numFmt && !axis.numFmt.sourceLinked ? axis.numFmt.formatCode : '',
				},
			],
		},
	];
}
