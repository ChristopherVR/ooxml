import {
	CHART_MARKER_SYMBOLS,
	type ChartEditIntent,
	type ChartFormatTarget,
} from 'ooxml-core/chart';
import { choice, t } from './format-pane-common';
import type {
	ChartFormatField,
	ChartFormatFieldValue,
	ChartFormatSection,
} from './format-pane-types';
import type { ChartFormatSeriesOptionsView } from './format-pane-view';

const TYPES = [
	['linear', 'Linear'],
	['exp', 'Exponential'],
	['log', 'Logarithmic'],
	['poly', 'Polynomial'],
	['power', 'Power'],
	['movingAvg', 'Moving average'],
] as const;
const ERRORS = [
	['fixedVal', 'Fixed value'],
	['percentage', 'Percentage'],
	['stdDev', 'Standard deviation'],
	['stdErr', 'Standard error'],
] as const;
const number = (
	id: string,
	key: string,
	fallback: string,
	value: number | '',
	min: number,
	max: number,
	auto = false,
): ChartFormatField => ({
	id,
	kind: 'number',
	label: t(key, fallback),
	value,
	min,
	max,
	step: 1,
	...(auto ? { auto: true } : {}),
});

export function chartAnalysisSections(series: ChartFormatSeriesOptionsView): ChartFormatSection[] {
	const sections: ChartFormatSection[] = [];
	if (series.marker)
		sections.push({
			id: 'markers',
			title: t('section.markers', 'Markers'),
			fields: [
				{
					id: 'marker.symbol',
					kind: 'select',
					label: t('marker.symbol', 'Marker'),
					value: series.marker.symbol ?? 'auto',
					choices: CHART_MARKER_SYMBOLS.map((symbol) => choice(symbol, `marker.${symbol}`, symbol)),
				},
				number('marker.size', 'marker.size', 'Marker size', series.marker.size ?? 5, 2, 72),
			],
		});
	const analysis = series.analysis;
	if (!analysis) return sections;
	const trend = analysis.trendline;
	const error = analysis.errorBars;
	const fields: ChartFormatField[] = [
		{
			id: 'trendline.type',
			kind: 'select',
			label: t('trendline.type', 'Trendline'),
			value: trend?.trendlineType ?? 'none',
			disabled: analysis.multipleTrendlines,
			choices: [
				choice('none', 'marker.none', 'None'),
				...TYPES.map(([value, label]) => choice(value, `trendline.${value}`, label)),
			],
		},
	];
	if (trend && !analysis.multipleTrendlines) {
		if (trend.trendlineType === 'poly')
			fields.push(
				number('trendline.order', 'trendline.order', 'Polynomial order', trend.order ?? 2, 2, 6),
			);
		if (trend.trendlineType === 'movingAvg')
			fields.push(
				number('trendline.period', 'trendline.period', 'Period', trend.period ?? 2, 2, 100000),
			);
		fields.push(
			number(
				'trendline.forward',
				'trendline.forward',
				'Forecast forward',
				trend.forward ?? 0,
				0,
				100000,
			),
			number(
				'trendline.backward',
				'trendline.backward',
				'Forecast backward',
				trend.backward ?? 0,
				0,
				100000,
			),
		);
		if (trend.trendlineType !== 'movingAvg')
			fields.push(
				{
					id: 'trendline.equation',
					kind: 'checkbox',
					label: t('trendline.equation', 'Display equation'),
					value: trend.displayEquation ?? false,
				},
				{
					id: 'trendline.rSquared',
					kind: 'checkbox',
					label: t('trendline.rSquared', 'Display R-squared'),
					value: trend.displayRSquared ?? false,
				},
			);
	}
	sections.push({
		id: 'trendlines',
		title: t('section.trendlines', 'Trendlines'),
		fields,
		...(analysis.multipleTrendlines
			? {
					note: t(
						'analysis.multipleTrendlines',
						'This series has multiple trendlines. Their settings are preserved.',
					),
				}
			: {}),
	});
	const errorFields: ChartFormatField[] = [
		{
			id: 'errorBars.type',
			kind: 'select',
			label: t('errorBars.amount', 'Error amount'),
			value: error?.valueType ?? 'none',
			disabled: analysis.multipleErrorBars,
			choices: [
				choice('none', 'marker.none', 'None'),
				...ERRORS.map(([value, label]) => choice(value, `errorBars.${value}`, label)),
				...(error?.valueType === 'cust'
					? [choice('cust', 'errorBars.custom', 'Custom source')]
					: []),
			],
		},
	];
	if (error && !analysis.multipleErrorBars) {
		errorFields.push(
			{
				id: 'errorBars.side',
				kind: 'select',
				label: t('errorBars.direction', 'Error bar direction'),
				value: error.barType ?? 'both',
				choices: [
					['both', 'Both'],
					['plus', 'Plus'],
					['minus', 'Minus'],
				].map(([value, label]) => choice(value!, `errorBars.${value}`, label!)),
			},
			{
				id: 'errorBars.cap',
				kind: 'checkbox',
				label: t('errorBars.cap', 'End caps'),
				value: !error.noEndCap,
			},
		);
		if (error.valueType !== 'cust' && error.valueType !== 'stdErr')
			errorFields.push(
				number('errorBars.value', 'errorBars.value', 'Error value', error.value ?? 1, 0, 100000),
			);
	}
	sections.push({
		id: 'errorBars',
		title: t('section.errorBars', 'Error bars'),
		fields: errorFields,
		...(analysis.multipleErrorBars
			? {
					note: t(
						'analysis.multipleErrorBars',
						'Ambiguous error bars are preserved and cannot be edited here.',
					),
				}
			: {}),
	});
	return sections;
}

/** Converts only the selected analysis property, retaining other trendlines, bars and styles. */
export function chartAnalysisFieldIntent(
	series: ChartFormatSeriesOptionsView | undefined,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null | undefined {
	if (!/^(marker|trendline|errorBars)\./u.test(field)) return undefined;
	if (!series || target.kind !== 'series') return null;
	const seriesIndex = target.seriesIndex;
	const text = String(value);
	const n = value === '' || typeof value === 'boolean' ? NaN : Number(value);
	if (field.startsWith('marker.')) {
		if (!series.marker) return null;
		if (field === 'marker.symbol')
			return CHART_MARKER_SYMBOLS.some((symbol) => symbol === text)
				? { op: 'marker', seriesIndex, edit: { symbol: text } }
				: null;
		return field === 'marker.size' && Number.isInteger(n) && n >= 2 && n <= 72
			? { op: 'marker', seriesIndex, edit: { size: n } }
			: null;
	}
	const analysis = series.analysis;
	if (!analysis) return null;
	if (field.startsWith('trendline.')) {
		if (analysis.multipleTrendlines) return null;
		const wrap = (
			edit: Extract<ChartEditIntent, { op: 'trendline' }>['edit'],
		): ChartEditIntent => ({ op: 'trendline', seriesIndex, index: 0, edit });
		if (field === 'trendline.type')
			return text === 'none'
				? analysis.trendline
					? wrap(null)
					: null
				: TYPES.some(([type]) => type === text)
					? wrap({ type: text as 'linear' })
					: null;
		if (!analysis.trendline) return null;
		if (field === 'trendline.equation') return wrap({ displayEquation: Boolean(value) });
		if (field === 'trendline.rSquared') return wrap({ displayRSquared: Boolean(value) });
		const key = field.slice(10);
		if (!['order', 'period', 'forward', 'backward'].includes(key) || !Number.isFinite(n) || n < 0)
			return null;
		return wrap({ [key]: n });
	}
	if (analysis.multipleErrorBars) return null;
	const wrap = (edit: Extract<ChartEditIntent, { op: 'errorBars' }>['edit']): ChartEditIntent => ({
		op: 'errorBars',
		seriesIndex,
		direction: 'y',
		edit,
	});
	if (field === 'errorBars.type')
		return text === 'none'
			? analysis.errorBars
				? wrap(null)
				: null
			: ERRORS.some(([type]) => type === text)
				? wrap({ valueType: text as 'fixedVal' })
				: null;
	if (!analysis.errorBars) return null;
	if (field === 'errorBars.cap') return wrap({ noEndCap: !value });
	if (field === 'errorBars.side')
		return ['both', 'plus', 'minus'].includes(text) ? wrap({ barType: text as 'both' }) : null;
	return field === 'errorBars.value' && Number.isFinite(n) && n >= 0 ? wrap({ value: n }) : null;
}
