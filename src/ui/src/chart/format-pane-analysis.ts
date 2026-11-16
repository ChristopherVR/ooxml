import {
	CHART_MARKER_SYMBOLS,
	chartErrorBarSourceFromText,
	chartErrorBarSourceText,
	type ChartDataSource,
	type ChartEditIntent,
	type ChartErrorBars,
	type ChartFormatTarget,
} from 'ooxml-core/chart';
import { choice, t } from './format-pane-common';
import {
	chartFormatNumberField as number,
	chartTrendlineFieldIntent,
	chartTrendlineSections,
} from './format-pane-trendlines';
import type {
	ChartFormatField,
	ChartFormatFieldValue,
	ChartFormatSection,
} from './format-pane-types';
import type { ChartFormatSeriesOptionsView } from './format-pane-view';

const ERRORS = [
	['fixedVal', 'Fixed value'],
	['percentage', 'Percentage'],
	['stdDev', 'Standard deviation'],
	['stdErr', 'Standard error'],
	['cust', 'Custom source'],
] as const;

type Analysis = NonNullable<ChartFormatSeriesOptionsView['analysis']>;

/** The source box of one side of custom error bars. */
const sourceField = (side: 'plus' | 'minus', error: ChartErrorBars): ChartFormatField => ({
	id: `errorBars.${side}Values`,
	kind: 'text',
	label: t(
		`errorBars.${side}Values`,
		side === 'plus' ? 'Positive error value' : 'Negative error value',
	),
	value: chartErrorBarSourceText(error[side]),
});

function errorBarSection(analysis: Analysis): ChartFormatSection {
	const error = analysis.errorBars;
	const fields: ChartFormatField[] = [
		{
			id: 'errorBars.type',
			kind: 'select',
			label: t('errorBars.amount', 'Error amount'),
			value: error?.valueType ?? 'none',
			disabled: analysis.multipleErrorBars,
			choices: [
				choice('none', 'marker.none', 'None'),
				...ERRORS.map(([value, label]) =>
					choice(value, value === 'cust' ? 'errorBars.custom' : `errorBars.${value}`, label),
				),
			],
		},
	];
	if (error && !analysis.multipleErrorBars) {
		fields.push(
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
		if (error.valueType === 'cust') {
			if (error.barType !== 'minus') fields.push(sourceField('plus', error));
			if (error.barType !== 'plus') fields.push(sourceField('minus', error));
		} else if (error.valueType !== 'stdErr')
			fields.push(
				number('errorBars.value', 'errorBars.value', 'Error value', error.value ?? 1, 0, 100000),
			);
	}
	const note = analysis.multipleErrorBars
		? t(
				'analysis.multipleErrorBars',
				'Ambiguous error bars are preserved and cannot be edited here.',
			)
		: error?.valueType === 'cust'
			? !analysis.resolveRange
				? t('errorBars.customValuesHint', 'Enter values such as 1, 2, 3.')
				: analysis.workbookRanges
					? t(
							'errorBars.customHint',
							'Enter values such as 1, 2, 3 or a range such as Sheet1!$C$2:$C$4.',
						)
					: t(
							'errorBars.customCachedHint',
							'Enter values such as 1, 2, 3, or a range of cells this chart already plots.',
						)
			: undefined;
	return {
		id: 'errorBars',
		title: t('section.errorBars', 'Error bars'),
		fields,
		...(note ? { note } : {}),
	};
}

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
	// A product that does not edit trendlines here leaves `trendlines` out.
	if (analysis.trendlines) sections.push(...chartTrendlineSections(analysis.trendlines));
	sections.push(errorBarSection(analysis));
	return sections;
}

const ones = (count: number): ChartDataSource =>
	chartErrorBarSourceFromText('1', { pointCount: count })!;

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
	if (field.startsWith('trendline.'))
		return chartTrendlineFieldIntent(analysis.trendlines ?? [], seriesIndex, field, value);
	if (analysis.multipleErrorBars) return null;
	const wrap = (edit: Extract<ChartEditIntent, { op: 'errorBars' }>['edit']): ChartEditIntent => ({
		op: 'errorBars',
		seriesIndex,
		direction: 'y',
		edit,
	});
	const error = analysis.errorBars;
	const count = analysis.pointCount ?? 1;
	if (field === 'errorBars.type') {
		if (text === 'none') return error ? wrap(null) : null;
		if (!ERRORS.some(([type]) => type === text)) return null;
		// Custom starts from one on each side, the amount Office's Custom box first shows.
		return text === 'cust'
			? error?.valueType === 'cust'
				? null
				: wrap({ valueType: 'cust', plus: ones(count), minus: ones(count) })
			: wrap({ valueType: text as 'fixedVal' });
	}
	if (!error) return null;
	if (field === 'errorBars.cap') return wrap({ noEndCap: !value });
	if (field === 'errorBars.side') {
		if (!['both', 'plus', 'minus'].includes(text)) return null;
		const edit: Extract<ChartEditIntent, { op: 'errorBars' }>['edit'] = {
			barType: text as 'both',
		};
		// A custom side that becomes visible needs a source of its own.
		if (error.valueType === 'cust') {
			if (text !== 'minus' && !error.plus) edit.plus = ones(count);
			if (text !== 'plus' && !error.minus) edit.minus = ones(count);
		}
		return wrap(edit);
	}
	if (field === 'errorBars.plusValues' || field === 'errorBars.minusValues') {
		if (error.valueType !== 'cust') return null;
		const source = chartErrorBarSourceFromText(text, {
			pointCount: count,
			...(analysis.resolveRange ? { resolveRange: analysis.resolveRange } : {}),
		});
		return source
			? wrap(field === 'errorBars.plusValues' ? { plus: source } : { minus: source })
			: null;
	}
	return field === 'errorBars.value' && Number.isFinite(n) && n >= 0 ? wrap({ value: n }) : null;
}
