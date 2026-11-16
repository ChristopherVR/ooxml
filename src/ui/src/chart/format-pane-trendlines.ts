/**
 * The Trendlines sections of a series' Format pane: one section per trendline of the series, an
 * Add trendline choice, and the intents their fields send. The first trendline's fields keep the
 * ids `trendline.<option>`; a later one's are `trendline.<n>.<option>` (n from 1), so each field
 * edits exactly the trendline it shows and leaves the others as they are.
 *
 * @module chart/format-pane-trendlines
 */
import type { ChartEditIntent, ChartTrendline } from 'ooxml-core/chart';
import { choice, t } from './format-pane-common';
import type {
	ChartFormatField,
	ChartFormatFieldValue,
	ChartFormatSection,
} from './format-pane-types';

export const CHART_TRENDLINE_TYPES = [
	['linear', 'Linear'],
	['exp', 'Exponential'],
	['log', 'Logarithmic'],
	['poly', 'Polynomial'],
	['power', 'Power'],
	['movingAvg', 'Moving average'],
] as const;

const typeChoices = () =>
	CHART_TRENDLINE_TYPES.map(([value, label]) => choice(value, `trendline.${value}`, label));

export const chartFormatNumberField = (
	id: string,
	key: string,
	fallback: string,
	value: number | '',
	min: number,
	max: number,
): ChartFormatField => ({ id, kind: 'number', label: t(key, fallback), value, min, max, step: 1 });

/** The option fields of one trendline, their ids starting with `prefix`. */
function trendlineFields(prefix: string, trend: ChartTrendline | undefined): ChartFormatField[] {
	const fields: ChartFormatField[] = [
		{
			id: `${prefix}type`,
			kind: 'select',
			label: t('trendline.type', 'Trendline'),
			value: trend?.trendlineType ?? 'none',
			choices: [choice('none', 'marker.none', 'None'), ...typeChoices()],
		},
	];
	if (!trend) return fields;
	const number = (key: string, fallback: string, value: number, min: number, max: number) =>
		chartFormatNumberField(`${prefix}${key}`, `trendline.${key}`, fallback, value, min, max);
	if (trend.trendlineType === 'poly')
		fields.push(number('order', 'Polynomial order', trend.order ?? 2, 2, 6));
	if (trend.trendlineType === 'movingAvg')
		fields.push(number('period', 'Period', trend.period ?? 2, 2, 100000));
	fields.push(
		number('forward', 'Forecast forward', trend.forward ?? 0, 0, 100000),
		number('backward', 'Forecast backward', trend.backward ?? 0, 0, 100000),
	);
	if (trend.trendlineType !== 'movingAvg')
		fields.push(
			{
				id: `${prefix}equation`,
				kind: 'checkbox',
				label: t('trendline.equation', 'Display equation'),
				value: trend.displayEquation ?? false,
			},
			{
				id: `${prefix}rSquared`,
				kind: 'checkbox',
				label: t('trendline.rSquared', 'Display R-squared'),
				value: trend.displayRSquared ?? false,
			},
		);
	return fields;
}

/** One section per trendline (the first titled Trendlines), then Add trendline. */
export function chartTrendlineSections(trends: readonly ChartTrendline[]): ChartFormatSection[] {
	const sections: ChartFormatSection[] = [
		{
			id: 'trendlines',
			title: t('section.trendlines', 'Trendlines'),
			fields: trendlineFields('trendline.', trends[0]),
		},
	];
	trends.slice(1).forEach((trend, i) =>
		sections.push({
			id: `trendline-${i + 1}`,
			title: {
				key: 'chartFormat.section.trendlineN',
				fallback: 'Trendline {{n}}',
				params: { n: i + 2 },
			},
			fields: trendlineFields(`trendline.${i + 1}.`, trend),
		}),
	);
	if (trends.length > 0)
		sections[sections.length - 1]!.fields.push({
			id: 'trendline.add',
			kind: 'select',
			label: t('trendline.add', 'Add trendline'),
			value: 'none',
			choices: [choice('none', 'marker.none', 'None'), ...typeChoices()],
		});
	return sections;
}

/** The intent of a `trendline.*` field; null for a value or trendline the series does not take. */
export function chartTrendlineFieldIntent(
	trends: readonly ChartTrendline[],
	seriesIndex: number,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null {
	const text = String(value);
	const isType = CHART_TRENDLINE_TYPES.some(([type]) => type === text);
	if (field === 'trendline.add')
		return isType
			? {
					op: 'trendline',
					seriesIndex,
					index: trends.length,
					edit: { type: text as 'linear' },
				}
			: null;
	const match = /^trendline\.(?:(\d+)\.)?(\w+)$/u.exec(field);
	if (!match) return null;
	const index = match[1] === undefined ? 0 : Number(match[1]);
	const key = match[2]!;
	const wrap = (edit: Extract<ChartEditIntent, { op: 'trendline' }>['edit']): ChartEditIntent => ({
		op: 'trendline',
		seriesIndex,
		index,
		edit,
	});
	const trend = trends[index];
	if (key === 'type') {
		if (text === 'none') return trend ? wrap(null) : null;
		// A missing trendline is only added in place of the first (the None choice).
		return isType && (trend || index === trends.length) ? wrap({ type: text as 'linear' }) : null;
	}
	if (!trend) return null;
	if (key === 'equation') return wrap({ displayEquation: Boolean(value) });
	if (key === 'rSquared') return wrap({ displayRSquared: Boolean(value) });
	const n = value === '' || typeof value === 'boolean' ? NaN : Number(value);
	if (!['order', 'period', 'forward', 'backward'].includes(key) || !Number.isFinite(n) || n < 0)
		return null;
	return wrap({ [key]: n });
}
