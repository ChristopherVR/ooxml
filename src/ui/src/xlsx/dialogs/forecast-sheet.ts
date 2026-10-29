// Data > Forecast Sheet: the options of Excel's Create Forecast Worksheet dialog (timeline and
// values ranges, forecast end, confidence interval, seasonality, missing points, duplicates,
// statistics, chart). OK builds the forecast on a new sheet through the core's createForecastSheet.
import {
	dateToSerial,
	type CellRange,
	normalizeRange,
	parseRange,
	serialToDate,
} from 'ooxml-core/xlsx';
import {
	createForecastSheet,
	FORECAST_AGGREGATIONS,
	ForecastError,
	forecastMessage,
	regionOf,
	suggestForecastEnd,
	target,
} from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, fieldset, numberInput, radios, row, select, textInput } from './fields';
import { showDialog } from './frame';

/** A range typed in the dialog (`$A$1:$A$10`, optionally with a sheet prefix). */
function readRange(text: string): CellRange | undefined {
	const ref = text.trim().replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '');
	const range = ref ? parseRange(ref) : undefined;
	return range ? normalizeRange(range) : undefined;
}

const addressOf = (r: CellRange): string => {
	const a = (row: number, col: number) => {
		let letters = '';
		for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26))
			letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
		return `$${letters}$${row + 1}`;
	};
	const start = a(r.start.row, r.start.col);
	const end = a(r.end.row, r.end.col);
	return start === end ? start : `${start}:${end}`;
};

/** The timeline and values columns the selection suggests: first column and the one after it. */
function suggestedRanges(range: CellRange): { timeline: CellRange; values: CellRange } | undefined {
	if (range.end.col <= range.start.col) return undefined;
	return {
		timeline: { start: range.start, end: { row: range.end.row, col: range.start.col } },
		values: {
			start: { row: range.start.row, col: range.start.col + 1 },
			end: { row: range.end.row, col: range.start.col + 1 },
		},
	};
}

/** The forecast end as the user typed it: a number, or a date (ISO or as the browser parses it). */
function parseEnd(text: string, date1904: boolean): number {
	const trimmed = text.trim();
	if (trimmed === '') return Number.NaN;
	const number = Number(trimmed);
	if (Number.isFinite(number)) return number;
	const parsed = Date.parse(/^\d{4}-\d{2}-\d{2}$/.test(trimmed) ? `${trimmed}T00:00:00Z` : trimmed);
	return Number.isNaN(parsed) ? Number.NaN : Math.floor(dateToSerial(new Date(parsed), date1904));
}

export function openForecastSheet(ctx: EditorContext): Promise<number | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const guess = suggestedRanges(regionOf(t));
	const timeline = textInput(ctx, guess ? addressOf(guess.timeline) : '');
	const values = textInput(ctx, guess ? addressOf(guess.values) : '');
	const date1904 = Boolean(t.workbook.date1904);
	const suggestion = guess ? suggestForecastEnd(t.session, { sheet: t.sheet, ...guess }) : undefined;
	const endText = suggestion
		? suggestion.date
			? serialToDate(suggestion.end, date1904).toISOString().slice(0, 10)
			: String(suggestion.end)
		: '';
	const end = textInput(ctx, endText);
	const interval = checkbox(ctx, 'Confidence interval', true);
	const level = numberInput(ctx, 95, 1, 99);
	level.setAttribute('aria-label', ctx.t('Confidence level (%)'));
	const seasonMode = radios(
		ctx,
		'Seasonality',
		[
			['auto', 'Detect automatically'],
			['manual', 'Set manually'],
		],
		'auto',
	);
	const season = numberInput(ctx, 1, 0, 8760);
	season.setAttribute('aria-label', ctx.t('Season length'));
	const fill = select(
		ctx,
		[
			['interpolate', 'Interpolation'],
			['zeros', 'Zeros'],
		],
		'interpolate',
	);
	const aggregate = select(
		ctx,
		FORECAST_AGGREGATIONS.map(([code, name]) => [String(code), name] as const),
		'1',
	);
	const statistics = checkbox(ctx, 'Include forecast statistics', false);
	const chart = select(
		ctx,
		[
			['line', 'Line chart'],
			['column', 'Column chart'],
		],
		'line',
	);
	const syncEnabled = (): void => {
		level.disabled = !interval.input.checked;
		season.disabled = seasonMode.get() !== 'manual';
	};
	interval.input.addEventListener('change', syncEnabled);
	for (const input of seasonMode.inputs) input.addEventListener('change', syncEnabled);
	syncEnabled();
	return showDialog<number>(ctx, {
		name: 'forecast-sheet',
		heading: 'Create Forecast Worksheet',
		body: [
			field(ctx, 'Timeline range:', timeline),
			field(ctx, 'Values range:', values),
			field(ctx, 'Forecast end:', end),
			row(ctx, interval.wrapper, level),
			fieldset(ctx, 'Seasonality', seasonMode.element, field(ctx, 'Season length', season)),
			field(ctx, 'Fill missing points using:', fill),
			field(ctx, 'Aggregate duplicates using:', aggregate),
			statistics.wrapper,
			field(ctx, 'Chart type', chart),
		],
		opened: () => end.focus(),
		submit: () => {
			const timelineRange = readRange(timeline.value);
			const valuesRange = readRange(values.value);
			if (!timelineRange || !valuesRange) {
				ctx.toast(ctx.t('Reference is not valid.'), 'warning');
				(timelineRange ? values : timeline).focus();
				return undefined;
			}
			const endValue = parseEnd(end.value, date1904);
			if (Number.isNaN(endValue)) {
				ctx.toast(ctx.t('Enter the forecast end as a date or a number.'), 'warning');
				end.focus();
				return undefined;
			}
			try {
				const result = createForecastSheet(t.session, {
					sheet: t.sheet,
					timeline: timelineRange,
					values: valuesRange,
					end: endValue,
					confidence: interval.input.checked ? Number(level.value) / 100 : null,
					seasonality: seasonMode.get() === 'manual' ? Number(season.value) : 'auto',
					fill: fill.value === 'zeros' ? 'zeros' : 'interpolate',
					aggregate: Number(aggregate.value),
					statistics: statistics.input.checked,
					chart: chart.value === 'column' ? 'column' : 'line',
				});
				return result.sheet;
			} catch (error) {
				if (!(error instanceof ForecastError)) throw error;
				ctx.toast(ctx.t(forecastMessage(error.problem)), 'warning');
				return undefined;
			}
		},
	});
}
