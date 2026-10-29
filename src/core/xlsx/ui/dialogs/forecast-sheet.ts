// Data > Forecast Sheet: builds the forecast on a new sheet from a timeline and a values range.
// The sheet has the timeline, the values, and for the forecast rows FORECAST.ETS with lower and
// upper confidence bounds (FORECAST.ETS.CONFINT), optionally a statistics table, and a line or
// column chart over them. The whole job is one undo step.
import {
	type CellRange,
	type CellValue,
	type EditSession,
	getCell,
	isDateFormat,
	normalizeRange,
	quoteSheetName,
	styleAt,
	type Worksheet,
} from '../../index';
import { timelineGrid } from '../../formula/functions/ets-grid';
import { buildChart } from './insert-chart-model';

/** Aggregation codes of the FORECAST.ETS functions. */
export const FORECAST_AGGREGATIONS = [
	[1, 'Average'],
	[2, 'Count'],
	[3, 'CountA'],
	[4, 'Max'],
	[5, 'Min'],
	[6, 'Median'],
	[7, 'Sum'],
] as const;

export interface ForecastSheetRequest {
	/** The sheet holding the data. */
	sheet: number;
	/** One column or row of timeline values (numbers or dates), with or without a header cell. */
	timeline: CellRange;
	/** The values, the same shape as the timeline. */
	values: CellRange;
	/** The last timeline value of the forecast (a date serial or number). */
	end: number;
	/** Confidence level of the interval (default 0.95); `null` leaves the bounds out. */
	confidence?: number | null;
	/** 'auto' detects the season length; a number sets it (0 turns seasonality off). */
	seasonality?: 'auto' | number;
	/** How missing points are filled (default interpolation). */
	fill?: 'interpolate' | 'zeros';
	/** The aggregation code of duplicate timestamps (default 1, average). */
	aggregate?: number;
	/** Adds the table of fit statistics (alpha, beta, gamma, MASE, SMAPE, MAE, RMSE). */
	statistics?: boolean;
	/** The chart to draw (default a line chart); `null` draws none. */
	chart?: 'line' | 'column' | null;
	/** Name of the new sheet (default the next free one). */
	name?: string;
}

export type ForecastProblem =
	| 'shape'
	| 'not-numeric'
	| 'few-points'
	| 'uneven'
	| 'end-before-last'
	| 'too-long'
	| 'confidence'
	| 'seasonality'
	| 'aggregate';

const MESSAGES: Record<ForecastProblem, string> = {
	shape: 'The timeline and the values must each be one column or row of the same size.',
	'not-numeric': 'The timeline must hold numbers or dates and the values numbers.',
	'few-points': 'The data needs at least two different timeline points.',
	uneven: 'The timeline must have a constant step between its points.',
	'end-before-last': 'The forecast end must come after the last timeline point.',
	'too-long': 'The forecast would be longer than 10000 points.',
	confidence: 'The confidence level must be between 0 and 100 percent.',
	seasonality: 'The seasonality must be a whole number from 0 to 8760.',
	aggregate: 'Choose one of the aggregation functions.',
};

export const forecastMessage = (problem: ForecastProblem): string => MESSAGES[problem];

export class ForecastError extends Error {
	constructor(readonly problem: ForecastProblem) {
		super(MESSAGES[problem]);
	}
}

export interface ForecastSheetResult {
	/** Index of the new sheet. */
	sheet: number;
	historyRows: number;
	forecastRows: number;
}

const MAX_FORECAST_ROWS = 10_000;
const STATISTICS = ['Alpha', 'Beta', 'Gamma', 'MASE', 'SMAPE', 'MAE', 'RMSE'] as const;

interface Source {
	times: number[];
	values: (number | null)[];
	header?: { timeline: string; values: string };
	dateFormat?: string;
}

function cellsOf(ws: Worksheet, range: CellRange): CellValue[] {
	const r = normalizeRange(range);
	const out: CellValue[] = [];
	for (let row = r.start.row; row <= r.end.row; row++)
		for (let col = r.start.col; col <= r.end.col; col++)
			out.push(getCell(ws, row, col)?.value ?? null);
	return out;
}

const isVector = (r: CellRange): boolean => r.start.row === r.end.row || r.start.col === r.end.col;

function readSource(session: EditSession, request: ForecastSheetRequest): Source {
	const ws = session.workbook.sheets[request.sheet];
	if (!ws) throw new RangeError(`No sheet at index ${request.sheet}`);
	const t = normalizeRange(request.timeline);
	const v = normalizeRange(request.values);
	const size = (r: CellRange): number =>
		Math.max(r.end.row - r.start.row, r.end.col - r.start.col) + 1;
	if (!isVector(t) || !isVector(v) || size(t) !== size(v)) throw new ForecastError('shape');
	let times = cellsOf(ws, t);
	let values = cellsOf(ws, v);
	const header =
		typeof values[0] === 'string' && (times[0] === null || typeof times[0] === 'string')
			? { timeline: String(times[0] ?? ''), values: String(values[0]) }
			: undefined;
	if (header) {
		times = times.slice(1);
		values = values.slice(1);
	}
	const nums: number[] = [];
	const vals: (number | null)[] = [];
	for (let i = 0; i < times.length; i++) {
		const time = times[i];
		const value = values[i] ?? null;
		if (time === null && value === null) continue;
		if (typeof time !== 'number' || (value !== null && typeof value !== 'number'))
			throw new ForecastError('not-numeric');
		nums.push(time);
		vals.push(value);
	}
	const skip = header ? 1 : 0;
	const vertical = t.start.col === t.end.col;
	const first = getCell(
		ws,
		t.start.row + (vertical ? skip : 0),
		t.start.col + (vertical ? 0 : skip),
	);
	const format = styleAt(session.workbook, first?.styleId).numFmt;
	const source: Source = { times: nums, values: vals };
	if (header) source.header = header;
	if (format && isDateFormat(format)) source.dateFormat = format;
	return source;
}

/**
 * A default forecast end for the data: a quarter of the history beyond its last point, at least
 * one step. `date` says whether the timeline is formatted as dates. Undefined when the data
 * cannot be forecast (the dialog then shows the problem on OK).
 */
export function suggestForecastEnd(
	session: EditSession,
	request: Pick<ForecastSheetRequest, 'sheet' | 'timeline' | 'values'>,
): { end: number; date: boolean } | undefined {
	try {
		const source = readSource(session, request as ForecastSheetRequest);
		const distinct = [...new Set(source.times)].sort((a, b) => a - b);
		if (distinct.length < 2) return undefined;
		const grid = timelineGrid(distinct);
		const last = grid.index[grid.index.length - 1] as number;
		const end = grid.timeAt(last + Math.max(1, Math.round((last + 1) / 4)));
		return Number.isFinite(end) ? { end, date: source.dateFormat !== undefined } : undefined;
	} catch {
		return undefined;
	}
}

const abs = (col: string, from: number, to: number): string => `$${col}$${from}:$${col}$${to}`;

/** Builds the forecast sheet; throws {@link ForecastError} for a request it cannot honour. */
export function createForecastSheet(
	session: EditSession,
	request: ForecastSheetRequest,
): ForecastSheetResult {
	const confidence = request.confidence === undefined ? 0.95 : request.confidence;
	if (confidence !== null && !(confidence > 0 && confidence < 1))
		throw new ForecastError('confidence');
	const seasonality =
		request.seasonality === undefined || request.seasonality === 'auto' ? 1 : request.seasonality;
	if (!Number.isInteger(seasonality) || seasonality < 0 || seasonality > 8760)
		throw new ForecastError('seasonality');
	const aggregate = request.aggregate ?? 1;
	if (!FORECAST_AGGREGATIONS.some(([code]) => code === aggregate))
		throw new ForecastError('aggregate');
	const completion = request.fill === 'zeros' ? 0 : 1;

	const source = readSource(session, request);
	const order = source.times
		.map((_t, i) => i)
		.sort((a, b) => (source.times[a] as number) - (source.times[b] as number));
	const distinct = [...new Set(order.map((i) => source.times[i] as number))];
	if (distinct.length < 2) throw new ForecastError('few-points');
	let grid;
	try {
		grid = timelineGrid(distinct);
	} catch {
		throw new ForecastError('uneven');
	}
	const lastIndex = grid.index[grid.index.length - 1] as number;
	const endPosition = Math.round(grid.position(request.end));
	if (!(request.end > (distinct[distinct.length - 1] as number)) || endPosition <= lastIndex)
		throw new ForecastError('end-before-last');
	const forecastRows = endPosition - lastIndex;
	if (forecastRows > MAX_FORECAST_ROWS) throw new ForecastError('too-long');

	const history = order.length;
	const lastHistory = history + 1; // 1-based sheet row of the last historical point
	const lastRow = lastHistory + forecastRows;
	const withBounds = confidence !== null;
	const timelineRange = abs('A', 2, lastHistory);
	const valuesRange = abs('B', 2, lastHistory);
	const options = `${seasonality},${completion},${aggregate}`;

	let sheetIndex = -1;
	session.batch('Forecast sheet', () => {
		sheetIndex = session.addSheet(request.name, request.sheet + 1);
		const s = sheetIndex;
		const headers = ['Timeline', 'Values', 'Forecast'];
		if (withBounds) headers.push('Lower Confidence Bound', 'Upper Confidence Bound');
		session.setRangeValues(s, { row: 0, col: 0 }, [headers]);
		session.setRangeValues(
			s,
			{ row: 1, col: 0 },
			order.map((i) => [source.times[i] as number, source.values[i] ?? null]),
		);
		// The forecast line starts from the last actual value so the chart lines connect.
		session.setCellInput(s, history, 2, `=B${lastHistory}`);
		if (withBounds) {
			session.setCellInput(s, history, 3, `=B${lastHistory}`);
			session.setCellInput(s, history, 4, `=B${lastHistory}`);
		}
		for (let k = 1; k <= forecastRows; k++) {
			const row = lastHistory + k; // 1-based
			const r0 = row - 1;
			session.setCellValue(s, r0, 0, grid.timeAt(lastIndex + k));
			const target = `A${row}`;
			session.setCellInput(
				s,
				r0,
				2,
				`=FORECAST.ETS(${target},${valuesRange},${timelineRange},${options})`,
			);
			if (withBounds) {
				const interval = `FORECAST.ETS.CONFINT(${target},${valuesRange},${timelineRange},${confidence},${options})`;
				session.setCellInput(s, r0, 3, `=C${row}-${interval}`);
				session.setCellInput(s, r0, 4, `=C${row}+${interval}`);
			}
		}
		if (source.dateFormat)
			session.applyStyle(s, [{ start: { row: 1, col: 0 }, end: { row: lastRow - 1, col: 0 } }], {
				numFmt: source.dateFormat,
			});
		session.setColumnWidth(s, [0], 14);
		session.setColumnWidth(s, [1, 2], 12);
		if (withBounds) session.setColumnWidth(s, [3, 4], 24);
		if (request.statistics) {
			const col = withBounds ? 6 : 4;
			session.setRangeValues(s, { row: 0, col }, [['Statistic', 'Value']]);
			STATISTICS.forEach((name, i) => {
				session.setCellValue(s, 1 + i, col, name);
				session.setCellInput(
					s,
					1 + i,
					col + 1,
					`=FORECAST.ETS.STAT(${valuesRange},${timelineRange},${i + 1},${options})`,
				);
			});
			session.setColumnWidth(s, [col], 14);
		}
		if (request.chart !== null) {
			const lastCol = withBounds ? 4 : 2;
			const chart = buildChart(
				session.workbook,
				s,
				{ start: { row: 0, col: 1 }, end: { row: lastRow - 1, col: lastCol } },
				request.chart ?? 'line',
			);
			const prefix = `${quoteSheetName(session.workbook.sheets[s]?.name ?? '')}!`;
			const sheet = session.workbook.sheets[s] as Worksheet;
			const labels: CellValue[] = [];
			for (let r = 1; r < lastRow; r++) labels.push(getCell(sheet, r, 0)?.value ?? null);
			for (const series of chart.series) {
				series.categoriesRef = `${prefix}${abs('A', 2, lastRow)}`;
				series.categories = labels.map((v) => (typeof v === 'number' ? v : ''));
			}
			chart.anchor = {
				from: {
					row: 0,
					col: withBounds ? (request.statistics ? 9 : 6) : request.statistics ? 7 : 4,
					rowOffset: 0,
					colOffset: 0,
				},
				to: {
					row: 18,
					col: withBounds ? (request.statistics ? 17 : 14) : request.statistics ? 15 : 12,
					rowOffset: 0,
					colOffset: 0,
				},
			};
			chart.title = source.header?.values || 'Forecast';
			const { kind: _kind, ...model } = chart;
			session.addChart(s, model);
		}
	});
	return { sheet: sheetIndex, historyRows: history, forecastRows };
}
