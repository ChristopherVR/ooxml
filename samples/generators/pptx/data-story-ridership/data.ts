/**
 * The one invented dataset behind "How Port Averly moved". Every chart, table, headline and
 * speaker note reads from here, so the numbers agree across the deck by construction.
 * Boardings are in millions unless a name says otherwise.
 */
import { round, seeded } from '../../lib/random';

export const MONTHS = [
	'Jan',
	'Feb',
	'Mar',
	'Apr',
	'May',
	'Jun',
	'Jul',
	'Aug',
	'Sep',
	'Oct',
	'Nov',
	'Dec',
] as const;

export const MONTHLY_2025 = [
	3.61, 3.58, 3.94, 3.97, 4.12, 4.06, 3.82, 3.79, 4.19, 4.58, 4.52, 4.42,
];
export const MONTHLY_2019 = [3.78, 3.71, 4.02, 3.95, 4.1, 3.98, 3.66, 3.6, 4.12, 4.21, 4.05, 3.72];

export function sum(values: readonly number[]): number {
	return round(
		values.reduce((total, value) => total + value, 0),
		2,
	);
}

export const TOTAL_2025 = sum(MONTHLY_2025);
export const TOTAL_2019 = sum(MONTHLY_2019);
export const TOTAL_2024 = 44.1;

export const MODES = ['Bus', 'Tram', 'Ferry', 'On-demand'] as const;
export const MODE_2025 = [27.9, 14.1, 4.3, 2.3];
export const MODE_2024 = [26.4, 12.0, 3.9, 1.8];

/** Percent change, one decimal. */
export function change(now: number, before: number): number {
	return round(((now - before) / before) * 100, 1);
}

/** A signed percentage as the deck prints it (`+10.2%`). */
export function signed(value: number, digits = 1): string {
	return `${value >= 0 ? '+' : '-'}${Math.abs(value).toFixed(digits)}%`;
}

/** October 2025 against the best month of 2019 (also an October). */
export const PEAK_2019 = Math.max(...MONTHLY_2019);
export const RECORD_MONTH = Math.max(...MONTHLY_2025);
/** Months of 2025 that beat the same month of 2019. */
export const MONTHS_AHEAD = MONTHLY_2025.filter(
	(value, index) => value > MONTHLY_2019[index]!,
).length;

export const DISTRICTS = ['Harbourside', 'Old Quay', 'Northgate', 'Marrow Hill', 'Saltmeadow'];
/** Boardings by district (rows) and mode (columns, as `MODES`). */
export const DISTRICT_MODE = [
	[5.0, 5.6, 2.6, 0.2],
	[5.6, 4.2, 1.2, 0.2],
	[6.4, 2.9, 0, 0.5],
	[6.2, 1.4, 0, 0.5],
	[4.7, 0, 0.5, 0.9],
];

export const FARES = [
	{ name: 'Capped tap-and-go', share: 38 },
	{ name: 'Monthly pass', share: 27 },
	{ name: 'Stored-value card', share: 19 },
	{ name: 'Single ticket', share: 9 },
	{ name: 'Concession pass', share: 7 },
];
/** Single tickets a year earlier, for the comparison the fare slide draws. */
export const SINGLE_TICKET_2024 = 21;

export const QUARTERS = ['Q1 24', 'Q2 24', 'Q3 24', 'Q4 24', 'Q1 25', 'Q2 25', 'Q3 25', 'Q4 25'];
const QUARTERLY_2024 = [10.3, 11.2, 10.9, 11.7];
const QUARTERLY_2025 = [0, 1, 2, 3].map((quarter) =>
	sum(MONTHLY_2025.slice(quarter * 3, quarter * 3 + 3)),
);
export const QUARTERLY_BOARDINGS = [...QUARTERLY_2024, ...QUARTERLY_2025];
/** Revenue service hours, in thousands. */
export const SERVICE_HOURS = [365, 368, 370, 371, 372, 378, 381, 396];
/** Boardings per revenue service hour. */
export const PRODUCTIVITY = QUARTERLY_BOARDINGS.map((boardings, index) =>
	round((boardings * 1000) / SERVICE_HOURS[index]!, 1),
);
export const HOURS_2024 = sum(SERVICE_HOURS.slice(0, 4));
export const HOURS_2025 = sum(SERVICE_HOURS.slice(4));
export const PRODUCTIVITY_2024 = round((TOTAL_2024 * 1000) / HOURS_2024, 1);
export const PRODUCTIVITY_2025 = round((TOTAL_2025 * 1000) / HOURS_2025, 1);

export interface RouteRow {
	route: string;
	name: string;
	mode: string;
	boardings: number;
	before: number;
	headway: number;
}
export const TOP_ROUTES: RouteRow[] = [
	{ route: 'T1', name: 'Harbour Line', mode: 'Tram', boardings: 7.9, before: 6.2, headway: 5 },
	{ route: 'T2', name: 'Northgate Line', mode: 'Tram', boardings: 6.2, before: 5.8, headway: 6 },
	{
		route: '4',
		name: 'Old Quay to Marrow Hill',
		mode: 'Bus',
		boardings: 3.4,
		before: 2.9,
		headway: 8,
	},
	{
		route: '9',
		name: 'Saltmeadow Crosstown',
		mode: 'Bus',
		boardings: 2.8,
		before: 2.6,
		headway: 10,
	},
	{ route: 'F1', name: 'Harbour Ferry', mode: 'Ferry', boardings: 2.7, before: 2.4, headway: 15 },
];
export const TOP_SHARE = round((sum(TOP_ROUTES.map((row) => row.boardings)) / TOTAL_2025) * 100, 0);

/** The week the Harbour Line extension opened (1 September 2025). */
export const EXTENSION_WEEK = 36;

/**
 * Weekly boardings by mode, in thousands. The weekly totals follow the monthly totals day by
 * day; the split between modes is fitted so that each mode's year adds up to `MODE_2025`.
 */
function weeklyByMode(): number[][] {
	const days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
	const daily: number[] = [];
	days.forEach((count, month) => {
		for (let day = 0; day < count; day++) daily.push((MONTHLY_2025[month]! * 1000) / count);
	});
	const random = seeded(2025);
	let totals = Array.from({ length: 52 }, (_, week) => {
		const base = daily.slice(week * 7, week * 7 + 7).reduce((total, value) => total + value, 0);
		return base * (1 + random.range(-0.015, 0.015));
	});
	const scale = (TOTAL_2025 * 1000) / totals.reduce((total, value) => total + value, 0);
	totals = totals.map((value) => value * scale);
	let cells = totals.map((_, index) => {
		const week = index + 1;
		const summer = Math.exp(-((week - 29) ** 2) / (2 * 6 ** 2));
		return [
			0.6,
			week < EXTENSION_WEEK ? 0.262 : 0.345,
			0.075 * (1 + 0.5 * summer),
			0.038 + 0.014 * (week / 52),
		];
	});
	// Iterative proportional fitting: rows to the weekly totals, columns to the mode totals.
	for (let pass = 0; pass < 40; pass++) {
		cells = cells.map((row, week) => {
			const total = row.reduce((a, b) => a + b, 0);
			return row.map((value) => (value * totals[week]!) / total);
		});
		const columns = MODES.map((_, mode) => cells.reduce((total, row) => total + row[mode]!, 0));
		cells = cells.map((row) =>
			row.map((value, mode) => (value * MODE_2025[mode]! * 1000) / columns[mode]!),
		);
	}
	return MODES.map((_, mode) => cells.map((row) => round(row[mode]!, 1)));
}

export const WEEKLY = weeklyByMode();
export const WEEKS = Array.from({ length: 52 }, (_, index) => `W${index + 1}`);

function mean(values: readonly number[]): number {
	return values.reduce((total, value) => total + value, 0) / values.length;
}
/** Average weekly tram boardings (thousands) before and after the extension opened. */
export const TRAM_BEFORE = Math.round(mean(WEEKLY[1]!.slice(0, EXTENSION_WEEK - 1)));
export const TRAM_AFTER = Math.round(mean(WEEKLY[1]!.slice(EXTENSION_WEEK - 1)));

export interface RoutePoint {
	/** Vehicles per hour in the peak. */
	frequency: number;
	/** Boardings per revenue service hour. */
	productivity: number;
}
function busRoutes(count: number, frequencies: readonly number[], seed: number): RoutePoint[] {
	const random = seeded(seed);
	return Array.from({ length: count }, (_, index) => {
		const frequency = frequencies[index % frequencies.length]!;
		return { frequency, productivity: round(6 + 4.5 * frequency + random.normal(0, 2.6), 1) };
	});
}
/** The six routes upgraded to run every 8 minutes or better, and the other 18 bus routes. */
export const FREQUENT_ROUTES = busRoutes(6, [7.5, 8, 10, 12, 8.5, 10], 61);
export const OTHER_ROUTES = busRoutes(
	18,
	[2, 3, 4, 6, 3, 5, 2, 4, 6, 3, 4, 5, 2, 3, 6, 4, 5, 3],
	62,
);
export const FREQUENT_MEAN = round(mean(FREQUENT_ROUTES.map((point) => point.productivity)), 1);
export const OTHER_MEAN = round(mean(OTHER_ROUTES.map((point) => point.productivity)), 1);
export const FREQUENCY_RATIO = round(FREQUENT_MEAN / OTHER_MEAN, 1);

/** The target the 2026 plan is built on. */
export const TARGET_2026 = 52;

/** A count of thousands written out in full (`1474` becomes `1,474,000`). */
export function thousands(value: number): string {
	const whole = String(Math.round(value));
	return `${whole.replace(/\B(?=(\d{3})+$)/gu, ',')},000`;
}
