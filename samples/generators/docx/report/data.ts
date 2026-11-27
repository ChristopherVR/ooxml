/**
 * The invented measurements behind the Greywharf report. Every number the text, the tables and the
 * charts show is derived from these arrays, so they cannot disagree.
 */
import { round } from '../../lib/random';

export const PALETTE = {
	ink: '1B2A33',
	deep: '0F4C5C',
	teal: '1F7A8C',
	sea: '5FA8B8',
	mist: 'E3EEF0',
	foam: 'F3F8F9',
	amber: 'E09F3E',
	rust: '9E2A2B',
	slate: '5B6B73',
	rule: 'C5D3D8',
	white: 'FFFFFF',
} as const;

export const REPORT = {
	title: 'Greywharf tidal energy pilot',
	subtitle: 'First-year technical report',
	number: 'SICE-TR-2026-01',
	institute: 'Saltmarsh Institute for Coastal Engineering',
	date: '15 January 2026',
	period: '1 November 2024 to 31 October 2025',
	/** Turbines in the array and the rating of each, kW. */
	turbines: 4,
	ratedKw: 500,
} as const;

export const AUTHORS = [
	{ name: 'Mara Ellingwood', role: 'Principal investigator' },
	{ name: 'Tobias Renner', role: 'Resource modelling' },
	{ name: 'Priya Vasanth', role: 'Environmental monitoring' },
	{ name: 'Idris Okonjo', role: 'Operations and maintenance' },
] as const;

export const REVIEWERS = {
	leona: { author: 'Leona Castellanos', initials: 'LC' },
	hugo: { author: 'Hugo Falkenrath', initials: 'HF' },
} as const;

/**
 * One row a month: label, long name, days, modelled and measured net energy (MWh), turbine-hours
 * lost to outages, mean and peak depth-averaged flow speed at the array (m/s), and a note.
 */
type MonthRow = [string, string, number, number, number, number, number, number, string];
// prettier-ignore
const ROWS: MonthRow[] = [
	['Nov', 'November 2024', 30, 492, 318, 829, 1.84, 3.31, 'Commissioning: turbines energised in turn'],
	['Dec', 'December 2024', 31, 508, 402, 458, 1.9, 3.38, 'Converter firmware update on T2 and T4'],
	['Jan', 'January 2025', 31, 511, 446, 259, 1.93, 3.41, 'Storm stand-down, 9 to 11 January'],
	['Feb', 'February 2025', 28, 458, 421, 140, 1.91, 3.36, 'First inspection by remote vehicle'],
	['Mar', 'March 2025', 31, 521, 489, 116, 1.96, 3.47, 'Equinoctial springs; highest flow'],
	['Apr', 'April 2025', 30, 498, 472, 89, 1.92, 3.4, 'Normal operation'],
	['May', 'May 2025', 31, 496, 361, 702, 1.88, 3.29, 'T3 recovered for a gearbox seal'],
	['Jun', 'June 2025', 30, 472, 448, 107, 1.83, 3.22, 'T3 returned to service on 2 June'],
	['Jul', 'July 2025', 31, 489, 471, 83, 1.86, 3.27, 'Normal operation'],
	['Aug', 'August 2025', 31, 503, 486, 74, 1.91, 3.35, 'Blade strain survey completed'],
	['Sep', 'September 2025', 30, 507, 492, 63, 1.95, 3.44, 'Normal operation'],
	['Oct', 'October 2025', 31, 518, 505, 57, 1.97, 3.46, 'Grid curtailment test, 14 October'],
];
const INPUT = ROWS.map(
	([label, long, days, modelled, measured, outage, meanFlow, peakFlow, note]) => ({
		label,
		long,
		days,
		modelled,
		measured,
		outage,
		meanFlow,
		peakFlow,
		note,
	}),
);

const ratedMw = (REPORT.turbines * REPORT.ratedKw) / 1000;

export const MONTHS = INPUT.map((month) => {
	const turbineHours = month.days * 24 * REPORT.turbines;
	return {
		...month,
		/** Measured over modelled, percent. */
		ratio: round((100 * month.measured) / month.modelled, 1),
		/** Share of turbine-hours the turbines were able to generate, percent. */
		availability: round(100 * (1 - month.outage / turbineHours), 1),
		/** Net energy over the energy at rated power for the whole month, percent. */
		capacityFactor: round((100 * month.measured) / (ratedMw * month.days * 24), 1),
	};
});

const sum = (values: number[]) => values.reduce((total, value) => total + value, 0);
const days = sum(MONTHS.map((month) => month.days));

export const YEAR = {
	days,
	ratedMw,
	modelled: sum(MONTHS.map((month) => month.modelled)),
	measured: sum(MONTHS.map((month) => month.measured)),
	outage: sum(MONTHS.map((month) => month.outage)),
	get ratio() {
		return round((100 * this.measured) / this.modelled, 1);
	},
	get availability() {
		return round(100 * (1 - this.outage / (days * 24 * REPORT.turbines)), 1);
	},
	get capacityFactor() {
		return round((100 * this.measured) / (ratedMw * days * 24), 1);
	},
	get modelledCapacityFactor() {
		return round((100 * this.modelled) / (ratedMw * days * 24), 1);
	},
	meanFlow: round(sum(MONTHS.map((month) => month.meanFlow * month.days)) / days, 2),
	peakFlow: Math.max(...MONTHS.map((month) => month.peakFlow)),
};

/** The four quarters of the reporting year, three months each. */
export const QUARTERS = [0, 1, 2, 3].map((index) => {
	const months = MONTHS.slice(index * 3, index * 3 + 3);
	const modelled = sum(months.map((month) => month.modelled));
	const measured = sum(months.map((month) => month.measured));
	const hours = sum(months.map((month) => month.days)) * 24 * REPORT.turbines;
	return {
		label: `Q${index + 1}`,
		span: `${months[0]!.label} to ${months[2]!.label}`,
		modelled,
		measured,
		ratio: round((100 * measured) / modelled, 1),
		availability: round(100 * (1 - sum(months.map((month) => month.outage)) / hours), 1),
	};
});

/** The months after commissioning and outside the T3 recovery: the steady-state record. */
const steady = MONTHS.filter((month) => month.availability >= 94);
export const STEADY = {
	months: steady.length,
	ratio: round(
		(100 * sum(steady.map((month) => month.measured))) / sum(steady.map((month) => month.modelled)),
		1,
	),
	availability: round(
		100 *
			(1 -
				sum(steady.map((month) => month.outage)) /
					(sum(steady.map((month) => month.days)) * 24 * REPORT.turbines)),
		1,
	),
};

/** Lost turbine-hours by cause: the commissioning months, the May recovery of T3, and the rest. */
export const LOSSES = (() => {
	const commissioning = MONTHS[0]!.outage + MONTHS[1]!.outage;
	const seal = MONTHS[6]!.outage;
	const other = YEAR.outage - commissioning - seal;
	const share = (hours: number) => Math.round((100 * hours) / YEAR.outage);
	return {
		commissioning,
		seal,
		other,
		commissioningShare: share(commissioning),
		sealShare: share(seal),
		otherShare: 100 - share(commissioning) - share(seal),
	};
})();

/** Rotor and fluid constants of the kinetic power expression. */
export const ROTOR = { diameter: 16, density: 1025, powerCoefficient: 0.41, cutIn: 0.8 } as const;
const sweptArea = (Math.PI * ROTOR.diameter ** 2) / 4;
/** Hydrodynamic power of one turbine at flow speed `v`, kW, limited to its rating. */
export const modelledPower = (v: number): number =>
	v < ROTOR.cutIn
		? 0
		: Math.min(
				REPORT.ratedKw,
				(0.5 * ROTOR.density * sweptArea * ROTOR.powerCoefficient * v ** 3) / 1000,
			);
export const SWEPT_AREA = round(sweptArea, 0);
/** The flow speed at which the modelled curve reaches rated power, m/s. */
export const RATED_SPEED = round(
	((REPORT.ratedKw * 1000) / (0.5 * ROTOR.density * sweptArea * ROTOR.powerCoefficient)) ** (1 / 3),
	2,
);

/** Binned power curve of one turbine: flow speed bin centres and the mean power in each, kW. */
export const POWER_CURVE = Array.from({ length: 15 }, (_, index) => {
	const speed = round(0.6 + index * 0.2, 1);
	const modelled = modelledPower(speed);
	// Measured power sits a few percent under the model below rated speed and just under rating above it.
	const shortfall = modelled >= REPORT.ratedKw ? 0.012 : 0.07 + 0.025 * Math.sin(index * 1.3);
	return { speed, modelled: round(modelled, 0), measured: round(modelled * (1 - shortfall), 0) };
});

/** Levelised cost of energy over the design life, from the year-one cost record. */
export const COST = (() => {
	const capex = 13.6;
	const opex = 0.62;
	const rate = 0.065;
	const life = 20;
	const steadyEnergy = 5720;
	const annuity = (1 - (1 + rate) ** -life) / rate;
	const cost = capex + opex * annuity;
	const energy = steadyEnergy * annuity;
	return {
		capex,
		opex,
		rate,
		life,
		steadyEnergy,
		annuity: round(annuity, 2),
		presentCost: round(cost, 1),
		presentEnergy: Math.round(energy / 100) * 100,
		lcoe: Math.round((cost * 1e6) / energy),
	};
})();

/** Thousands separator for the document's numbers (a thin, locale-free comma grouping). */
export const n = (value: number, digits = 0): string => {
	const [whole, fraction] = value.toFixed(digits).split('.');
	const grouped = whole!.replace(/\B(?=(\d{3})+(?!\d))/gu, ',');
	return fraction ? `${grouped}.${fraction}` : grouped;
};
