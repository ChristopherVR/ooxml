import { ERR, fail } from '../values.js';
import { num, spec, str } from './helpers.js';
import type { FunctionSpec } from './types.js';

/**
 * CONVERT's unit table. It covers the common units of each dimension Excel documents but not
 * every one (the Pica area and volume units, for example, are absent); a unit not listed here
 * returns `#N/A`, as an unknown unit does in Excel.
 *
 * Each entry: `factor` converts one of the unit to the dimension's base unit, `power` is 2 or 3 for
 * area and volume units (a metric prefix is raised to it), and `prefix` says whether metric
 * prefixes may be attached (information units also take the binary ones).
 */
interface Unit {
	dimension: string;
	factor: number;
	power?: number;
	prefix?: 'metric' | 'binary';
}

const table = new Map<string, Unit>();
const add = (
	dimension: string,
	names: string[],
	factor: number,
	extra: Partial<Unit> = {},
): void => {
	for (const name of names) table.set(name, { dimension, factor, ...extra });
};

// Mass, base gram.
add('mass', ['g'], 1, { prefix: 'metric' });
add('mass', ['sg'], 14593.9029372064);
add('mass', ['lbm'], 453.59237);
add('mass', ['u'], 1.6605390666e-24, { prefix: 'metric' });
add('mass', ['ozm'], 28.349523125);
add('mass', ['grain'], 0.06479891);
add('mass', ['cwt', 'shweight'], 45359.237);
add('mass', ['uk_cwt', 'lcwt', 'hweight'], 50802.34544);
add('mass', ['stone'], 6350.29318);
add('mass', ['ton'], 907184.74);
add('mass', ['uk_ton', 'LTON', 'brton'], 1016046.9088);
// Distance, base metre.
add('distance', ['m'], 1, { prefix: 'metric' });
add('distance', ['mi'], 1609.344);
add('distance', ['Nmi'], 1852);
add('distance', ['in'], 0.0254);
add('distance', ['ft'], 0.3048);
add('distance', ['yd'], 0.9144);
add('distance', ['ang'], 1e-10, { prefix: 'metric' });
add('distance', ['ell'], 1.143);
add('distance', ['ly'], 9460730472580800);
add('distance', ['parsec', 'pc'], 30856775814913670, { prefix: 'metric' });
add('distance', ['Pica'], 0.0254 / 72);
add('distance', ['pica'], 0.0254 / 6);
add('distance', ['survey_mi'], 1609.3472186944);
// Time, base second.
add('time', ['yr'], 31557600);
add('time', ['day', 'd'], 86400);
add('time', ['hr'], 3600);
add('time', ['mn', 'min'], 60);
add('time', ['sec', 's'], 1, { prefix: 'metric' });
// Pressure, base pascal.
add('pressure', ['Pa', 'p'], 1, { prefix: 'metric' });
add('pressure', ['atm', 'at'], 101325, { prefix: 'metric' });
add('pressure', ['mmHg'], 133.322, { prefix: 'metric' });
add('pressure', ['psi'], 6894.757293168);
add('pressure', ['Torr'], 101325 / 760);
// Force, base newton.
add('force', ['N'], 1, { prefix: 'metric' });
add('force', ['dyn', 'dy'], 1e-5, { prefix: 'metric' });
add('force', ['lbf'], 4.4482216152605);
add('force', ['pond'], 0.00980665, { prefix: 'metric' });
// Energy, base joule.
add('energy', ['J'], 1, { prefix: 'metric' });
add('energy', ['e'], 1e-7, { prefix: 'metric' });
add('energy', ['c'], 4.184, { prefix: 'metric' });
add('energy', ['cal'], 4.1868, { prefix: 'metric' });
add('energy', ['eV', 'ev'], 1.602176487e-19, { prefix: 'metric' });
add('energy', ['HPh', 'hh'], 2684519.537696172);
add('energy', ['Wh', 'wh'], 3600, { prefix: 'metric' });
add('energy', ['flb'], 1.3558179483314);
add('energy', ['BTU', 'btu'], 1055.05585262);
// Power, base watt.
add('power', ['HP', 'h'], 745.69987158227);
add('power', ['PS'], 735.49875);
add('power', ['W', 'w'], 1, { prefix: 'metric' });
// Magnetism, base tesla.
add('magnetism', ['T'], 1, { prefix: 'metric' });
add('magnetism', ['ga'], 1e-4, { prefix: 'metric' });
// Volume, base cubic metre.
add('volume', ['tsp'], 4.92892159375e-6);
add('volume', ['tspm'], 5e-6);
add('volume', ['tbs'], 1.478676478125e-5);
add('volume', ['oz'], 2.95735295625e-5);
add('volume', ['cup'], 2.365882365e-4);
add('volume', ['pt', 'us_pt'], 4.73176473e-4);
add('volume', ['uk_pt'], 5.6826125e-4);
add('volume', ['qt'], 9.46352946e-4);
add('volume', ['uk_qt'], 1.1365225e-3);
add('volume', ['gal'], 3.785411784e-3);
add('volume', ['uk_gal'], 4.54609e-3);
add('volume', ['l', 'L', 'lt'], 1e-3, { prefix: 'metric' });
add('volume', ['m3', 'm^3'], 1, { prefix: 'metric', power: 3 });
add('volume', ['ang3', 'ang^3'], 1e-30, { prefix: 'metric', power: 3 });
add('volume', ['barrel'], 0.158987294928);
add('volume', ['bushel'], 0.03523907016688);
add('volume', ['regton', 'GRT'], 2.8316846592);
add('volume', ['MTON'], 1.13267386368);
add('volume', ['in3', 'in^3'], 1.6387064e-5);
add('volume', ['ft3', 'ft^3'], 0.028316846592);
add('volume', ['yd3', 'yd^3'], 0.764554857984);
add('volume', ['mi3', 'mi^3'], 4168181825.4405794);
add('volume', ['Nmi3', 'Nmi^3'], 6352182208);
add('volume', ['ly3', 'ly^3'], 8.46786664623715e47);
// Area, base square metre.
add('area', ['m2', 'm^2'], 1, { prefix: 'metric', power: 2 });
add('area', ['ang2', 'ang^2'], 1e-20, { prefix: 'metric', power: 2 });
add('area', ['ar'], 100, { prefix: 'metric' });
add('area', ['ha'], 10000, { prefix: 'metric' });
add('area', ['uk_acre'], 4046.8564224);
add('area', ['us_acre'], 4046.87260987425);
add('area', ['ft2', 'ft^2'], 0.09290304);
add('area', ['in2', 'in^2'], 6.4516e-4);
add('area', ['yd2', 'yd^2'], 0.83612736);
add('area', ['mi2', 'mi^2'], 2589988.110336);
add('area', ['Nmi2', 'Nmi^2'], 3429904);
add('area', ['ly2', 'ly^2'], 8.95054210748189e31);
add('area', ['Morgen'], 2500);
// Information, base bit.
add('information', ['bit'], 1, { prefix: 'binary' });
add('information', ['byte'], 8, { prefix: 'binary' });
// Speed, base metre per second.
add('speed', ['m/s', 'm/sec'], 1, { prefix: 'metric' });
add('speed', ['m/h', 'm/hr'], 1 / 3600, { prefix: 'metric' });
add('speed', ['mph'], 0.44704);
add('speed', ['kn'], 1852 / 3600);
add('speed', ['admkn'], 0.514773333333333);

const METRIC: Record<string, number> = {
	Y: 1e24,
	Z: 1e21,
	E: 1e18,
	P: 1e15,
	T: 1e12,
	G: 1e9,
	M: 1e6,
	k: 1e3,
	h: 1e2,
	e: 1e1,
	d: 1e-1,
	c: 1e-2,
	m: 1e-3,
	u: 1e-6,
	n: 1e-9,
	p: 1e-12,
	f: 1e-15,
	a: 1e-18,
	z: 1e-21,
	y: 1e-24,
};
const BINARY: Record<string, number> = {
	Yi: 2 ** 80,
	Zi: 2 ** 70,
	Ei: 2 ** 60,
	Pi: 2 ** 50,
	Ti: 2 ** 40,
	Gi: 2 ** 30,
	Mi: 2 ** 20,
	ki: 2 ** 10,
};

interface Resolved {
	dimension: string;
	/** Multiplies a value in this unit to give the dimension's base unit. */
	factor: number;
}

function resolve(name: string): Resolved {
	const direct = table.get(name);
	if (direct) return { dimension: direct.dimension, factor: direct.factor };
	for (const length of [2, 1]) {
		const unit = table.get(name.slice(length));
		const prefix = name.slice(0, length);
		if (!unit?.prefix) continue;
		const scale = (unit.prefix === 'binary' ? BINARY[prefix] : undefined) ?? METRIC[prefix];
		if (scale !== undefined)
			return { dimension: unit.dimension, factor: unit.factor * scale ** (unit.power ?? 1) };
	}
	return fail(ERR.NA);
}

const TEMPERATURES: Record<
	string,
	{ toKelvin: (v: number) => number; fromKelvin: (k: number) => number }
> = {
	C: { toKelvin: (v) => v + 273.15, fromKelvin: (k) => k - 273.15 },
	F: { toKelvin: (v) => ((v + 459.67) * 5) / 9, fromKelvin: (k) => (k * 9) / 5 - 459.67 },
	K: { toKelvin: (v) => v, fromKelvin: (k) => k },
	Rank: { toKelvin: (v) => (v * 5) / 9, fromKelvin: (k) => (k * 9) / 5 },
	Reau: { toKelvin: (v) => (v * 5) / 4 + 273.15, fromKelvin: (k) => ((k - 273.15) * 4) / 5 },
};
const TEMPERATURE_NAMES: Record<string, string> = { cel: 'C', fah: 'F', kel: 'K' };
const temperature = (name: string) => TEMPERATURES[TEMPERATURE_NAMES[name] ?? name];

export function convertUnits(value: number, from: string, to: string): number {
	const [tf, tt] = [temperature(from), temperature(to)];
	if (tf || tt) {
		if (!tf || !tt) fail(ERR.NA);
		return tt.fromKelvin(tf.toKelvin(value));
	}
	const [a, b] = [resolve(from), resolve(to)];
	if (a.dimension !== b.dimension) fail(ERR.NA);
	return (value * a.factor) / b.factor;
}

export const CONVERT_FUNCTIONS: FunctionSpec[] = [
	spec(
		'CONVERT',
		'Engineering',
		'CONVERT(number, from_unit, to_unit)',
		'Converts a number from one measurement system to another.',
		3,
		3,
		(args) => convertUnits(num(args[0]), str(args[1]), str(args[2])),
	),
];
