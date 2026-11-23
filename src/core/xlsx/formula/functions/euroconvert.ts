import { ERR, isError } from '../values';
import { bool, num, scalar, spec, str } from './helpers';
import type { FunctionSpec } from './types';

interface EuroCurrency {
	/** Fixed EU conversion rate: units of this currency per 1 EUR. */
	readonly rate: number;
	/** Default calculation precision (decimal places) when full_precision is FALSE. */
	readonly precision: number;
}

/**
 * Official EU Council Regulation fixed conversion rates (units per 1 EUR) and
 * Excel EUROCONVERT calculation precisions.
 */
const EURO_CURRENCIES: Readonly<Record<string, EuroCurrency>> = {
	EUR: { rate: 1, precision: 2 },
	ATS: { rate: 13.7603, precision: 2 },
	BEF: { rate: 40.3399, precision: 0 },
	CYP: { rate: 0.585274, precision: 2 },
	DEM: { rate: 1.95583, precision: 2 },
	EEK: { rate: 15.6466, precision: 2 },
	ESP: { rate: 166.386, precision: 0 },
	FIM: { rate: 5.94573, precision: 2 },
	FRF: { rate: 6.55957, precision: 2 },
	GRD: { rate: 340.75, precision: 0 },
	HRK: { rate: 7.5345, precision: 2 },
	IEP: { rate: 0.787564, precision: 2 },
	ITL: { rate: 1936.27, precision: 0 },
	LTL: { rate: 3.4528, precision: 2 },
	LUF: { rate: 40.3399, precision: 0 },
	LVL: { rate: 0.702804, precision: 2 },
	MTL: { rate: 0.4293, precision: 2 },
	NLG: { rate: 2.20371, precision: 2 },
	PTE: { rate: 200.482, precision: 0 },
	SIT: { rate: 239.64, precision: 2 },
	SKK: { rate: 30.126, precision: 2 },
};

function roundDecimal(value: number, decimals: number): number {
	if (value === 0) return 0;
	const factor = 10 ** decimals;
	return (Math.sign(value) * Math.round(Math.abs(value) * factor + Number.EPSILON)) / factor;
}

function roundSignificant(value: number, digits: number): number {
	if (value === 0) return 0;
	const abs = Math.abs(value);
	const shift = digits - 1 - Math.floor(Math.log10(abs));
	if (shift >= 0) {
		const factor = 10 ** shift;
		return (Math.sign(value) * Math.round(abs * factor + Number.EPSILON)) / factor;
	}
	const factor = 10 ** -shift;
	return Math.sign(value) * Math.round(abs / factor + Number.EPSILON) * factor;
}

export const EUROCONVERT: FunctionSpec = spec(
	'EUROCONVERT',
	'Financial',
	'EUROCONVERT(number, source, target, [full_precision], [triangulation_precision])',
	'Converts a number to euros, from euros to a euro member currency, or between euro member currencies using fixed EU rates.',
	3,
	5,
	(args) => {
		const nArg = scalar(args[0]);
		const srcArg = scalar(args[1]);
		const dstArg = scalar(args[2]);
		if (isError(nArg)) return nArg;
		if (isError(srcArg)) return srcArg;
		if (isError(dstArg)) return dstArg;
		const amount = num(nArg);
		const sourceCode = str(srcArg).toUpperCase();
		const targetCode = str(dstArg).toUpperCase();
		const source = EURO_CURRENCIES[sourceCode];
		const target = EURO_CURRENCIES[targetCode];
		if (!source || !target) return ERR.VALUE;

		const fullPrecision = args[3] === undefined || args[3] === null ? false : bool(scalar(args[3]));
		let triPrecision: number | undefined;
		if (args[4] !== undefined && args[4] !== null) {
			const triArg = scalar(args[4]);
			if (isError(triArg)) return triArg;
			const parsed = Math.trunc(num(triArg));
			if (parsed < 3) return ERR.VALUE;
			triPrecision = parsed;
		}

		if (sourceCode === targetCode) return amount;

		let raw: number;
		if (sourceCode === 'EUR') {
			raw = amount * target.rate;
		} else if (targetCode === 'EUR') {
			raw = amount / source.rate;
		} else {
			let eur = amount / source.rate;
			if (triPrecision !== undefined) eur = roundSignificant(eur, triPrecision);
			raw = eur * target.rate;
		}

		return fullPrecision ? raw : roundDecimal(raw, target.precision);
	},
);
