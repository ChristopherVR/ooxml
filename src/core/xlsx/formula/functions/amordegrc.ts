import { ERR, fail } from '../values';
import { serialArg, yearFrac } from './date-calc';
import { finite, num, optNum, spec } from './helpers';

const round = (value: number): number => Math.floor(value + 0.5);

/**
 * Declining-balance French depreciation with a life-dependent coefficient.
 * https://support.microsoft.com/en-us/office/amordegrc-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51
 */
export const AMORDEGRC = spec(
	'AMORDEGRC',
	'Financial',
	'AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis])',
	'Depreciation of each accounting period using a depreciation coefficient.',
	6,
	7,
	(args, ctx) => {
		let cost = num(args[0]);
		const salvage = num(args[3]);
		const period = Math.floor(num(args[4]));
		let rate = num(args[5]);
		const basis = Math.trunc(optNum(args, 6, 0));
		const purchase = Math.floor(serialArg(args[1]));
		const firstPeriod = Math.floor(serialArg(args[2]));
		if (
			cost <= 0 ||
			salvage < 0 ||
			salvage > cost ||
			rate <= 0 ||
			period < 0 ||
			![0, 1, 3, 4].includes(basis) ||
			purchase > firstPeriod
		)
			fail(ERR.NUM);
		const life = 1 / rate;
		rate *= life < 3 ? 1 : life < 5 ? 1.5 : life <= 6 ? 2 : 2.5;
		let depreciation = round(yearFrac(purchase, firstPeriod, basis, ctx) * rate * cost);
		cost -= depreciation;
		let rest = cost - salvage;
		for (let n = 0; n < period; n++) {
			depreciation = round(rate * cost);
			rest -= depreciation;
			if (rest < 0) return period - n === 1 ? round(cost * 0.5) : 0;
			cost -= depreciation;
		}
		return finite(depreciation);
	},
);
