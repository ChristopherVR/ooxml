import { ERR, fail } from '../values.js';
import { isLeapYear, serialToYmd } from '../date-serial.js';
import { serialArg, yearFrac } from './date-calc.js';
import { finite, num, optNum, spec } from './helpers.js';

/**
 * Straight-line French depreciation, including a prorated initial period.
 * https://support.microsoft.com/en-us/excel/functions/amorlinc-function
 * Native Excel 16 evidence also establishes equal-date
 * full period, fractional period/basis truncation and omitted-required-slot errors.
 */
export const AMORLINC = {
	...spec(
		'AMORLINC',
		'Financial',
		'AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])',
		'Straight-line depreciation with a prorated initial accounting period.',
		6,
		7,
		(args, ctx) => {
			const cost = num(args[0]),
				salvage = num(args[3]),
				period = num(args[4]),
				rate = num(args[5]);
			const rawBasis = optNum(args, 6, 0),
				basis = Math.trunc(rawBasis);
			let purchase = Math.floor(serialArg(args[1])),
				firstPeriod = Math.floor(serialArg(args[2]));
			if (
				cost <= 0 ||
				salvage < 0 ||
				salvage > cost ||
				rate <= 0 ||
				period < 0 ||
				rawBasis < 0 ||
				![0, 1, 3, 4].includes(basis) ||
				purchase > firstPeriod
			)
				fail(ERR.NUM);
			const annual = finite(cost * rate);
			// AMORLINC's actual-day bases normalize Feb 29 to Feb 28. Basis 1
			// divides by the purchase year's length, unlike general YEARFRAC.
			if (basis === 1 || basis === 3) {
				const normalize = (serial: number) => {
					const date = serialToYmd(serial, ctx.date1904);
					return date.month === 2 && date.day === 29 ? serial - 1 : serial;
				};
				purchase = normalize(purchase);
				firstPeriod = normalize(firstPeriod);
			}
			const fraction =
				purchase === firstPeriod
					? 1
					: basis === 1
						? (firstPeriod - purchase) /
							(isLeapYear(serialToYmd(purchase, ctx.date1904).year) ? 366 : 365)
						: yearFrac(purchase, firstPeriod, basis, ctx);
			const first = Math.min(cost - salvage, annual * fraction);
			const index = Math.trunc(period);
			if (index === 0) return first;
			return Math.max(0, Math.min(annual, cost - salvage - first - (index - 1) * annual));
		},
	),
	missingRequiredError: '#N/A' as const,
};
