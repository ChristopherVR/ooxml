import { localDateToSerial, serialSeconds, serialToYmd } from '../date-serial.js';
import { parseNumberText } from '../text-number.js';
import { ERR, fail, isError } from '../values.js';
import { num, optNum, scalar, spec, str } from './helpers.js';
import { WORKDAY_FUNCTIONS } from './workdays.js';
import type { FunctionSpec } from './types.js';

const C = 'Date & Time';
import {
	addMonths,
	dateSerial,
	datedif,
	days360,
	isoWeek,
	serialArg,
	weekNum,
	weekday,
	yearFrac,
} from './date-calc.js';

export const DATETIME_FUNCTIONS: FunctionSpec[] = [
	...WORKDAY_FUNCTIONS,
	spec('DATE', C, 'DATE(year, month, day)', 'The serial number of a date.', 3, 3, (args, ctx) =>
		dateSerial(num(args[0]), num(args[1]), num(args[2]), ctx),
	),
	spec(
		'DATEVALUE',
		C,
		'DATEVALUE(date_text)',
		'Converts date text to a serial number.',
		1,
		1,
		(args) => {
			const v = scalar(args[0]);
			if (typeof v !== 'string') return isError(v) ? v : fail(ERR.VALUE);
			const parsed = parseNumberText(v);
			return parsed === undefined || /^[+-]?[\d.,\s$]+%?$/.test(v.trim())
				? fail(ERR.VALUE)
				: Math.floor(parsed);
		},
	),
	spec(
		'TIMEVALUE',
		C,
		'TIMEVALUE(time_text)',
		'Converts time text to a fraction of a day.',
		1,
		1,
		(args) => {
			const v = scalar(args[0]);
			if (typeof v !== 'string') return isError(v) ? v : fail(ERR.VALUE);
			const parsed = parseNumberText(v);
			if (parsed === undefined || !/:|am|pm/i.test(v)) fail(ERR.VALUE);
			return parsed - Math.floor(parsed);
		},
	),
	spec(
		'DAY',
		C,
		'DAY(serial_number)',
		'The day of the month.',
		1,
		1,
		(args, ctx) => serialToYmd(serialArg(args[0]), ctx.date1904).day,
	),
	spec(
		'MONTH',
		C,
		'MONTH(serial_number)',
		'The month (1-12).',
		1,
		1,
		(args, ctx) => serialToYmd(serialArg(args[0]), ctx.date1904).month,
	),
	spec(
		'YEAR',
		C,
		'YEAR(serial_number)',
		'The year.',
		1,
		1,
		(args, ctx) => serialToYmd(serialArg(args[0]), ctx.date1904).year,
	),
	spec('HOUR', C, 'HOUR(serial_number)', 'The hour (0-23).', 1, 1, (args) =>
		Math.floor(serialSeconds(serialArg(args[0])) / 3600),
	),
	spec(
		'MINUTE',
		C,
		'MINUTE(serial_number)',
		'The minute (0-59).',
		1,
		1,
		(args) => Math.floor(serialSeconds(serialArg(args[0])) / 60) % 60,
	),
	spec(
		'SECOND',
		C,
		'SECOND(serial_number)',
		'The second (0-59).',
		1,
		1,
		(args) => serialSeconds(serialArg(args[0])) % 60,
	),
	spec(
		'TIME',
		C,
		'TIME(hour, minute, second)',
		'The fraction of a day for a time.',
		3,
		3,
		(args) => {
			const h = Math.trunc(num(args[0]));
			const m = Math.trunc(num(args[1]));
			const s = Math.trunc(num(args[2]));
			if (h > 32_767 || m > 32_767 || s > 32_767) fail(ERR.NUM);
			const total = h * 3600 + m * 60 + s;
			if (total < 0) fail(ERR.NUM);
			return (total % 86_400) / 86_400;
		},
	),
	spec(
		'NOW',
		C,
		'NOW()',
		'The current date and time.',
		0,
		0,
		(_a, ctx) => localDateToSerial(ctx.frame.host.now(), ctx.date1904),
		undefined,
		true,
	),
	spec(
		'TODAY',
		C,
		'TODAY()',
		'The current date.',
		0,
		0,
		(_a, ctx) => Math.floor(localDateToSerial(ctx.frame.host.now(), ctx.date1904)),
		undefined,
		true,
	),
	spec(
		'WEEKDAY',
		C,
		'WEEKDAY(serial_number, [return_type])',
		'The day of the week.',
		1,
		2,
		(args, ctx) => weekday(serialArg(args[0]), Math.trunc(optNum(args, 1, 1)), ctx),
	),
	spec(
		'WEEKNUM',
		C,
		'WEEKNUM(serial_number, [return_type])',
		'The week number of the year.',
		1,
		2,
		(args, ctx) => weekNum(serialArg(args[0]), Math.trunc(optNum(args, 1, 1)), ctx),
	),
	spec('ISOWEEKNUM', C, 'ISOWEEKNUM(date)', 'The ISO week number of the year.', 1, 1, (args, ctx) =>
		isoWeek(serialArg(args[0]), ctx),
	),
	spec(
		'EDATE',
		C,
		'EDATE(start_date, months)',
		'The date a number of months away.',
		2,
		2,
		(args, ctx) => addMonths(num(args[0]), num(args[1]), ctx, false),
	),
	spec(
		'EOMONTH',
		C,
		'EOMONTH(start_date, months)',
		'The last day of a month a number of months away.',
		2,
		2,
		(args, ctx) => addMonths(num(args[0]), num(args[1]), ctx, true),
	),
	spec(
		'DATEDIF',
		C,
		'DATEDIF(start_date, end_date, unit)',
		'The difference between dates in years, months or days.',
		3,
		3,
		(args, ctx) => datedif(serialArg(args[0]), serialArg(args[1]), str(args[2]), ctx),
	),
	spec(
		'DAYS',
		C,
		'DAYS(end_date, start_date)',
		'The number of days between two dates.',
		2,
		2,
		(args) => Math.floor(serialArg(args[0])) - Math.floor(serialArg(args[1])),
	),
	spec(
		'DAYS360',
		C,
		'DAYS360(start_date, end_date, [method])',
		'Days between dates on a 360-day year.',
		2,
		3,
		(args, ctx) =>
			days360(
				serialArg(args[0]),
				serialArg(args[1]),
				args.length > 2 && Boolean(num(args[2])),
				ctx,
			),
	),
	spec(
		'YEARFRAC',
		C,
		'YEARFRAC(start_date, end_date, [basis])',
		'The fraction of a year between dates.',
		2,
		3,
		(args, ctx) =>
			yearFrac(serialArg(args[0]), serialArg(args[1]), Math.trunc(optNum(args, 2, 0)), ctx),
	),
];
