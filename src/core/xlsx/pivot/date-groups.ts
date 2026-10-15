import { serialToYmd } from '../formula/date-serial';
import type { CellValue, PivotDateGroup, PivotDateUnit, PivotTable } from '../model';

/** The units in Excel's Group dialog order (finest first), which is also the cache field order. */
export const PIVOT_DATE_UNITS: readonly PivotDateUnit[] = [
	'seconds',
	'minutes',
	'hours',
	'days',
	'months',
	'quarters',
	'years',
];
const UNIT_TITLES: Record<PivotDateUnit, string> = {
	seconds: 'Seconds',
	minutes: 'Minutes',
	hours: 'Hours',
	days: 'Days',
	months: 'Months',
	quarters: 'Quarters',
	years: 'Years',
};
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
/** Month lengths of a leap year: the by-day groups of one day each list 29 February too. */
const LEAP_MONTHS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
const MAX_ITEMS = 10_000;
const pad = (value: number, width = 2) => String(value).padStart(width, '0');

type DateGroupPivot = Pick<PivotTable, 'fields' | 'records'> & {
	dateGroups?: PivotDateGroup[] | undefined;
	date1904?: boolean | undefined;
};

/** The units of a group, finest first, without repeats. */
export const dateGroupUnits = (group: Pick<PivotDateGroup, 'by'>): PivotDateUnit[] =>
	PIVOT_DATE_UNITS.filter((unit) => group.by.includes(unit));

/** The cache field of a unit: `Years (Date)`, or the name the file gave it in its language. */
export const dateGroupFieldName = (
	unit: PivotDateUnit,
	field: string,
	group?: Pick<PivotDateGroup, 'localized'>,
): string => group?.localized?.[unit]?.name ?? `${UNIT_TITLES[unit]} (${field})`;

/** Days per group when grouping by days alone; every other grouping counts single days. */
export const dateGroupDays = (group: Pick<PivotDateGroup, 'by' | 'days'>): number =>
	group.by.length === 1 && group.by[0] === 'days' ? (group.days ?? 1) : 1;

export interface DateRange {
	start: number;
	end: number;
}

/**
 * The grouped range: the chosen start and end, else the earliest date and the latest one (Excel
 * ends a range of whole days one day after the latest, so that day is inside the range).
 */
export function dateGroupRange(
	group: Pick<PivotDateGroup, 'start' | 'end'>,
	serials: readonly number[],
): DateRange {
	let min = Infinity;
	let max = -Infinity;
	if (group.start === undefined || group.end === undefined)
		for (const value of serials) {
			if (value < min) min = value;
			if (value > max) max = value;
		}
	return {
		start: group.start ?? min,
		end: group.end ?? (Number.isInteger(max) ? max + 1 : max),
	};
}

const isoDay = (serial: number, date1904: boolean): string => {
	const { year, month, day } = serialToYmd(serial, date1904);
	return `${pad(year, 4)}-${pad(month)}-${pad(day)}`;
};
const secondsOfDay = (serial: number) =>
	Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;

function dayBuckets(range: DateRange, days: number): number {
	return Math.max(1, Math.ceil((Math.floor(range.end) - Math.floor(range.start) + 1) / days));
}

/** The items of one unit in Excel's order: `<start`, the unit's items, `>end`. */
export function englishDateUnitLabels(
	unit: PivotDateUnit,
	range: DateRange,
	days: number,
	date1904: boolean,
): string[] {
	const labels = [`<${isoDay(range.start, date1904)}`];
	if (unit === 'seconds' || unit === 'minutes')
		for (let i = 0; i < 60; i++) labels.push(`:${pad(i)}`);
	else if (unit === 'hours')
		for (let i = 0; i < 24; i++) labels.push(`${i % 12 || 12} ${i < 12 ? 'AM' : 'PM'}`);
	else if (unit === 'months') labels.push(...MONTHS);
	else if (unit === 'quarters') for (let i = 1; i <= 4; i++) labels.push(`Qtr${i}`);
	else if (unit === 'years') {
		const first = serialToYmd(range.start, date1904).year;
		const last = serialToYmd(range.end, date1904).year;
		for (let year = first; year <= last; year++) labels.push(String(year));
	} else if (days <= 1)
		LEAP_MONTHS.forEach((length, month) => {
			for (let day = 1; day <= length; day++) labels.push(`${pad(day)}-${MONTHS[month]}`);
		});
	else {
		const first = Math.floor(range.start);
		const last = Math.floor(range.end);
		for (let k = 0; k < dayBuckets(range, days); k++) {
			const from = first + k * days;
			labels.push(
				`${isoDay(from, date1904)} - ${isoDay(Math.min(from + days - 1, last), date1904)}`,
			);
		}
	}
	labels.push(`>${isoDay(range.end, date1904)}`);
	return labels;
}

/**
 * The items of one unit as the file shows them: the labels a group read from another language
 * gave them when its range is still the one they were read for, else the English ones.
 */
export function dateUnitLabels(
	unit: PivotDateUnit,
	range: DateRange,
	days: number,
	date1904: boolean,
	group?: Pick<PivotDateGroup, 'localized'>,
): string[] {
	const english = englishDateUnitLabels(unit, range, days, date1904);
	const localized = group?.localized?.[unit]?.labels;
	return localized?.length === english.length ? localized : english;
}

/** The index in `dateUnitLabels` of a date: 0 before the range, the last item after it. */
export function dateUnitIndex(
	unit: PivotDateUnit,
	range: DateRange,
	days: number,
	date1904: boolean,
	serial: number,
	count: number,
): number {
	if (serial < range.start) return 0;
	if (serial > range.end) return count - 1;
	if (unit === 'seconds') return 1 + (secondsOfDay(serial) % 60);
	if (unit === 'minutes') return 1 + (Math.floor(secondsOfDay(serial) / 60) % 60);
	if (unit === 'hours') return 1 + Math.floor(secondsOfDay(serial) / 3600);
	const { year, month, day } = serialToYmd(serial, date1904);
	if (unit === 'months') return month;
	if (unit === 'quarters') return 1 + Math.floor((month - 1) / 3);
	if (unit === 'years') return 1 + year - serialToYmd(range.start, date1904).year;
	if (days <= 1) {
		let before = 0;
		for (let m = 0; m < month - 1; m++) before += LEAP_MONTHS[m]!;
		return 1 + before + day - 1;
	}
	return 1 + Math.floor((Math.floor(serial) - Math.floor(range.start)) / days);
}

export interface DerivedDateField {
	name: string;
	unit: PivotDateUnit;
	group: PivotDateGroup;
	/** Position in the extended field list (the cache fields after the database fields). */
	index: number;
}

/** The derived fields of every date group, in cache order, with their extended positions. */
export function derivedDateFields(pivot: DateGroupPivot): DerivedDateField[] {
	const derived: DerivedDateField[] = [];
	// Expanded copies already list their derived fields; the database fields come first.
	const databaseFields = pivot.fields.length - countDerived(pivot);
	for (const group of pivot.dateGroups ?? [])
		for (const unit of dateGroupUnits(group))
			derived.push({
				name: dateGroupFieldName(unit, group.field, group),
				unit,
				group,
				index: databaseFields + derived.length,
			});
	return derived;
}
function countDerived(pivot: DateGroupPivot): number {
	let count = 0;
	for (const group of pivot.dateGroups ?? [])
		for (const unit of dateGroupUnits(group))
			if (pivot.fields.includes(dateGroupFieldName(unit, group.field, group))) count++;
	return count;
}

/** Every name an axis can use: the source fields and the derived date fields. */
export const pivotAxisFieldNames = (pivot: DateGroupPivot): string[] => [
	...pivot.fields,
	...derivedDateFields(pivot)
		.filter((entry) => !pivot.fields.includes(entry.name))
		.map((entry) => entry.name),
];

/** The groups with their start and end made explicit from the cached dates. */
export function resolveDateGroups(pivot: DateGroupPivot): PivotDateGroup[] {
	return (pivot.dateGroups ?? []).map((group) => {
		const index = pivot.fields.indexOf(group.field);
		const serials = pivot.records
			.map((record) => record[index])
			.filter((value): value is number => typeof value === 'number');
		const range = dateGroupRange(group, serials);
		return { ...group, start: range.start, end: range.end };
	});
}

/** The ordered item labels of a derived date field (its group items), if `name` is one. */
export function derivedDateLabels(pivot: DateGroupPivot, name: string): string[] | undefined {
	const entry = derivedDateFields(pivot).find((candidate) => candidate.name === name);
	if (!entry) return undefined;
	const [group] = resolveDateGroups({ ...pivot, dateGroups: [entry.group] });
	const range = { start: group!.start!, end: group!.end! };
	return dateUnitLabels(entry.unit, range, dateGroupDays(group!), pivot.date1904 ?? false, group);
}

/** Checks the groups against the cached dates and the axes; throws a user-facing message. */
export function validateDateGroups(
	pivot: DateGroupPivot & Pick<PivotTable, 'dateFields' | 'pageFilter' | 'dateFilter'>,
): void {
	const groups = pivot.dateGroups;
	if (groups === undefined) return;
	const invalid = () => new Error('Choose a valid date group.');
	if (!Array.isArray(groups)) throw invalid();
	const seen = new Set<string>();
	for (const group of groups) {
		const index = pivot.fields.indexOf(group?.field as string);
		const finite = (value: number | undefined) => value === undefined || Number.isFinite(value);
		if (
			!group ||
			typeof group !== 'object' ||
			index < 0 ||
			seen.has(group.field) ||
			!Array.isArray(group.by) ||
			!group.by.length ||
			group.by.some((unit) => !PIVOT_DATE_UNITS.includes(unit)) ||
			new Set(group.by).size !== group.by.length ||
			!finite(group.start) ||
			!finite(group.end) ||
			(group.days !== undefined &&
				(!Number.isInteger(group.days) || group.days < 1 || group.days > 32_767)) ||
			(group.days !== undefined && group.days > 1 && group.by.join() !== 'days')
		)
			throw invalid();
		seen.add(group.field);
		if (!(pivot.dateFields ?? []).includes(group.field))
			throw new Error('Only fields that contain dates can be grouped by date.');
		const serials = pivot.records.map((record) => record[index]);
		if (!serials.length || serials.some((v) => typeof v !== 'number' || !Number.isFinite(v)))
			throw new Error('Only fields that contain dates can be grouped by date.');
		const range = dateGroupRange(group, serials as number[]);
		if (range.start >= range.end) throw new Error('The date group must end after it starts.');
		for (const unit of dateGroupUnits(group)) {
			const name = dateGroupFieldName(unit, group.field, group);
			if (pivot.fields.includes(name))
				throw new Error('A source field already has the name of a date group field.');
			if (
				dateUnitLabels(unit, range, dateGroupDays(group), pivot.date1904 ?? false, group).length >
				MAX_ITEMS
			)
				throw new Error('The date group has too many items.');
		}
	}
}

/**
 * The pivot with each date group's derived fields appended: their names to `fields` and, to every
 * record, the group item each date falls in. The cache is unchanged; this is the compute view.
 */
export function expandDateGroups(pivot: PivotTable): PivotTable {
	if (!pivot.dateGroups?.length || countDerived(pivot) > 0) return pivot;
	const groups = resolveDateGroups(pivot);
	const resolved = { ...pivot, dateGroups: groups };
	const derived = derivedDateFields(resolved);
	const date1904 = pivot.date1904 ?? false;
	const tables = derived.map((entry) => {
		const range = { start: entry.group.start!, end: entry.group.end! };
		const days = dateGroupDays(entry.group);
		const labels = dateUnitLabels(entry.unit, range, days, date1904, entry.group);
		return {
			base: pivot.fields.indexOf(entry.group.field),
			labels,
			index: (serial: number) =>
				dateUnitIndex(entry.unit, range, days, date1904, serial, labels.length),
		};
	});
	return {
		...resolved,
		fields: [...pivot.fields, ...derived.map((entry) => entry.name)],
		records: pivot.records.map((record) => [
			...record,
			...tables.map((table): CellValue => {
				const value = record[table.base];
				return typeof value === 'number' ? table.labels[table.index(value)]! : null;
			}),
		]),
	};
}

/** The group the dialog offers: the cached dates by months and years (or one day), automatic range. */
export function defaultDateGroup(
	pivot: DateGroupPivot,
	field: string,
): PivotDateGroup & { start: number; end: number } {
	const existing = pivot.dateGroups?.find((group) => group.field === field);
	const index = pivot.fields.indexOf(field);
	const serials = pivot.records
		.map((record) => record[index])
		.filter((value): value is number => typeof value === 'number');
	const group: PivotDateGroup = existing ?? { field, by: ['months'] };
	return { ...group, ...dateGroupRange(group, serials) };
}

/**
 * The fields an axis can offer: the source fields without the dates that have groups, plus the
 * derived group fields (`Years (Date)`).
 */
export const pivotAxisChoices = (pivot: DateGroupPivot): string[] => {
	const grouped = new Set((pivot.dateGroups ?? []).map((group) => group.field));
	return pivotAxisFieldNames(pivot).filter((name) => !grouped.has(name));
};
