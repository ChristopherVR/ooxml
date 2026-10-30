// Excel's built-in sort custom lists (Data > Sort > Order > Custom List...).
const WEEKDAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS_SHORT = [
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
];
const MONTHS = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
];

export const BUILT_IN_SORT_LISTS: readonly (readonly string[])[] = [
	WEEKDAYS_SHORT,
	WEEKDAYS,
	MONTHS_SHORT,
	MONTHS,
];

/** Splits user text (one item per line or comma-separated) into list items. */
export function parseSortList(text: string): string[] {
	return text
		.split(/[\r\n,]+/)
		.map((item) => item.trim())
		.filter((item) => item !== '');
}

/** The `customList` attribute of a sort condition: the items joined by commas. */
export const formatSortList = (items: readonly string[]): string => items.join(',');
