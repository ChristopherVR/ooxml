// The Text, Number and Date Filters submenus of the AutoFilter drop-down as data: a label and
// either a dialog to open or criteria to apply.
import type { CustomOperator, DynamicFilterType, FilterColumnKind } from 'ooxml-core/xlsx';
import type { CustomAutoFilterProps } from '../dialogs/custom-autofilter';

export interface PresetItem {
	/** English key. */
	label: string;
	separatorBefore?: boolean;
	/** Opens the Custom AutoFilter dialog. */
	custom?: CustomAutoFilterProps;
	/** Opens the Top 10 dialog. */
	top10?: true;
	/** Applies a dynamic filter directly. */
	dynamic?: DynamicFilterType;
	children?: PresetItem[];
}

const op = (label: string, operator: CustomOperator): PresetItem => ({
	label,
	custom: { operator },
});
const dyn = (label: string, dynamic: DynamicFilterType): PresetItem => ({ label, dynamic });

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

export function presetItems(kind: FilterColumnKind): PresetItem[] {
	if (kind === 'text')
		return [
			op('Equals...', 'equals'),
			op('Does Not Equal...', 'notEquals'),
			{ ...op('Begins With...', 'beginsWith'), separatorBefore: true },
			op('Ends With...', 'endsWith'),
			{ ...op('Contains...', 'contains'), separatorBefore: true },
			op('Does Not Contain...', 'notContains'),
			{ label: 'Custom Filter...', custom: {}, separatorBefore: true },
		];
	if (kind === 'number')
		return [
			op('Equals...', 'equals'),
			op('Does Not Equal...', 'notEquals'),
			{ ...op('Greater Than...', 'greaterThan'), separatorBefore: true },
			op('Greater Than Or Equal To...', 'greaterThanOrEqual'),
			op('Less Than...', 'lessThan'),
			op('Less Than Or Equal To...', 'lessThanOrEqual'),
			{ label: 'Between...', custom: { between: true } },
			{ label: 'Top 10...', top10: true, separatorBefore: true },
			{ ...dyn('Above Average', 'aboveAverage'), separatorBefore: true },
			dyn('Below Average', 'belowAverage'),
			{ label: 'Custom Filter...', custom: {}, separatorBefore: true },
		];
	return [
		op('Equals...', 'equals'),
		{ ...op('Before...', 'lessThan'), separatorBefore: true },
		op('After...', 'greaterThan'),
		{ label: 'Between...', custom: { between: true } },
		{ ...dyn('Tomorrow', 'tomorrow'), separatorBefore: true },
		dyn('Today', 'today'),
		dyn('Yesterday', 'yesterday'),
		{ ...dyn('Next Week', 'nextWeek'), separatorBefore: true },
		dyn('This Week', 'thisWeek'),
		dyn('Last Week', 'lastWeek'),
		{ ...dyn('Next Month', 'nextMonth'), separatorBefore: true },
		dyn('This Month', 'thisMonth'),
		dyn('Last Month', 'lastMonth'),
		{ ...dyn('Next Quarter', 'nextQuarter'), separatorBefore: true },
		dyn('This Quarter', 'thisQuarter'),
		dyn('Last Quarter', 'lastQuarter'),
		{ ...dyn('Next Year', 'nextYear'), separatorBefore: true },
		dyn('This Year', 'thisYear'),
		dyn('Last Year', 'lastYear'),
		{ ...dyn('Year to Date', 'yearToDate'), separatorBefore: true },
		{
			label: 'All Dates in the Period',
			children: [
				...[1, 2, 3, 4].map((q) => dyn(`Quarter ${q}`, `Q${q}` as DynamicFilterType)),
				...MONTHS.map((m, i) => ({
					...dyn(m, `M${i + 1}` as DynamicFilterType),
					...(i === 0 ? { separatorBefore: true } : {}),
				})),
			],
		},
		{ label: 'Custom Filter...', custom: {}, separatorBefore: true },
	];
}

/** The submenu title for a column kind. */
export const presetTitle = (kind: FilterColumnKind): string =>
	kind === 'text' ? 'Text Filters' : kind === 'number' ? 'Number Filters' : 'Date Filters';

/** Sort menu labels for a column kind, ascending then descending. */
export const sortLabels = (kind: FilterColumnKind): [string, string] =>
	kind === 'text'
		? ['Sort A to Z', 'Sort Z to A']
		: kind === 'number'
			? ['Sort Smallest to Largest', 'Sort Largest to Smallest']
			: ['Sort Oldest to Newest', 'Sort Newest to Oldest'];
