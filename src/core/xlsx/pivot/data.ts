import { getCell } from '../cells';
import { isCellError, type CellValue, type PivotTable, type Workbook } from '../model';
import { formatRange, quoteSheetName } from '../address';
import { sparklineDataRange } from '../edit/sparkline-ranges';
import { parseXml } from '../../xml/index';
import { pivotColumnFields, pivotRowFields } from './axes';
import { pivotCaptions, pivotMeasureName, type CaptionHolder } from './captions';
import { pivotAxisFieldNames } from './date-groups';

export { pivotColumnFields, pivotRowFields };
import { groupLabelsOf, groupRank, validateFieldGroups } from './groups';
import { validateFieldSubtotals } from './subtotals';
import { validatePivotItemFilters } from './item-filters';
import { validatePivotPageFilter } from './page-filter';

export const PIVOT_AGGREGATES = [
	'sum',
	'count',
	'average',
	'min',
	'max',
	'product',
	'countNums',
	'stdDev',
	'stdDevp',
	'var',
	'varp',
] as const;

/** Reads a rectangular, header-led source, retaining blanks and cached formula values. */
export function pivotSourceData(
	workbook: Workbook,
	sheet: number,
	source: string,
): Pick<PivotTable, 'source' | 'fields' | 'records'> {
	const data = sparklineDataRange(workbook, sheet, source);
	const ws = workbook.sheets[data.sheet]!;
	const { start, end } = data.range;
	if (end.row <= start.row || (end.row - start.row + 1) * (end.col - start.col + 1) > 1_000_000)
		throw new Error('A PivotTable source needs headers and data, with at most one million cells.');
	const fields: string[] = [];
	for (let col = start.col; col <= end.col; col++) {
		const value = getCell(ws, start.row, col)?.value;
		if (value === undefined || value === null || isCellError(value) || !String(value).trim())
			throw new Error('Every PivotTable source column needs a unique, non-empty header.');
		fields.push(String(value));
	}
	if (new Set(fields.map((name) => name.toLowerCase())).size !== fields.length)
		throw new Error('Every PivotTable source column needs a unique, non-empty header.');
	const records: CellValue[][] = [];
	for (let row = start.row + 1; row <= end.row; row++) {
		const record: CellValue[] = [];
		for (let col = start.col; col <= end.col; col++)
			record.push(getCell(ws, row, col)?.value ?? null);
		records.push(record);
	}
	return { source: `${quoteSheetName(ws.name)}!${formatRange(data.range)}`, fields, records };
}

/** Type tags keep text "1", the number 1, booleans and blanks in distinct groups. */
export function pivotValueKey(value: CellValue): string {
	return value === null || value === ''
		? 'blank'
		: isCellError(value)
			? `error:${value.error}`
			: `${typeof value}:${String(value).toLowerCase()}`;
}

export function pivotItems(
	pivot: Pick<PivotTable, 'fields' | 'records'> & {
		fieldGroups?: PivotTable['fieldGroups'];
		dateGroups?: PivotTable['dateGroups'];
		date1904?: boolean | undefined;
	},
	field: string,
): CellValue[] {
	const group = groupLabelsOf(pivot, field);
	const index = pivot.fields.indexOf(field);
	const items = new Map<string, CellValue>();
	for (const record of pivot.records) {
		const value = record[index] ?? null;
		const key = pivotValueKey(value);
		if (!items.has(key)) items.set(key, value);
	}
	const rank = (value: CellValue) =>
		value === null || value === ''
			? 4
			: isCellError(value)
				? 3
				: typeof value === 'boolean'
					? 2
					: typeof value === 'string'
						? 1
						: 0;
	return [...items.values()].sort(
		(a, b) =>
			(group ? groupRank(group, a) - groupRank(group, b) : 0) ||
			rank(a) - rank(b) ||
			(typeof a === 'number' && typeof b === 'number'
				? a - b
				: String(pivotItemLabel(a)).localeCompare(String(pivotItemLabel(b)), 'en', {
						sensitivity: 'base',
						numeric: true,
					})),
	);
}

/** Native PivotTable labels display booleans and errors as text; numeric items remain numeric. */
export function pivotItemLabel(value: CellValue): CellValue {
	if (value === null || value === '') return '(blank)';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	return isCellError(value) ? value.error : value;
}

export function validatePivotFields(
	pivot: Pick<
		PivotTable,
		| 'fields'
		| 'rowField'
		| 'layout'
		| 'subtotalPosition'
		| 'secondRowField'
		| 'moreRowFields'
		| 'subtotals'
		| 'fieldSubtotals'
		| 'dateGroups'
		| 'date1904'
		| 'repeatLabels'
		| 'columnField'
		| 'moreColumnFields'
		| 'valueField'
		| 'aggregate'
		| 'valueCaption'
		| 'additionalValues'
		| 'sourceXml'
		| 'valuesAxis'
		| 'valuesPosition'
		| 'captions'
		| 'pageFilter'
		| 'itemFilters'
		| 'dateFilter'
		| 'dateFields'
		| 'fieldGroups'
		| 'records'
	>,
): void {
	validatePivotPageFilter(pivot);
	validatePivotItemFilters(pivot);
	validateFieldGroups(pivot);
	const rows = pivotRowFields(pivot);
	if (
		pivot.moreRowFields !== undefined &&
		(!Array.isArray(pivot.moreRowFields) ||
			(pivot.moreRowFields.length > 0 && !pivot.secondRowField) ||
			pivot.moreRowFields.length > 62 ||
			pivot.moreRowFields.some((name) => typeof name !== 'string'))
	)
		throw new Error('Choose valid PivotTable row fields.');
	const columns = pivotColumnFields(pivot);
	if (
		pivot.moreColumnFields !== undefined &&
		(!Array.isArray(pivot.moreColumnFields) ||
			(pivot.moreColumnFields.length > 0 && !pivot.columnField) ||
			pivot.moreColumnFields.length > 62 ||
			pivot.moreColumnFields.some((name) => typeof name !== 'string'))
	)
		throw new Error('Choose valid PivotTable column fields.');
	if (pivot.pageFilter && columns.length > 1)
		throw new Error('Report filters support one column field.');
	validateFieldSubtotals(pivot, [...rows, ...columns]);
	const axisNames = pivotAxisFieldNames(pivot);
	if (
		(pivot.dateGroups ?? []).some((group) =>
			[...rows, ...columns].some((name) => name === group.field),
		)
	)
		throw new Error('A date field with groups is replaced on the axes by its group fields.');
	const extras = pivot.additionalValues ?? [];
	if (pivot.valuesAxis !== undefined && !['columns', 'rows'].includes(pivot.valuesAxis))
		throw new Error('Choose a valid Values axis.');
	const position = pivot.valuesPosition;
	if (
		position !== undefined &&
		!['last', 'between', 'first'].includes(position as string) &&
		!(typeof position === 'number' && Number.isInteger(position) && position >= 0)
	)
		throw new Error('Choose a valid Values position.');
	if (
		position !== undefined &&
		position !== 'last' &&
		(pivot.valuesAxis !== 'rows' ||
			(position === 'between' && !pivot.secondRowField) ||
			(typeof position === 'number' && position > rows.length))
	)
		throw new Error('Values before or between fields require two row fields and Values on rows.');
	if (pivot.valuesAxis === 'rows' && (!Array.isArray(extras) || !extras.length))
		throw new Error('Values on rows require multiple measures.');
	if (
		extras.length > 0 &&
		pivot.sourceXml &&
		parseXml(pivot.sourceXml).documentElement.getAttribute('dataCaption') !==
			pivotCaptions(pivot).values
	)
		throw new Error('This PivotTable uses unsupported Values header captions.');
	if (
		!Array.isArray(extras) ||
		extras.length > 15 ||
		extras.some(
			(measure) =>
				!measure ||
				!pivot.fields.includes(measure.field) ||
				!PIVOT_AGGREGATES.includes(measure.aggregate) ||
				typeof measure.caption !== 'string' ||
				!measure.caption.trim() ||
				measure.caption.length > 255 ||
				pivot.fields.some((field) => field.toLowerCase() === measure.caption.toLowerCase()),
		) ||
		new Set(
			[pivot.valueCaption, ...extras.map((measure) => measure.caption)].map((caption) =>
				caption?.toLowerCase(),
			),
		).size !==
			extras.length + 1
	)
		throw new Error('Choose PivotTable value fields with valid aggregates and unique captions.');
	if (
		(pivot.layout !== undefined && !['tabular', 'compact', 'outline'].includes(pivot.layout)) ||
		(pivot.subtotalPosition !== undefined && !['top', 'bottom'].includes(pivot.subtotalPosition)) ||
		!PIVOT_AGGREGATES.includes(pivot.aggregate) ||
		!pivot.fields.includes(pivot.valueField) ||
		[...rows, ...columns].some((field) => !axisNames.includes(field)) ||
		new Set([...rows, ...columns]).size !== rows.length + columns.length ||
		(pivot.subtotals !== undefined && typeof pivot.subtotals !== 'boolean') ||
		(pivot.repeatLabels !== undefined && typeof pivot.repeatLabels !== 'boolean')
	)
		throw new Error('Choose valid PivotTable fields; row and column fields must differ.');
}

export function pivotCaption(
	aggregate: PivotTable['aggregate'],
	field: string,
	pivot?: CaptionHolder,
): string {
	return pivotMeasureName(aggregate, field, pivot);
}
