import type { XmlElement } from '../../xml/index';
import type {
	AutoFilterColumn,
	DateGroupItem,
	DateGroupLevel,
	DifferentialStyle,
	DynamicFilterType,
	FilterOperator,
} from '../model';
import { att, boolAttr, numAttr, outerXml, xChildren, xFirst } from './xml-util';

/** The default SpreadsheetML declaration xmldom adds to a serialized fragment's first tag. */
const MAIN_NS_DECL =
	/(?<=^<[^>]*) xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/;

const OPERATORS = new Set<string>([
	'equal',
	'lessThan',
	'lessThanOrEqual',
	'notEqual',
	'greaterThanOrEqual',
	'greaterThan',
]);
const LEVELS = new Set<string>(['year', 'month', 'day', 'hour', 'minute', 'second']);

function readDateGroup(node: XmlElement): DateGroupItem | undefined {
	const grouping = att(node, 'dateTimeGrouping');
	const year = numAttr(node, 'year');
	if (year === undefined || !grouping || !LEVELS.has(grouping)) return undefined;
	const item: DateGroupItem = { year, grouping: grouping as DateGroupLevel };
	for (const key of ['month', 'day', 'hour', 'minute', 'second'] as const) {
		const value = numAttr(node, key);
		if (value !== undefined) item[key] = value;
	}
	return item;
}

/** Reads a `<filterColumn>`; `dxfs` resolves a colour filter's differential format. */
export function readFilterColumn(
	column: XmlElement,
	dxfs?: readonly DifferentialStyle[],
): AutoFilterColumn {
	const entry: AutoFilterColumn = { offset: numAttr(column, 'colId') ?? 0 };
	const filters = xFirst(column, 'filters');
	if (filters) {
		entry.values = xChildren(filters, 'filter').map((f) => att(f, 'val') ?? '');
		if (boolAttr(filters, 'blank')) entry.blank = true;
		const groups = xChildren(filters, 'dateGroupItem')
			.map(readDateGroup)
			.filter((g): g is DateGroupItem => g !== undefined);
		if (groups.length) entry.dateGroups = groups;
	}
	const custom = xFirst(column, 'customFilters');
	if (custom) {
		const conditions = xChildren(custom, 'customFilter').map((c) => {
			const op = att(c, 'operator') ?? 'equal';
			return {
				operator: (OPERATORS.has(op) ? op : 'equal') as FilterOperator,
				value: att(c, 'val') ?? '',
			};
		});
		entry.custom = { ...(boolAttr(custom, 'and') ? { and: true as const } : {}), conditions };
	}
	const top = xFirst(column, 'top10');
	if (top) {
		const filterValue = numAttr(top, 'filterVal');
		entry.top10 = {
			top: boolAttr(top, 'top', true),
			percent: boolAttr(top, 'percent', false),
			value: numAttr(top, 'val') ?? 10,
			...(filterValue !== undefined ? { filterValue } : {}),
		};
	}
	const dynamic = xFirst(column, 'dynamicFilter');
	if (dynamic) {
		const value = numAttr(dynamic, 'val');
		const maxValue = numAttr(dynamic, 'maxVal');
		entry.dynamic = {
			type: (att(dynamic, 'type') ?? 'null') as DynamicFilterType,
			...(value !== undefined ? { value } : {}),
			...(maxValue !== undefined ? { maxValue } : {}),
		};
	}
	const color = xFirst(column, 'colorFilter');
	if (color) {
		const dxfId = numAttr(color, 'dxfId');
		const style = dxfId === undefined ? undefined : dxfs?.[dxfId];
		entry.color = {
			cellColor: boolAttr(color, 'cellColor', true),
			...(dxfId !== undefined ? { dxfId } : {}),
			...(style ? { style: structuredClone(style) } : {}),
		};
	}
	const icon = xFirst(column, 'iconFilter');
	const iconSet = att(icon, 'iconSet');
	if (iconSet) {
		const iconId = numAttr(icon, 'iconId');
		entry.icon = { iconSet, ...(iconId !== undefined ? { iconId } : {}) };
	}
	if (boolAttr(column, 'hiddenButton')) entry.hiddenButton = true;
	entry.sourceXml = outerXml(column).replace(MAIN_NS_DECL, '');
	return entry;
}
