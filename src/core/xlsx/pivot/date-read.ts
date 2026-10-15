import { NS, children, elements, first, type XmlElement } from '../../xml/index';
import type { PivotDateFilter } from '../model';
import { att } from '../read/xml-util';
import { isoToSerial } from './dates';

/** Cache fields that hold only dates (`containsDate`), shown with the date format 14. */
export function readDateFields(fields: XmlElement[], names: string[]): string[] {
	const dates: string[] = [];
	fields.forEach((field, index) => {
		const items = first(field, 'sharedItems', NS.x);
		if (att(items, 'containsDate') !== '1') {
			if (att(field, 'numFmtId') !== '0') throw new Error('Unsupported pivot field number format');
			return;
		}
		if (
			att(items, 'containsNonDate') !== '0' ||
			att(items, 'containsString') !== '0' ||
			att(items, 'containsBlank') === '1' ||
			att(field, 'numFmtId') !== '14'
		)
			throw new Error('Unsupported pivot date field');
		dates.push(names[index] ?? '');
	});
	return dates;
}

/**
 * The timeline date range of a PivotTable: one `dateBetween` filter with two inclusive
 * `customFilter` bounds. Any other filter is preserved unsupported.
 */
export function readDateFilter(
	root: XmlElement,
	names: string[],
	dateFields: string[],
	date1904: boolean,
): PivotDateFilter | undefined {
	const list = first(root, 'filters', NS.x);
	if (!list) return undefined;
	const entries = children(list, 'filter', NS.x);
	const filter = entries[0];
	if (entries.length !== 1 || att(list, 'count') !== '1' || !filter)
		throw new Error('Only one timeline filter is modeled');
	const field = names[Number(att(filter, 'fld'))];
	const columns = children(first(filter, 'autoFilter', NS.x) ?? list, 'filterColumn', NS.x);
	const custom = first(columns[0], 'customFilters', NS.x);
	const bounds = custom ? children(custom, 'customFilter', NS.x) : [];
	if (
		att(filter, 'type') !== 'dateBetween' ||
		!field ||
		!dateFields.includes(field) ||
		columns.length !== 1 ||
		att(custom, 'and') !== '1' ||
		bounds.length !== 2 ||
		att(bounds[0], 'operator') !== 'greaterThanOrEqual' ||
		att(bounds[1], 'operator') !== 'lessThanOrEqual' ||
		elements(filter).some((node) => !['autoFilter', 'extLst'].includes(node.localName ?? ''))
	)
		throw new Error('Unsupported pivot filter');
	const [start, end] = bounds.map((node) => {
		const raw = att(node, 'val') ?? '';
		const value = /^\d+(\.\d+)?$/.test(raw) ? Number(raw) : isoToSerial(raw, date1904);
		if (value === undefined) throw new Error('Unsupported pivot filter date');
		return value;
	});
	if (start === undefined || end === undefined || start > end)
		throw new Error('Unsupported pivot filter range');
	return { field, start, end };
}
