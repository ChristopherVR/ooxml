import { NS, parseXml } from '../../xml/index';
import type { AutoFilter, AutoFilterColumn } from '../model';
import { readAutoFilter } from '../read/sheet-props';
import { sameModel } from './snapshot';
import type { StyleWriter } from './styles';
import { el } from './xml-out';

/** The criteria of a column without its resolved dxf (the file only stores the id). */
export function criteriaOf(column: AutoFilterColumn) {
	const { values, blank, dateGroups, custom, top10, dynamic, color, icon, hiddenButton } = column;
	const bare = color && { cellColor: color.cellColor, dxfId: color.dxfId };
	return { values, blank, dateGroups, custom, top10, dynamic, color: bare, icon, hiddenButton };
}

/** `filter` without resolved colour formats, to compare a model with its re-read source. */
export function unresolvedFilter(filter: AutoFilter | undefined) {
	return (
		filter && {
			range: filter.range,
			sortState: filter.sortState && {
				...filter.sortState,
				conditions: filter.sortState.conditions.map(({ style: _style, ...rest }) => rest),
			},
			columns: filter.columns?.map((column) => ({
				offset: column.offset,
				sourceXml: column.sourceXml,
				...criteriaOf(column),
			})),
		}
	);
}

/** The modelled criteria of a kept `<filterColumn>`, to tell whether an edit replaced them. */
function sourceCriteria(xml: string) {
	try {
		const wrapped = `<autoFilter xmlns="${NS.x}" ref="A1">${xml}</autoFilter>`;
		const column = readAutoFilter(parseXml(wrapped).documentElement)?.columns?.[0];
		return column ? criteriaOf(column) : undefined;
	} catch {
		return undefined;
	}
}

function hasCriteria(column: AutoFilterColumn): boolean {
	return !!(
		column.values?.length ||
		column.blank ||
		column.dateGroups?.length ||
		column.custom?.conditions.length ||
		column.top10 ||
		column.dynamic ||
		column.color ||
		column.icon
	);
}

function criteriaXml(column: AutoFilterColumn, styles: StyleWriter | undefined): string {
	if (column.top10) {
		const { top, percent, value, filterValue } = column.top10;
		return el('top10', {
			top: top ? undefined : false,
			percent: percent || undefined,
			val: value,
			filterVal: filterValue,
		});
	}
	if (column.custom?.conditions.length) {
		const items = column.custom.conditions
			.slice(0, 2)
			.map((c) =>
				el('customFilter', {
					operator: c.operator === 'equal' ? undefined : c.operator,
					val: c.value,
				}),
			)
			.join('');
		return el('customFilters', { and: column.custom.and || undefined }, items);
	}
	if (column.dynamic) {
		const { type, value, maxValue } = column.dynamic;
		return el('dynamicFilter', { type, val: value, maxVal: maxValue });
	}
	if (column.color) {
		const { cellColor, style, dxfId } = column.color;
		const id = style && styles ? styles.dxfId(style) : dxfId;
		return el('colorFilter', { dxfId: id, cellColor: cellColor ? undefined : false });
	}
	if (column.icon)
		return el('iconFilter', { iconSet: column.icon.iconSet, iconId: column.icon.iconId });
	const values = (column.values ?? []).map((value) => el('filter', { val: value }));
	const groups = (column.dateGroups ?? []).map((g) =>
		el('dateGroupItem', {
			year: g.year,
			month: g.month,
			day: g.day,
			hour: g.hour,
			minute: g.minute,
			second: g.second,
			dateTimeGrouping: g.grouping,
		}),
	);
	return el('filters', { blank: column.blank || undefined }, values.join('') + groups.join(''));
}

export function filterColumnXml(column: AutoFilterColumn, styles?: StyleWriter): string {
	if (column.sourceXml && sameModel(sourceCriteria(column.sourceXml), criteriaOf(column))) {
		const head = /^<filterColumn\b[^>]*?>/.exec(column.sourceXml)?.[0];
		if (head) {
			const colId = ` colId="${column.offset}"`;
			const patched = /\scolId="[^"]*"/.test(head)
				? head.replace(/\scolId="[^"]*"/, colId)
				: head.replace(/^<filterColumn\b/, `<filterColumn${colId}`);
			return patched + column.sourceXml.slice(head.length);
		}
	}
	// Criteria the model does not hold are never written as an empty `<filterColumn>`.
	if (!hasCriteria(column)) return '';
	return el(
		'filterColumn',
		{ colId: column.offset, hiddenButton: column.hiddenButton || undefined },
		criteriaXml(column, styles),
	);
}
