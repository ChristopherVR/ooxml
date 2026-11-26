import { NS, elements, first, type XmlElement } from '../../xml/index';
import type { CellValue, PivotCondition, PivotTable } from '../model';
import { pivotValueKey } from './data';
import { pivotItemLabeler } from './item-labels';
import { pivotMeasureNumber, pivotMeasures } from './measures';
import { attrs, encodeEscapes } from '../write/xml-out';
import { decodeEscapes } from '../read/xml-util';
import { compileWildcard } from '../formula/functions/wildcard';

function captionCriterion(pattern: string): { wildcard: boolean; value: string } {
	// Excel stores escaped literals in a checklist and uses a custom filter for wildcards.
	if (!/[*?]/.test(pattern)) return { wildcard: false, value: pattern };
	const tokens = pattern.match(/~[\s\S]|[*?]/g) ?? [];
	if (tokens.some((token) => token.length === 1)) return { wildcard: true, value: pattern };
	return { wildcard: false, value: pattern.replace(/~([\s\S]|$)/g, '$1') };
}

export function supportsPivotCondition(pivot: PivotTable): boolean {
	return !(
		pivot.columnField ||
		pivot.pageFilter ||
		pivot.dateFilter ||
		pivot.dateGroups?.length ||
		pivot.fieldGroups?.length ||
		pivot.itemFilters?.length ||
		pivot.slicerCacheId !== undefined ||
		pivot.calculatedField
	);
}

export function applyPivotCondition(pivot: PivotTable): PivotTable {
	const condition = pivot.condition;
	if (!condition) return pivot;
	const field = pivot.fields.indexOf(pivot.rowField);
	if (field < 0 || !supportsPivotCondition(pivot))
		throw new Error('This PivotTable layout does not support caption or value filters.');
	const label = pivotItemLabeler(pivot);
	if (condition.kind === 'captionEqual') {
		if (typeof condition.value !== 'string' || condition.value.length > 255)
			throw new Error('Enter a caption of at most 255 characters.');
		const criterion = captionCriterion(condition.value);
		const matcher = criterion.wildcard ? compileWildcard(criterion.value) : undefined;
		const literal = criterion.value.toLowerCase();
		return {
			...pivot,
			records: pivot.records.filter((record) => {
				const caption = String(label(pivot.rowField, record[field] ?? null));
				return matcher ? matcher.test(caption) : caption.toLowerCase() === literal;
			}),
		};
	}
	const measure = pivotMeasures(pivot)[condition.measure];
	if (
		!measure ||
		!Number.isInteger(condition.measure) ||
		!Number.isFinite(condition.value) ||
		(condition.kind !== 'valueGreaterThan' && condition.kind !== 'valueLessThan')
	)
		throw new Error('Choose a measure and enter a finite filter value.');
	const groups = new Map<string, CellValue[][]>();
	for (const record of pivot.records) {
		const key = pivotValueKey(record[field] ?? null);
		const group = groups.get(key);
		if (group) group.push(record);
		else groups.set(key, [record]);
	}
	const included = new Set<string>();
	for (const [key, records] of groups) {
		const total = pivotMeasureNumber(pivot, records, measure);
		if (
			typeof total === 'number' &&
			(condition.kind === 'valueGreaterThan' ? total > condition.value : total < condition.value)
		)
			included.add(key);
	}
	return {
		...pivot,
		records: pivot.records.filter((record) => included.has(pivotValueKey(record[field] ?? null))),
	};
}

const operator = (condition: Exclude<PivotCondition, { kind: 'captionEqual' }>) =>
	condition.kind === 'valueGreaterThan' ? 'greaterThan' : 'lessThan';

export function pivotConditionXml(pivot: PivotTable, condition: PivotCondition): string {
	const common = {
		fld: pivot.fields.indexOf(pivot.rowField),
		type: condition.kind,
		evalOrder: -1,
		id: 1,
	};
	const caption = condition.kind === 'captionEqual';
	const field = caption
		? { stringValue1: encodeEscapes(condition.value) }
		: { iMeasureFld: condition.measure };
	const criterion = caption ? captionCriterion(condition.value) : undefined;
	const criteria =
		criterion && !criterion.wildcard
			? `<filters><filter${attrs({ val: encodeEscapes(criterion.value) })}/></filters>`
			: `<customFilters><customFilter${attrs({
					operator: caption ? undefined : operator(condition),
					val: caption ? encodeEscapes(condition.value) : condition.value,
				})}/></customFilters>`;
	return `<filters count="1"><filter${attrs({ ...common, ...field })}><autoFilter ref="A1"><filterColumn colId="0">${criteria}</filterColumn></autoFilter></filter></filters>`;
}

export function readPivotCondition(root: XmlElement, row: number): PivotCondition | undefined {
	const filters = first(root, 'filters', NS.x);
	const filter = first(filters, 'filter', NS.x);
	const kind = filter?.getAttribute('type');
	if (kind !== 'captionEqual' && kind !== 'valueGreaterThan' && kind !== 'valueLessThan') return;
	const auto = first(filter, 'autoFilter', NS.x);
	const column = first(auto, 'filterColumn', NS.x);
	const pattern = kind === 'captionEqual' ? filter?.getAttribute('stringValue1') : undefined;
	const caption = pattern != null ? captionCriterion(decodeEscapes(pattern)) : undefined;
	const custom = !caption || caption.wildcard;
	const criteria = first(column, custom ? 'customFilters' : 'filters', NS.x);
	const criterion = first(criteria, custom ? 'customFilter' : 'filter', NS.x);
	const value = criterion?.getAttribute('val');
	if (
		filters?.getAttribute('count') !== '1' ||
		filter?.getAttribute('fld') !== String(row) ||
		auto?.getAttribute('ref') !== 'A1' ||
		column?.getAttribute('colId') !== '0' ||
		value == null ||
		[filters, filter, auto, column, criteria].some(
			(node) => !node || elements(node).length !== 1,
		) ||
		!criterion ||
		elements(criterion).length ||
		criteria?.hasAttribute('and')
	)
		throw new Error('Unsupported PivotTable condition.');
	if (
		kind === 'captionEqual' &&
		pattern != null &&
		caption &&
		decodeEscapes(value) === caption.value &&
		(!custom ||
			!criterion.hasAttribute('operator') ||
			criterion.getAttribute('operator') === 'equal') &&
		!filter?.hasAttribute('iMeasureFld')
	)
		return { kind, value: decodeEscapes(pattern) };
	const index = filter?.getAttribute('iMeasureFld') ?? '';
	if (
		kind !== 'captionEqual' &&
		!filter?.hasAttribute('stringValue1') &&
		/^\d+$/.test(index) &&
		value.trim() &&
		Number.isFinite(Number(value))
	) {
		const condition: PivotCondition = { kind, measure: Number(index), value: Number(value) };
		if (criterion?.getAttribute('operator') === operator(condition)) return condition;
	}
	throw new Error('Unsupported PivotTable condition.');
}
