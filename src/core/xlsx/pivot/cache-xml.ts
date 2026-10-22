import { NS } from '../../xml/index';
import type { CellValue, PivotTable, Workbook } from '../model';
import { isCellError } from '../model';
import { formatRange } from '../address';
import { sparklineDataRange } from '../edit/sparkline-ranges';
import { XML_HEADER, attrs, encodeEscapes, escapeAttr } from '../write/xml-out';
import { pivotValueKey } from './data';
import { serialToIso } from './dates';
import {
	dateGroupDays,
	dateGroupRange,
	dateUnitLabels,
	derivedDateFields,
	resolveDateGroups,
	type DerivedDateField,
} from './date-groups';
import { groupOf, numberGroupLabels } from './groups';

export interface PivotCacheData {
	items: CellValue[][];
	indices: Map<string, number>[];
}

export function pivotCacheData(pivot: PivotTable): PivotCacheData {
	const items = pivot.fields.map(() => [] as CellValue[]);
	const indices = pivot.fields.map(() => new Map<string, number>());
	for (const record of pivot.records)
		pivot.fields.forEach((_field, index) => {
			const value = record[index] ?? null;
			const key = pivotValueKey(value);
			if (!indices[index]!.has(key)) {
				indices[index]!.set(key, items[index]!.length);
				items[index]!.push(value);
			}
		});
	return { items, indices };
}

function valueXml(value: CellValue): string {
	if (value === null || value === '') return '<m/>';
	if (isCellError(value)) return `<e${attrs({ v: value.error })}/>`;
	if (typeof value === 'string') return `<s${attrs({ v: encodeEscapes(value) })}/>`;
	if (typeof value === 'boolean') return `<b v="${value ? 1 : 0}"/>`;
	if (!Number.isFinite(value)) throw new Error('The PivotTable cache contains an invalid value.');
	return `<n v="${value}"/>`;
}

export function pivotCacheDefinitionXml(
	pivot: PivotTable,
	workbook: Workbook,
	data: PivotCacheData,
	recordRel: string,
): string {
	const source = sparklineDataRange(workbook, 0, pivot.source);
	const iso = (serial: number) => serialToIso(serial, workbook.date1904);
	const derived = derivedFields(pivot, workbook.date1904);
	const fields = pivot.fields
		.map((name, index) => {
			const items = data.items[index]!;
			if (pivot.dateFields?.includes(name) && items.every((v) => typeof v === 'number')) {
				const serials = items as number[];
				// The earliest date to a day after the latest whole-day date, found in a loop: a
				// spread would pass one argument per date and overflow on large caches.
				const range = dateGroupRange({}, serials);
				const parent = derived.filter((entry) => entry.group.field === name).at(-1);
				const link = parent ? `<fieldGroup par="${parent.index}"/>` : '';
				return `<cacheField${attrs({ name: encodeEscapes(name), numFmtId: 14 })}><sharedItems${attrs({ containsSemiMixedTypes: false, containsNonDate: false, containsDate: true, containsString: false, minDate: iso(range.start), maxDate: iso(range.end), count: serials.length })}>${serials.map((serial) => `<d v="${iso(serial)}"/>`).join('')}</sharedItems>${link}</cacheField>`;
			}
			const types = new Set(
				items.map((value) => (value === null || value === '' ? 'blank' : typeof value)),
			);
			const nums = items.filter((value): value is number => typeof value === 'number');
			// Excel classifies boolean/error items with strings, and blanks are not a mixed type.
			const text = items.some((item) => item !== null && item !== '' && typeof item !== 'number');
			return `<cacheField${attrs({ name: encodeEscapes(name), numFmtId: 0 })}><sharedItems${attrs({ count: items.length, containsString: !nums.length || text, containsNumber: !!nums.length, containsInteger: nums.length > 0 && nums.every(Number.isInteger), containsBlank: types.has('blank'), containsMixedTypes: nums.length > 0 && text, containsSemiMixedTypes: !nums.length || types.size > 1 })}>${items.map(valueXml).join('')}</sharedItems>${groupXml(pivot, name, index)}</cacheField>`;
		})
		.join('');
	const extra = derived.map((entry) => derivedFieldXml(pivot, entry, workbook.date1904)).join('');
	return `${XML_HEADER}<pivotCacheDefinition xmlns="${NS.x}" xmlns:r="${NS.r}" r:id="${recordRel}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${pivot.records.length}"><cacheSource type="worksheet"><worksheetSource${attrs({ ref: formatRange(source.range), sheet: workbook.sheets[source.sheet]!.name })}/></cacheSource><cacheFields count="${pivot.fields.length + derived.length}">${fields}${extra}</cacheFields>${slicerExt(pivot)}</pivotCacheDefinition>`;
}

interface DerivedCacheField extends DerivedDateField {
	autoStart: boolean;
	autoEnd: boolean;
}

/** The derived date fields with their resolved ranges and whether Excel picks start and end. */
function derivedFields(pivot: PivotTable, date1904: boolean): DerivedCacheField[] {
	const resolved = { ...pivot, date1904, dateGroups: resolveDateGroups({ ...pivot, date1904 }) };
	return derivedDateFields(resolved).map((entry) => {
		const chosen = pivot.dateGroups!.find((group) => group.field === entry.group.field)!;
		return { ...entry, autoStart: chosen.start === undefined, autoEnd: chosen.end === undefined };
	});
}

/** A derived field of a date group: no records, only the range and the group items. */
function derivedFieldXml(pivot: PivotTable, entry: DerivedCacheField, date1904: boolean): string {
	const range = { start: entry.group.start!, end: entry.group.end! };
	const days = dateGroupDays(entry.group);
	const labels = dateUnitLabels(entry.unit, range, days, date1904, entry.group);
	const base = pivot.fields.indexOf(entry.group.field);
	return `<cacheField${attrs({ name: encodeEscapes(entry.name), numFmtId: 0, databaseField: false })}><fieldGroup base="${base}"><rangePr${attrs({ autoStart: entry.autoStart ? undefined : false, autoEnd: entry.autoEnd ? undefined : false, groupBy: entry.unit, startDate: serialToIso(range.start, date1904), endDate: serialToIso(range.end, date1904), groupInterval: days > 1 ? days : undefined })}/><groupItems count="${labels.length}">${labels.map((label) => `<s v="${escapeAttr(label)}"/>`).join('')}</groupItems></fieldGroup></cacheField>`;
}

/** `fieldGroup` of a number-grouped field: the range and the bucket labels. */
function groupXml(pivot: PivotTable, name: string, index: number): string {
	const group = groupOf(pivot, name);
	if (!group) return '';
	const labels = numberGroupLabels(group);
	return `<fieldGroup base="${index}"><rangePr autoStart="0" autoEnd="0" startNum="${group.start}" endNum="${group.end}" groupInterval="${group.by}"/><groupItems count="${labels.length}">${labels.map((label) => `<s v="${escapeAttr(label)}"/>`).join('')}</groupItems></fieldGroup>`;
}

/** The identity slicers find a cache by (`x14:pivotCacheDefinition/@pivotCacheId`). */
function slicerExt(pivot: PivotTable): string {
	if (pivot.slicerCacheId === undefined) return '';
	const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
	return `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${x14}"><x14:pivotCacheDefinition pivotCacheId="${pivot.slicerCacheId}"/></ext></extLst>`;
}

export function pivotCacheRecordsXml(pivot: PivotTable, data: PivotCacheData): string {
	const records = pivot.records
		.map(
			(record) =>
				`<r>${pivot.fields.map((_field, index) => `<x v="${data.indices[index]!.get(pivotValueKey(record[index] ?? null))}"/>`).join('')}</r>`,
		)
		.join('');
	return `${XML_HEADER}<pivotCacheRecords xmlns="${NS.x}" count="${pivot.records.length}">${records}</pivotCacheRecords>`;
}
