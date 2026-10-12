import { NS } from '../../xml/index';
import type { CellValue, PivotTable, Workbook } from '../model';
import { isCellError } from '../model';
import { formatRange } from '../address';
import { sparklineDataRange } from '../edit/sparkline-ranges';
import { XML_HEADER, attrs, encodeEscapes } from '../write/xml-out';
import { pivotValueKey } from './data';

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
	const fields = pivot.fields
		.map((name, index) => {
			const items = data.items[index]!;
			const types = new Set(
				items.map((value) => (value === null || value === '' ? 'blank' : typeof value)),
			);
			const nums = items.filter((value): value is number => typeof value === 'number');
			return `<cacheField${attrs({ name: encodeEscapes(name), numFmtId: 0 })}><sharedItems${attrs({ count: items.length, containsString: types.has('string'), containsNumber: !!nums.length, containsInteger: nums.length > 0 && nums.every(Number.isInteger), containsBlank: types.has('blank'), containsMixedTypes: types.size > 1, containsSemiMixedTypes: !nums.length || types.size > 1 })}>${items.map(valueXml).join('')}</sharedItems></cacheField>`;
		})
		.join('');
	return `${XML_HEADER}<pivotCacheDefinition xmlns="${NS.x}" xmlns:r="${NS.r}" r:id="${recordRel}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${pivot.records.length}"><cacheSource type="worksheet"><worksheetSource${attrs({ ref: formatRange(source.range), sheet: workbook.sheets[source.sheet]!.name })}/></cacheSource><cacheFields count="${pivot.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`;
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
