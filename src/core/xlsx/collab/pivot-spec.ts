// PivotTables as keyed entries of `xlsx:pivots` (see keyed.ts): minted key -> the whole
// PivotTable (definition, layout, filters, groups and the cached source fields and records), last
// writer wins per PivotTable. The output cells travel as ordinary cells, so peers see what the
// author's refresh wrote. Part names and the cache id belong to one peer's package and are never
// shared; a PivotTable a peer does not have is written as new parts on its save. A PivotTable
// whose shared form is over `MAX_SHARED_PIVOT_CHARS` is not shared: it stays on its own peer.
import type { CellValue, PivotTable } from '../model';
import { PIVOT_AGGREGATES } from '../pivot/data';
import { MAX_SHARED_PIVOT_CHARS, pivotShareable, sharedPivotForm } from '../pivot/share-limit';
import type { PartSpec } from './keyed';
import { isRecord } from './schema';

const MAX_FIELDS = 16_384;
const MAX_RECORDS = 1_048_576;
const isText = (v: unknown, max: number): v is string => typeof v === 'string' && v.length <= max;
const isIndex = (v: unknown, max: number): v is number =>
	typeof v === 'number' && Number.isInteger(v) && v >= 0 && v <= max;

function isRange(raw: unknown): boolean {
	return (
		isRecord(raw) &&
		isRecord(raw.start) &&
		isRecord(raw.end) &&
		isIndex(raw.start.row, 1_048_575) &&
		isIndex(raw.start.col, 16_383) &&
		isIndex(raw.end.row, 1_048_575) &&
		isIndex(raw.end.col, 16_383)
	);
}

function isCellValue(value: unknown): value is CellValue {
	if (value === null || typeof value === 'string' || typeof value === 'boolean') return true;
	if (typeof value === 'number') return Number.isFinite(value);
	return isRecord(value) && typeof value.error === 'string' && value.error.length <= 16;
}

const aggregates: readonly string[] = PIVOT_AGGREGATES;

/** The top-level shape every PivotTable needs; nested options are checked by the engine's use. */
function isPivot(raw: Record<string, unknown>): boolean {
	const { fields, records } = raw;
	return (
		isText(raw.name, 255) &&
		isText(raw.source, 1024) &&
		isText(raw.rowField, 255) &&
		isText(raw.valueField, 255) &&
		isText(raw.valueCaption, 255) &&
		typeof raw.aggregate === 'string' &&
		aggregates.includes(raw.aggregate) &&
		isRange(raw.location) &&
		Array.isArray(fields) &&
		fields.length <= MAX_FIELDS &&
		fields.every((field) => isText(field, 255)) &&
		Array.isArray(records) &&
		records.length <= MAX_RECORDS &&
		records.every(
			(record) =>
				Array.isArray(record) && record.length <= fields.length && record.every(isCellValue),
		) &&
		(raw.sourceXml === undefined || isText(raw.sourceXml, 1_048_576))
	);
}

export const pivotSpec: PartSpec<PivotTable> = {
	encode: (pivot) => (pivotShareable(pivot) ? sharedPivotForm(pivot) : undefined),
	decode: (raw) => {
		if (!isRecord(raw) || !isPivot(raw)) return undefined;
		if (JSON.stringify(raw).length > MAX_SHARED_PIVOT_CHARS) return undefined;
		return JSON.parse(JSON.stringify(raw)) as PivotTable;
	},
	natural: (encoded) => {
		const name = isRecord(encoded) ? encoded.name : undefined;
		return typeof name === 'string' ? name : undefined;
	},
	// The local package's parts stay with the same PivotTable; a remote edit marks it for rewriting.
	adopt: (decoded, local) => {
		if (local?.partName === undefined) return;
		decoded.partName = local.partName;
		if (local.cacheId !== undefined) decoded.cacheId = local.cacheId;
		decoded.dirty = true;
	},
};
