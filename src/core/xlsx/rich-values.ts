import { inspectRasterImage, type RasterImageInfo } from '../drawingml/raster-image';
import { type Cell, type CellError, type ErrorCode, isCellError, type Workbook } from './model';

export const RICH_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
export const RICH_REL_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel';
export const RICH_TYPES_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2';
export const RICH_METADATA = 'XLRICHVALUE';
export const RICH_METADATA_EXT = '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}';
export const RICH_PARTS = {
	data: {
		rel: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue',
		type: 'application/vnd.ms-excel.rdrichvalue+xml',
		file: 'rdrichvalue.xml',
		root: 'rvData',
		ns: RICH_NS,
	},
	structures: {
		rel: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure',
		type: 'application/vnd.ms-excel.rdrichvaluestructure+xml',
		file: 'rdrichvaluestructure.xml',
		root: 'rvStructures',
		ns: RICH_NS,
	},
	types: {
		rel: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes',
		type: 'application/vnd.ms-excel.rdrichvaluetypes+xml',
		file: 'rdRichValueTypes.xml',
		root: 'rvTypesInfo',
		ns: RICH_TYPES_NS,
	},
	relationships: {
		rel: 'http://schemas.microsoft.com/office/2022/10/relationships/richValueRel',
		type: 'application/vnd.ms-excel.richvaluerel+xml',
		file: 'richValueRel.xml',
		root: 'richValueRels',
		ns: RICH_REL_NS,
	},
} as const;

export const LOCAL_IMAGE_KEY = '_rvRel:LocalImageIdentifier';
/** [MS-XLSX] rich error codes for external image results. */
export const RICH_ERRORS: ReadonlyMap<ErrorCode, string> = new Map([
	['#CONNECT!', '9'],
	['#BLOCKED!', '10'],
	['#BUSY!', '14'],
]);

export interface RichValueField {
	readonly name: string;
	readonly type: string;
	readonly value: string;
}

export interface EmbeddedCellPicture {
	readonly partName: string;
	readonly altText: string;
}

export interface CellPicture extends RasterImageInfo {
	readonly bytes: Uint8Array;
	readonly altText: string;
	readonly sizing: 0 | 1 | 2 | 3;
	readonly height?: number;
	readonly width?: number;
}

export interface RichValueRecord {
	readonly type: string;
	readonly fields: readonly RichValueField[];
	readonly picture?: EmbeddedCellPicture;
}

export interface RichValueStore {
	readonly records: readonly (RichValueRecord | undefined)[];
	/** One-based worksheet vm to zero-based record index. */
	readonly byMetadata: ReadonlyMap<number, number>;
	/** Keys suppressed by the global and per-type ShowInCardView flags. */
	readonly hiddenCardFields: ReadonlyMap<string, ReadonlySet<string>>;
}

export function richError(store: RichValueStore | undefined, vm: number): ErrorCode | undefined {
	const index = store?.byMetadata.get(vm);
	const record = index === undefined ? undefined : store?.records[index];
	if (record?.type !== '_error') return undefined;
	const type = record.fields.find(
		(field) => field.name === 'errorType' && field.type === 'i',
	)?.value;
	for (const [code, id] of RICH_ERRORS) if (id === type) return code;
	return undefined;
}

export interface LinkedDataCard {
	readonly text: string;
	readonly fields: readonly RichValueField[];
}

/** Reserved host contract. A future refresh operation must explicitly opt in to this provider. */
export interface LinkedDataTypeProvider {
	refresh(value: RichValueRecord): Promise<RichValueRecord>;
}

const inspected = new WeakMap<Uint8Array, RasterImageInfo | undefined>();
export function cellPictureInfo(bytes: Uint8Array): RasterImageInfo | undefined {
	if (!inspected.has(bytes)) {
		try {
			inspected.set(bytes, inspectRasterImage(bytes));
		} catch {
			inspected.set(bytes, undefined);
		}
	}
	return inspected.get(bytes);
}

export function embeddedPicture(workbook: Workbook, cell: Cell): EmbeddedCellPicture | undefined {
	if (cell.inCellPicture) return cell.inCellPicture;
	if (!isCellError(cell.value)) return undefined;
	const store = workbook.richValues;
	const index =
		cell.valueMetadata === undefined ? undefined : store?.byMetadata.get(cell.valueMetadata);
	return index === undefined ? undefined : store?.records[index]?.picture;
}

const formulaPictures = new WeakMap<CellError, CellPicture>();

/** Formula results carry an Excel-compatible fallback for consumers that only handle scalar values. */
export function pictureResult(picture: CellPicture): CellError {
	const placeholder: CellError = { error: '#VALUE!' };
	formulaPictures.set(placeholder, picture);
	return placeholder;
}

export function pictureView(workbook: Workbook, cell: Cell): CellPicture | undefined {
	if (cell.value && typeof cell.value === 'object') {
		const formula = formulaPictures.get(cell.value);
		if (formula) return formula;
	}
	const picture = embeddedPicture(workbook, cell);
	const bytes = picture && workbook.source?.parts.get(picture.partName);
	const info = bytes && cellPictureInfo(bytes);
	return picture && bytes && info
		? { ...info, bytes, altText: picture.altText, sizing: 0 }
		: undefined;
}

export function richIndex(text: string | undefined): number | undefined {
	if (!text || !/^\d{1,10}$/.test(text)) return undefined;
	const index = Number(text);
	return index <= 0xffffffff ? index : undefined;
}

const cards = new WeakMap<RichValueRecord, LinkedDataCard | undefined>();
export function linkedDataCard(workbook: Workbook, cell: Cell): LinkedDataCard | undefined {
	const store = workbook.richValues;
	const index =
		cell.valueMetadata === undefined ? undefined : store?.byMetadata.get(cell.valueMetadata);
	if (!store || index === undefined || cell.inCellPicture) return undefined;
	const record = store.records[index];
	if (!record || !/^_(?:entity|linkedentity2?)$/.test(record.type)) return undefined;
	if (!cards.has(record)) {
		const core = record.fields.find(({ name, type }) => name === '%cvi' && type === 'r');
		const coreIndex = richIndex(core?.value);
		const nested =
			coreIndex !== undefined && coreIndex < index ? store.records[coreIndex] : undefined;
		const body = nested && /^_linkedentity2?core$/.test(nested.type) ? nested : record;
		const label = (fields: readonly RichValueField[]) =>
			fields.find(({ name, type }) => name === '_DisplayString' && type === 's')?.value;
		const text = label(record.fields) ?? label(body.fields);
		const hidden = store.hiddenCardFields;
		if (text === undefined) {
			cards.set(record, undefined);
			return undefined;
		}
		const fields: RichValueField[] = [];
		const globalHidden = hidden.get('');
		const typeHidden = hidden.get(body.type);
		for (const { name, type, value } of body.fields) {
			if (
				!name ||
				name.startsWith('_') ||
				name.startsWith('%') ||
				globalHidden?.has(name) ||
				typeHidden?.has(name) ||
				!['s', 'i', 'd', 'b', 'e'].includes(type)
			)
				continue;
			fields.push({ name: name.slice(0, 256), type, value: value.slice(0, 1024) });
			if (fields.length === 100) break;
		}
		cards.set(record, { text: text.slice(0, 1024), fields });
	}
	return cards.get(record);
}
