import { fail } from './package-common';
import { target } from './edit-metadata-commands';

/** Largest embedded workbook or preview accepted for a chart object. */
export const VISIO_CHART_OBJECT_MAX_BYTES = 8 * 1024 * 1024;

/**
 * Embed an Excel workbook holding a chart as an OLE object, as Visio's Insert > Chart does: a
 * Foreign shape whose ForeignData is an Object, the workbook under visio/embeddings and an
 * optional EMF preview that Visio shows until the object is activated. Drawing inches,
 * bottom-left origin.
 */
export interface VisioChartObjectInsertEdit {
	type: 'insert-chart-object';
	pageId: string;
	shapeId: string;
	x: number;
	y: number;
	width: number;
	height: number;
	/** The .xlsx package (its first chart is the object). */
	workbook: Uint8Array;
	/** An enhanced metafile of the chart; omitted, Visio draws the object blank until activated. */
	preview?: Uint8Array;
}

/** Replace the workbook and preview of an existing chart object; placement is unchanged. */
export interface VisioChartObjectReplaceEdit {
	type: 'replace-chart-object';
	pageId: string;
	shapeId: string;
	workbook: Uint8Array;
	preview?: Uint8Array;
}

export type VisioChartObjectEdit = VisioChartObjectInsertEdit | VisioChartObjectReplaceEdit;

export const isVisioChartObjectEdit = (edit: { type: string }): edit is VisioChartObjectEdit =>
	edit.type === 'insert-chart-object' || edit.type === 'replace-chart-object';

function bytesOf(value: unknown, what: string, required: boolean): Uint8Array | undefined {
	if (value === undefined && !required) return undefined;
	if (!(value instanceof Uint8Array) || !value.byteLength)
		fail('INVALID_EDIT', `${what} must be a non-empty Uint8Array.`);
	if (value.byteLength > VISIO_CHART_OBJECT_MAX_BYTES)
		fail('LIMIT_CHART_OBJECT', `${what} exceeds the 8 MiB embedding limit.`);
	// Copied: a host must not be able to change the bytes after validation.
	return new Uint8Array(value);
}

const isZip = (bytes: Uint8Array) => bytes[0] === 0x50 && bytes[1] === 0x4b;
const isEmf = (bytes: Uint8Array) =>
	bytes.byteLength >= 44 &&
	new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true) === 1 &&
	new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(40, true) === 0x464d4520;

/** Copy and validate a chart object command; host properties never reach the transaction. */
export function snapshotChartObjectEdit(edit: VisioChartObjectEdit): VisioChartObjectEdit {
	const base = target(edit);
	const workbook = bytesOf(edit.workbook, 'The chart workbook', true)!;
	if (!isZip(workbook)) fail('INVALID_EDIT', 'The chart workbook is not an Open XML package.');
	const preview = bytesOf(edit.preview, 'The chart preview', false);
	if (preview && !isEmf(preview)) fail('INVALID_EDIT', 'The chart preview must be an EMF.');
	const previewPart = preview ? { preview } : {};
	if (edit.type === 'replace-chart-object')
		return { ...base, type: edit.type, workbook, ...previewPart };
	const values = [edit.x, edit.y, edit.width, edit.height];
	if (
		values.some((value) => typeof value !== 'number' || !Number.isFinite(value)) ||
		values.some((value) => Math.abs(value) > 1e6) ||
		!(edit.width > 0) ||
		!(edit.height > 0)
	)
		fail('INVALID_EDIT', 'Chart objects need finite drawing inches and positive dimensions.');
	return {
		...base,
		type: edit.type,
		x: edit.x,
		y: edit.y,
		width: edit.width,
		height: edit.height,
		workbook,
		...previewPart,
	};
}
