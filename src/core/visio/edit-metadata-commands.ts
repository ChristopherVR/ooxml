import { fail } from './package-common';
import { VISIO_RASTER_IMAGE_LIMITS } from './media';

/** Longest ScreenTip (Comment cell) and hyperlink field accepted by the editor. */
export const VISIO_METADATA_TEXT_LIMIT = 4096;
/** Largest picture accepted for insertion; the reader shares the same raster ceiling. */
export const VISIO_PICTURE_MAX_BYTES = VISIO_RASTER_IMAGE_LIMITS.maxImageBytes;

export interface VisioHyperlinkFields {
	/** External address (URL or path); empty for a link to a page or shape only. */
	address: string;
	/** Page or page/shape target inside the drawing; empty for none. */
	subAddress: string;
	description: string;
}
/** Add, replace or remove (`hyperlink: null`) one local Hyperlink section row. */
export interface VisioShapeHyperlinkEdit {
	type: 'set-shape-hyperlink';
	pageId: string;
	shapeId: string;
	/** Existing row name (`Row_1`) to replace or remove; omitted adds a new row. */
	row?: string;
	hyperlink: VisioHyperlinkFields | null;
}
/** Set the shape's Comment cell, which Visio shows as its ScreenTip; empty text removes it. */
export interface VisioShapeScreenTipEdit {
	type: 'set-shape-screentip';
	pageId: string;
	shapeId: string;
	text: string;
}
/** Embed a PNG, JPEG or GIF as a new Foreign shape. Drawing inches, bottom-left origin. */
export interface VisioPictureInsertEdit {
	type: 'insert-picture';
	pageId: string;
	shapeId: string;
	x: number;
	y: number;
	width: number;
	height: number;
	image: Uint8Array;
}
export type VisioMetadataEdit = VisioShapeHyperlinkEdit | VisioShapeScreenTipEdit;

export const isVisioMetadataEdit = (edit: { type: string }): edit is VisioMetadataEdit =>
	edit.type === 'set-shape-hyperlink' || edit.type === 'set-shape-screentip';

const invalidText =
	/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u;
function field(value: unknown, singleLine: boolean): string {
	if (typeof value !== 'string' || value.length > VISIO_METADATA_TEXT_LIMIT)
		fail('INVALID_EDIT', 'Link and ScreenTip text must be strings within the length limit.');
	if (invalidText.test(value) || (singleLine && /[\r\n\t]/.test(value)))
		fail('INVALID_EDIT_TEXT', 'Link or ScreenTip text contains invalid characters.');
	return value;
}
function target(edit: { pageId: unknown; shapeId: unknown }): { pageId: string; shapeId: string } {
	if (
		typeof edit.pageId !== 'string' ||
		!edit.pageId ||
		edit.pageId.length > 256 ||
		typeof edit.shapeId !== 'string' ||
		!/^[1-9]\d{0,9}$/.test(edit.shapeId) ||
		Number(edit.shapeId) > 0xffffffff
	)
		fail('INVALID_EDIT', 'Invalid edit shape target.');
	return { pageId: edit.pageId as string, shapeId: edit.shapeId as string };
}

/** Copy and validate a metadata command; host properties never reach the transaction. */
export function snapshotMetadataEdit(edit: VisioMetadataEdit): VisioMetadataEdit {
	const base = target(edit);
	if (edit.type === 'set-shape-screentip')
		return { ...base, type: edit.type, text: field(edit.text, true) };
	if (
		edit.row !== undefined &&
		(typeof edit.row !== 'string' || !/^[A-Za-z_]\w{0,63}$/.test(edit.row))
	)
		fail('INVALID_EDIT', 'Hyperlink row names must be ShapeSheet row names.');
	const link = edit.hyperlink;
	if (link === null) {
		if (edit.row === undefined) fail('INVALID_EDIT', 'Removing a hyperlink requires its row.');
		return { ...base, type: edit.type, row: edit.row!, hyperlink: null };
	}
	if (!link || typeof link !== 'object') fail('INVALID_EDIT', 'Invalid hyperlink fields.');
	const hyperlink = {
		address: field(link.address, true).trim(),
		subAddress: field(link.subAddress, true).trim(),
		description: field(link.description, true),
	};
	if (!hyperlink.address && !hyperlink.subAddress)
		fail('INVALID_EDIT', 'A hyperlink needs an address or a page target.');
	return {
		...base,
		type: edit.type,
		...(edit.row === undefined ? {} : { row: edit.row }),
		hyperlink,
	};
}

export function snapshotPictureInsert(edit: VisioPictureInsertEdit): VisioPictureInsertEdit {
	const base = target(edit);
	const values = [edit.x, edit.y, edit.width, edit.height];
	if (
		values.some((value) => typeof value !== 'number' || !Number.isFinite(value)) ||
		values.some((value) => Math.abs(value) > 1e6) ||
		!(edit.width > 0) ||
		!(edit.height > 0)
	)
		fail('INVALID_EDIT', 'Pictures need finite drawing inches and positive dimensions.');
	if (!(edit.image instanceof Uint8Array) || !edit.image.byteLength)
		fail('INVALID_EDIT', 'Picture bytes must be a non-empty Uint8Array.');
	if (edit.image.byteLength > VISIO_PICTURE_MAX_BYTES)
		fail('LIMIT_PICTURE', 'Picture exceeds the 8 MiB insertion limit.');
	return {
		...base,
		type: 'insert-picture',
		x: edit.x,
		y: edit.y,
		width: edit.width,
		height: edit.height,
		image: new Uint8Array(edit.image),
	};
}
