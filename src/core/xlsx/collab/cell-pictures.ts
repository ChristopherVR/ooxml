// In-cell pictures (Place in Cell) in a shared workbook. A picture cell's entry names its picture
// by content, `p: { h: <sha-256 hex>, a?: <alt text> }`, and the bytes live once per distinct
// picture in the top-level `xlsx:media` map, `h -> { t: <MIME type>, b: <bytes> }`, so a picture
// placed in many cells (or by two peers at once) is stored once. Entries are immutable and are
// never deleted (like the shared formats), so an undo or a concurrent delete can never take bytes
// another cell still names; bytes of pictures no cell uses any more stay in the room.
//
// Only PNG, JPEG and GIF pictures whose header the raster check accepts, of at most
// `MAX_SHARED_CELL_PICTURE_BYTES`, are shared. A larger picture (one loaded from a file: the editor
// refuses to place one while shared) reaches peers as the cell's `#VALUE!` fallback only.
import type * as Y from 'yjs';
import { sha256 } from '../../digest/sha256';
import { newMediaPart } from '../edit/charts';
import type { Cell, Workbook } from '../model';
import {
	MAX_SHARED_CELL_PICTURE_BYTES,
	cellPictureInfo,
	embeddedPicture,
	type EmbeddedCellPicture,
} from '../rich-values';
import { isRecord } from './schema';

export { MAX_SHARED_CELL_PICTURE_BYTES };

/** A picture cell's reference: the content hash of its bytes and its alt text. */
export interface PictureRef {
	h: string;
	a?: string;
}

const hashes = new WeakMap<Uint8Array, string>();

function hashOf(bytes: Uint8Array): string {
	let hash = hashes.get(bytes);
	if (hash === undefined) {
		hash = Array.from(sha256(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');
		hashes.set(bytes, hash);
	}
	return hash;
}

/** Local media parts by content hash, per workbook, so a picture read twice is stored once. */
const localParts = new WeakMap<Workbook, Map<string, string>>();

/** Whether `bytes` is a picture a room may carry. */
export const shareableCellPicture = (bytes: Uint8Array): boolean =>
	bytes.length <= MAX_SHARED_CELL_PICTURE_BYTES && cellPictureInfo(bytes) !== undefined;

/**
 * Translates between a workbook's picture cells and their shared references. Writers publish the
 * bytes of the pictures they reference; readers never write.
 */
export class PictureBridge {
	private readonly parts: Map<string, string>;

	constructor(
		private readonly workbook: Workbook,
		private readonly media: Y.Map<unknown>,
		private readonly writable: boolean,
	) {
		let parts = localParts.get(workbook);
		if (!parts) localParts.set(workbook, (parts = new Map()));
		this.parts = parts;
	}

	/** The shared reference of a picture cell, or undefined for any other cell. */
	refOf = (cell: Cell): PictureRef | undefined => {
		const picture = embeddedPicture(this.workbook, cell);
		const bytes = picture && this.workbook.source?.parts.get(picture.partName);
		if (!picture || !bytes || !shareableCellPicture(bytes)) return undefined;
		const h = hashOf(bytes);
		this.parts.set(h, picture.partName);
		if (this.writable && !this.media.has(h)) {
			const info = cellPictureInfo(bytes);
			if (info) this.media.set(h, { t: info.mimeType, b: bytes.slice() });
		}
		return picture.altText ? { h, a: picture.altText } : { h };
	};

	/**
	 * The local picture of a shared reference: the bytes are checked as a raster picture of the type
	 * they claim and stored as a media part of this workbook. Undefined when the reference is
	 * malformed or its bytes are missing or refused.
	 */
	pictureOf = (raw: unknown): EmbeddedCellPicture | undefined => {
		if (!isRecord(raw) || typeof raw.h !== 'string' || !/^[0-9a-f]{64}$/.test(raw.h)) return;
		if (raw.a !== undefined && (typeof raw.a !== 'string' || raw.a.length > 1024)) return;
		const altText = (raw.a as string | undefined) ?? '';
		const known = this.parts.get(raw.h);
		const local = known && this.workbook.source?.parts.get(known);
		if (known && local && hashOf(local) === raw.h) return { partName: known, altText };
		const entry = this.media.get(raw.h);
		if (!isRecord(entry) || !(entry.b instanceof Uint8Array) || typeof entry.t !== 'string')
			return undefined;
		const bytes = entry.b;
		if (!shareableCellPicture(bytes) || hashOf(bytes) !== raw.h) return undefined;
		const info = cellPictureInfo(bytes);
		if (!info || info.mimeType !== entry.t) return undefined;
		const partName = newMediaPart(this.workbook, info.mimeType);
		this.workbook.source ??= { parts: new Map() };
		this.workbook.source.parts.set(partName, bytes.slice());
		this.parts.set(raw.h, partName);
		return { partName, altText };
	};
}
