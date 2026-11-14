/**
 * Font embedding permissions: the OpenType `OS/2` table's `fsType` field, which the EOT header
 * also carries (offset 32). Format-neutral and DOM-free; every writer that embeds a font file
 * (a presentation's font parts, a PDF's subset fonts) decides through this one decoder.
 *
 * `fsType` bits 0-3 are the usage permission. A value of 0 there is "installable"; bit 1 is
 * "restricted licence", bit 2 "preview & print", bit 3 "editable". A font should set at most one
 * of them; when an older font sets several, the least restrictive one applies. Bit 8 forbids
 * subsetting and bit 9 allows only bitmaps to be embedded. The other bits are reserved.
 */

/** The usage permission a font's vendor grants for an embedded copy. */
export type EmbeddingLicence = 'installable' | 'editable' | 'preview-print' | 'restricted';

export interface EmbeddingPermission {
	readonly licence: EmbeddingLicence;
	/** Bit 8: the font may only be embedded whole. */
	readonly noSubsetting: boolean;
	/** Bit 9: only the bitmaps may be embedded, never the outlines. */
	readonly bitmapOnly: boolean;
}

const RESTRICTED = 0x0002;
const PREVIEW_PRINT = 0x0004;
const EDITABLE = 0x0008;
const NO_SUBSETTING = 0x0100;
const BITMAP_ONLY = 0x0200;

/** Decode an `fsType` value. */
export function embeddingPermission(fsType: number): EmbeddingPermission {
	const licence: EmbeddingLicence =
		fsType & EDITABLE
			? 'editable'
			: fsType & PREVIEW_PRINT
				? 'preview-print'
				: fsType & RESTRICTED
					? 'restricted'
					: 'installable';
	return {
		licence,
		noSubsetting: (fsType & NO_SUBSETTING) !== 0,
		bitmapOnly: (fsType & BITMAP_ONLY) !== 0,
	};
}

/** Whether an embedded copy may be used to display a document (every licence but restricted). */
export function mayDisplayEmbeddedFont(permission: EmbeddingPermission): boolean {
	return permission.licence !== 'restricted';
}

/**
 * Whether the whole font file (its outlines) may be newly embedded into an editable document:
 * installable and editable faces only, and not when the vendor allows bitmaps alone.
 */
export function mayEmbedFontInEditableDocument(permission: EmbeddingPermission): boolean {
	return (
		(permission.licence === 'installable' || permission.licence === 'editable') &&
		!permission.bitmapOnly
	);
}

/**
 * Whether a subset of the outlines may be embedded into a read-only document such as a PDF:
 * any licence but restricted, when subsetting and outlines are allowed.
 */
export function mayEmbedFontSubset(permission: EmbeddingPermission): boolean {
	return mayDisplayEmbeddedFont(permission) && !permission.noSubsetting && !permission.bitmapOnly;
}

const SFNT_VERSIONS = new Set([0x00010000, 0x4f54544f, 0x74727565]);
const TTC_TAG = 0x74746366; // 'ttcf'
const OS2_TAG = 0x4f532f32; // 'OS/2'

/**
 * Read `fsType` from an uncompressed sfnt (TrueType, OpenType, or the first font of a collection)
 * without parsing anything else. `header` replaces the leading bytes of `data` when given, so a
 * caller whose file is obfuscated only at its start (an OOXML font part) can pass the clear
 * header instead of copying the whole file. `undefined` when the bytes are not an sfnt (WOFF and
 * WOFF2 compress their tables) or have no `OS/2` table.
 */
export function readSfntFsType(data: Uint8Array, header?: Uint8Array): number | undefined {
	const byte = (offset: number): number | undefined =>
		header && offset < header.length ? header[offset] : data[offset];
	const u16 = (offset: number): number | undefined => {
		const hi = byte(offset);
		const lo = byte(offset + 1);
		return hi === undefined || lo === undefined ? undefined : (hi << 8) | lo;
	};
	const u32 = (offset: number): number | undefined => {
		const hi = u16(offset);
		const lo = u16(offset + 2);
		return hi === undefined || lo === undefined ? undefined : hi * 0x10000 + lo;
	};
	let base = 0;
	if (u32(0) === TTC_TAG) {
		const first = u32(12);
		if (first === undefined) return undefined;
		base = first;
	}
	const version = u32(base);
	if (version === undefined || !SFNT_VERSIONS.has(version)) return undefined;
	const numTables = u16(base + 4) ?? 0;
	for (let i = 0; i < numTables; i++) {
		const record = base + 12 + i * 16;
		const tag = u32(record);
		if (tag === undefined) return undefined;
		if (tag === OS2_TAG) {
			const offset = u32(record + 8);
			return offset === undefined ? undefined : u16(offset + 8);
		}
	}
	return undefined;
}
