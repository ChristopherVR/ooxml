/**
 * Embedding permissions of a presentation's font parts, read from the font's
 * OS/2 `fsType` (or the copy an EOT header carries) without decoding the part.
 * The decisions themselves are the format-neutral ones in `text/truetype/embedding`.
 *
 * A part whose permission cannot be read (WOFF/WOFF2 payloads compress their
 * tables, a font without an OS/2 table) is treated as unrestricted: there is
 * nothing to honour, and refusing it would hide fonts that decks rely on.
 */
import {
	embeddingPermission,
	mayDisplayEmbeddedFont,
	mayEmbedFontInEditableDocument,
	readSfntFsType,
} from '../../../text/truetype/embedding';
import type { PptxEmbeddedFont } from '../types';
import { readEotFsType } from './eot-parser';
import { deobfuscateFont } from './font-deobfuscation';

/** The XOR layer of an obfuscated part covers only its first 32 bytes. */
const OBFUSCATED_HEADER_BYTES = 32;

/**
 * `fsType` of a stored font part: from the EOT header, or from the OS/2 table
 * of an obfuscated sfnt, de-obfuscating only its header (`guid` is the part's
 * obfuscation key). `undefined` when it cannot be read.
 */
export function readEmbeddedFontPartFsType(
	partBytes: Uint8Array,
	guid: string | undefined,
): number | undefined {
	const eot = readEotFsType(partBytes);
	if (eot !== undefined) {
		return eot;
	}
	if (!guid) {
		return readSfntFsType(partBytes);
	}
	const header = deobfuscateFont(
		partBytes.subarray(0, Math.min(partBytes.length, OBFUSCATED_HEADER_BYTES)),
		guid,
	);
	return readSfntFsType(partBytes, header);
}

/** Whether a face with this `fsType` may be activated for display. */
export function mayDisplayFontWithFsType(fsType: number | undefined): boolean {
	return fsType === undefined || mayDisplayEmbeddedFont(embeddingPermission(fsType));
}

/**
 * Whether clear font bytes may be newly embedded into a saved presentation
 * (installable or editable, outlines allowed). Only new parts are gated: a
 * part the deck already carries is always written back unchanged.
 */
export function mayNewlyEmbedFontData(font: PptxEmbeddedFont, fontData: Uint8Array): boolean {
	const fsType = readSfntFsType(fontData) ?? font.fsType;
	return fsType === undefined || mayEmbedFontInEditableDocument(embeddingPermission(fsType));
}
