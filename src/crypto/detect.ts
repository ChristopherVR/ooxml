/**
 * Recognises an encrypted OOXML package: a compound file (CFB, magic D0 CF 11 E0 A1 B1 1A E1)
 * holding `EncryptionInfo` and `EncryptedPackage` streams, as opposed to a plain OOXML zip or a
 * legacy binary document (.xls, .doc, .ppt), which is a compound file too.
 *
 * Moved from `src/pptx/core/utils/encryption-detection.ts` (see PROVENANCE.md);
 * {@link isEncryptedOoxmlPackage} and {@link openEncryptedPackage} are new.
 */

import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';

/** Compound File Binary magic signature. */
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

/** Bytes as a Uint8Array (no copy). */
export const toBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
	input instanceof Uint8Array ? input : new Uint8Array(input);

/** An ArrayBuffer holding exactly `bytes` (copied only when it is a view of a larger buffer). */
export function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
		return bytes.buffer as ArrayBuffer;
	}
	return bytes.slice().buffer;
}

/** Whether the bytes start with the compound-file magic. */
export function isCompoundFile(input: Uint8Array | ArrayBuffer): boolean {
	const bytes = toBytes(input);
	return bytes.length >= 8 && OLE_MAGIC.every((value, index) => bytes[index] === value);
}

/** The two streams of an encrypted package. */
export interface EncryptedPackageStreams {
	encryptionInfo: Uint8Array;
	encryptedPackage: Uint8Array;
}

/**
 * The `EncryptionInfo` and `EncryptedPackage` streams of an encrypted OOXML package, or
 * `undefined` when the bytes are not a compound file or lack either stream (a legacy binary
 * document). Never throws.
 */
export function openEncryptedPackage(
	input: Uint8Array | ArrayBuffer,
): EncryptedPackageStreams | undefined {
	const bytes = toBytes(input);
	if (!isCompoundFile(bytes)) return undefined;
	try {
		const file = parseOle2(toArrayBuffer(bytes));
		// A legacy .xls/.doc/.ppt has no EncryptionInfo: answer before reading any large stream.
		if (!file.entries.some((entry) => entry.name === 'EncryptedPackage')) return undefined;
		const encryptionInfo = file.getStream('EncryptionInfo');
		const encryptedPackage = encryptionInfo && file.getStream('EncryptedPackage');
		return encryptionInfo && encryptedPackage ? { encryptionInfo, encryptedPackage } : undefined;
	} catch {
		return undefined;
	}
}

/**
 * Whether the bytes are an encrypted OOXML package (.xlsx, .docx, .pptx and their macro and
 * template variants saved with a password to open). A legacy .xls/.doc/.ppt is `false`.
 */
export function isEncryptedOoxmlPackage(input: Uint8Array | ArrayBuffer): boolean {
	const bytes = toBytes(input);
	if (!isCompoundFile(bytes)) return false;
	try {
		const names = new Set(parseOle2(toArrayBuffer(bytes)).entries.map((entry) => entry.name));
		return names.has('EncryptionInfo') && names.has('EncryptedPackage');
	} catch {
		return false;
	}
}

/** A coarse container sniff: zip, compound file (possibly an encrypted package) or unknown. */
export type FileFormatDetection =
	| { format: 'zip'; encrypted: false }
	| { format: 'ole'; encrypted: true }
	| { format: 'unknown'; encrypted: false };

/**
 * Detect the container format from the magic bytes. A compound file is reported as
 * `encrypted` without opening it (use {@link isEncryptedOoxmlPackage} to tell an encrypted
 * package from a legacy binary document). Must be called before parsing as a zip.
 */
export function detectFileFormat(data: ArrayBuffer | Uint8Array): FileFormatDetection {
	const bytes = toBytes(data);
	if (bytes.length < 8) return { format: 'unknown', encrypted: false };
	if (isCompoundFile(bytes)) return { format: 'ole', encrypted: true };
	if (bytes[0] === 0x50 && bytes[1] === 0x4b) return { format: 'zip', encrypted: false };
	return { format: 'unknown', encrypted: false };
}

/**
 * Thrown when an encrypted file is detected where a plain package was expected. Callers can
 * check `isEncrypted` (or `instanceof`) to tell it from generic parse failures.
 */
export class EncryptedFileError extends Error {
	public readonly isEncrypted = true;

	public constructor(message: string) {
		super(message);
		this.name = 'EncryptedFileError';
	}
}
