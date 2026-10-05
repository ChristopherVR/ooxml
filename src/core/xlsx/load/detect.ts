// Spreadsheet format detection, loading and saving: .xlsx/.xlsm (OPC zip), password-protected
// (encrypted) packages, Excel 97-2003 .xls (OLE2 compound file) and CSV text. Detection sniffs
// content; a file name only distinguishes macro-enabled and binary packages and never overrides
// what the bytes are.
import { PasswordRequiredError } from '../../crypto/errors.js';
import {
	decryptOoxmlPackage,
	encryptOoxmlPackage,
	isEncryptedOoxmlPackage,
	type EncryptionOptions,
} from '../../crypto/index.js';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { csvToWorkbook, sheetNameFromFileName, sheetToCsv } from './csv.js';
import { xlsbUnsupported } from './errors.js';
import { loadLegacyXls } from './legacy-xls.js';
import { decodeSpreadsheetText } from './text-decode.js';

/**
 * A detected spreadsheet format. `'encrypted'` is an OOXML package saved with a password to open
 * (a compound file holding EncryptionInfo and EncryptedPackage, not a legacy .xls); `'xlsb'` is
 * an Excel Binary Workbook, recognised so it can be refused honestly.
 */
export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xls' | 'csv' | 'encrypted' | 'xlsb';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

const toBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
	input instanceof Uint8Array ? input : new Uint8Array(input);

const startsWith = (bytes: Uint8Array, magic: readonly number[]): boolean =>
	magic.every((value, index) => bytes[index] === value);

/**
 * The lower-cased entry names of a zip, read from the uncompressed file names in the central
 * directory so detection stays synchronous (empty when the directory cannot be found).
 */
function zipEntryNames(bytes: Uint8Array): Set<string> {
	const names = new Set<string>();
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let end = -1;
	for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at--) {
		if (view.getUint32(at, true) === 0x06054b50) {
			end = at;
			break;
		}
	}
	if (end < 0) return names;
	const count = view.getUint16(end + 10, true);
	let at = view.getUint32(end + 16, true);
	const decoder = new TextDecoder();
	for (let i = 0; i < count && at + 46 <= bytes.length; i++) {
		if (view.getUint32(at, true) !== 0x02014b50) break;
		const nameLength = view.getUint16(at + 28, true);
		const extraLength = view.getUint16(at + 30, true);
		const commentLength = view.getUint16(at + 32, true);
		names.add(decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength)).toLowerCase());
		at += 46 + nameLength + extraLength + commentLength;
	}
	return names;
}

/** Text when the first bytes hold no NUL characters (UTF-16 text is recognised by its BOM). */
function looksLikeText(bytes: Uint8Array): boolean {
	if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))
		return true;
	const sample = bytes.subarray(0, 4096);
	return !sample.includes(0);
}

/** Sniffs a spreadsheet's format from its bytes (and, for macro-enabled packages, its name). */
export function detectWorkbookFormat(
	input: Uint8Array | ArrayBuffer,
	fileName?: string,
): WorkbookFormat {
	const bytes = toBytes(input);
	if (startsWith(bytes, CFB_MAGIC)) return isEncryptedOoxmlPackage(bytes) ? 'encrypted' : 'xls';
	if (startsWith(bytes, ZIP_MAGIC)) {
		const names = zipEntryNames(bytes);
		if (names.has('xl/workbook.bin')) return 'xlsb';
		if (fileName && /\.xl[st]m$/i.test(fileName)) return 'xlsm';
		return names.has('xl/vbaproject.bin') ? 'xlsm' : 'xlsx';
	}
	if (looksLikeText(bytes)) return 'csv';
	throw new Error('Unsupported file. Open an Excel workbook (.xlsx, .xlsm, .xls) or a CSV file.');
}

/**
 * Decodes CSV bytes: a UTF-16 or UTF-8 byte-order mark is honoured; otherwise strict UTF-8,
 * falling back to Windows-1252 (Excel's default "CSV (Comma delimited)" encoding).
 */
export function decodeText(bytes: Uint8Array): string {
	return decodeSpreadsheetText(bytes);
}

export interface LoadWorkbookOptions {
	/** The file name, used for the CSV sheet name and to recognise macro-enabled packages. */
	fileName?: string;
	/** CSV field delimiter; detected when omitted. */
	delimiter?: string;
	/**
	 * The password to open an encrypted (password-protected) .xlsx/.xlsm. Without it such a file
	 * throws `PasswordRequiredError` (`code: 'password-required'`); a wrong one throws
	 * `IncorrectPasswordError` (`code: 'incorrect-password'`).
	 */
	password?: string;
	/**
	 * Import CSV fields starting with `=` as formulas, as Excel does. Off by default: a CSV from an
	 * untrusted source would otherwise inject live formulas (`=HYPERLINK(...)`).
	 */
	formulas?: boolean;
}

const ENCRYPTED_WARNING =
	'This workbook was opened with a password. Saving writes it without the password unless one is set again when saving.';

/** The plain package of an encrypted workbook. */
async function decryptWorkbook(bytes: Uint8Array, password: string | undefined) {
	if (password === undefined) {
		throw new PasswordRequiredError(
			'This workbook is password protected. Enter the password to open it.',
		);
	}
	const plain = await decryptOoxmlPackage(bytes, password);
	if (!startsWith(plain, ZIP_MAGIC)) {
		throw new Error('The decrypted file is not an Excel workbook package.');
	}
	return plain;
}

/** Loads any supported spreadsheet: .xlsx, .xlsm, .xltx (also password-protected), .xls or CSV. */
export async function loadWorkbook(
	input: Uint8Array | ArrayBuffer,
	options: LoadWorkbookOptions = {},
): Promise<Workbook> {
	let bytes = toBytes(input);
	let format = detectWorkbookFormat(bytes, options.fileName);
	const encrypted = format === 'encrypted';
	if (encrypted) {
		bytes = await decryptWorkbook(bytes, options.password);
		format = detectWorkbookFormat(bytes, options.fileName);
	}
	if (format === 'xlsb') throw xlsbUnsupported();
	if (format === 'xls') return loadLegacyXls(bytes);
	if (format === 'csv') {
		return csvToWorkbook(decodeText(bytes), {
			sheetName: sheetNameFromFileName(options.fileName),
			...(options.delimiter === undefined ? {} : { delimiter: options.delimiter }),
			...(options.formulas === undefined ? {} : { formulas: options.formulas }),
		});
	}
	const workbook = await loadXlsx(bytes);
	if (format === 'xlsm' && workbook.format === 'xlsx') workbook.format = 'xlsm';
	if (encrypted) workbook.warnings.push(ENCRYPTED_WARNING);
	return workbook;
}

export interface SaveWorkbookOptions {
	/** The sheet a CSV save writes (default the active sheet). */
	sheetIndex?: number;
	/**
	 * Encrypts an `'xlsx'` save with this password to open (agile encryption, AES-256 and
	 * SHA-512, as Excel 2010 and later write). Not available for CSV.
	 */
	password?: string;
	/** Advanced encryption settings for a password save (scheme, key size, spin count). */
	encryption?: EncryptionOptions;
}

/**
 * Serializes a workbook. `'xlsx'` writes the whole workbook as an OOXML package (an .xls or CSV
 * source is converted: the result is always .xlsx/.xlsm content), encrypted when `password` is
 * given; `'csv'` writes one sheet (`sheetIndex`, default the active sheet) as UTF-8 CSV with a
 * byte-order mark, as Excel's "CSV UTF-8" does. A number as the third argument is the
 * `sheetIndex` (the original signature).
 */
export async function saveWorkbook(
	workbook: Workbook,
	format: 'xlsx' | 'csv',
	sheetIndexOrOptions?: number | SaveWorkbookOptions,
): Promise<Uint8Array> {
	const options: SaveWorkbookOptions =
		typeof sheetIndexOrOptions === 'number'
			? { sheetIndex: sheetIndexOrOptions }
			: (sheetIndexOrOptions ?? {});
	if (format === 'xlsx') {
		const bytes = await saveXlsx(workbook);
		if (options.password === undefined) return bytes;
		return encryptOoxmlPackage(bytes, options.password, options.encryption ?? {});
	}
	if (options.password !== undefined) {
		throw new Error('A CSV file cannot be password protected; save as .xlsx to set a password.');
	}
	const csv = sheetToCsv(workbook, options.sheetIndex ?? workbook.activeSheet);
	return new TextEncoder().encode(`﻿${csv}`);
}
