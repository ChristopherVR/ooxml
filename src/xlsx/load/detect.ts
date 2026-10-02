// Spreadsheet format detection, loading and saving: .xlsx/.xlsm (OPC zip), Excel 97-2003 .xls
// (OLE2 compound file) and CSV text. Detection sniffs content; a file name only distinguishes
// macro-enabled packages and never overrides what the bytes are.
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { csvToWorkbook, sheetNameFromFileName, sheetToCsv } from './csv.js';
import { loadLegacyXls } from './legacy-xls.js';

export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xls' | 'csv';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

const toBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
	input instanceof Uint8Array ? input : new Uint8Array(input);

const startsWith = (bytes: Uint8Array, magic: readonly number[]): boolean =>
	magic.every((value, index) => bytes[index] === value);

/**
 * Whether a zip holds a VBA project part (`xl/vbaProject.bin`), read from the uncompressed file
 * names in the central directory so detection stays synchronous.
 */
function zipHasVbaProject(bytes: Uint8Array): boolean {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let end = -1;
	for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at--) {
		if (view.getUint32(at, true) === 0x06054b50) {
			end = at;
			break;
		}
	}
	if (end < 0) return false;
	const count = view.getUint16(end + 10, true);
	let at = view.getUint32(end + 16, true);
	const decoder = new TextDecoder();
	for (let i = 0; i < count && at + 46 <= bytes.length; i++) {
		if (view.getUint32(at, true) !== 0x02014b50) return false;
		const nameLength = view.getUint16(at + 28, true);
		const extraLength = view.getUint16(at + 30, true);
		const commentLength = view.getUint16(at + 32, true);
		const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
		if (name.toLowerCase() === 'xl/vbaproject.bin') return true;
		at += 46 + nameLength + extraLength + commentLength;
	}
	return false;
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
	if (startsWith(bytes, CFB_MAGIC)) return 'xls';
	if (startsWith(bytes, ZIP_MAGIC)) {
		if (fileName && /\.xl[st]m$/i.test(fileName)) return 'xlsm';
		return zipHasVbaProject(bytes) ? 'xlsm' : 'xlsx';
	}
	if (looksLikeText(bytes)) return 'csv';
	throw new Error('Unsupported file. Open an Excel workbook (.xlsx, .xlsm, .xls) or a CSV file.');
}

/** Decodes CSV bytes: UTF-8 (with or without BOM) or UTF-16 with a BOM. */
export function decodeText(bytes: Uint8Array): string {
	if (bytes[0] === 0xff && bytes[1] === 0xfe)
		return new TextDecoder('utf-16le').decode(bytes.subarray(2));
	if (bytes[0] === 0xfe && bytes[1] === 0xff)
		return new TextDecoder('utf-16be').decode(bytes.subarray(2));
	const text = new TextDecoder('utf-8').decode(bytes);
	return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

export interface LoadWorkbookOptions {
	/** The file name, used for the CSV sheet name and to recognise macro-enabled packages. */
	fileName?: string;
	/** CSV field delimiter; detected when omitted. */
	delimiter?: string;
}

/** Loads any supported spreadsheet: .xlsx, .xlsm, .xltx, .xls or CSV. */
export async function loadWorkbook(
	input: Uint8Array | ArrayBuffer,
	options: LoadWorkbookOptions = {},
): Promise<Workbook> {
	const bytes = toBytes(input);
	const format = detectWorkbookFormat(bytes, options.fileName);
	if (format === 'xls') return loadLegacyXls(bytes);
	if (format === 'csv') {
		return csvToWorkbook(decodeText(bytes), {
			sheetName: sheetNameFromFileName(options.fileName),
			...(options.delimiter === undefined ? {} : { delimiter: options.delimiter }),
		});
	}
	const workbook = await loadXlsx(bytes);
	if (format === 'xlsm' && workbook.format === 'xlsx') workbook.format = 'xlsm';
	return workbook;
}

/**
 * Serializes a workbook. `'xlsx'` writes the whole workbook as an OOXML package (an .xls or CSV
 * source is converted: the result is always .xlsx/.xlsm content); `'csv'` writes one sheet
 * (`sheetIndex`, default the active sheet) as UTF-8 CSV with a byte-order mark, as Excel's
 * "CSV UTF-8" does.
 */
export async function saveWorkbook(
	workbook: Workbook,
	format: 'xlsx' | 'csv',
	sheetIndex?: number,
): Promise<Uint8Array> {
	if (format === 'xlsx') return saveXlsx(workbook);
	const csv = sheetToCsv(workbook, sheetIndex ?? workbook.activeSheet);
	return new TextEncoder().encode(`﻿${csv}`);
}
