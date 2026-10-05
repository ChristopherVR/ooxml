// Excel 97-2003 (.xls, BIFF8) import through the shared ole2 reader, which tsup inlines into the
// `xlsx/load` bundle. The result is a converted model: saving writes .xlsx, never .xls.
import {
	readXlsWorkbook,
	XlsReadError,
	type XlsWorkbook,
} from '@christophervr/ole2/legacy-excel-workbook';
import { readLegacyOfficeMetadata } from '@christophervr/ole2/legacy-office-metadata';
import type { DefinedName, Workbook, WorkbookProperties } from '../model.js';
import { internStyle } from '../styles.js';
import { createWorkbook, createWorksheet } from '../workbook.js';
import { mapSheet } from './legacy-xls-sheet.js';
import { defaultXfIndex, xfToCellStyle } from './legacy-xls-styles.js';

export type LegacyXlsErrorCode = 'encrypted' | 'unsupported-version' | 'corrupt';

/** The .xls file cannot be opened: encrypted, older than Excel 97 (BIFF5 and before) or corrupt. */
export class LegacyXlsError extends Error {
	readonly code: LegacyXlsErrorCode;
	constructor(message: string, code: LegacyXlsErrorCode = 'corrupt') {
		super(message);
		this.name = 'LegacyXlsError';
		this.code = code;
	}
}

const MESSAGES: Readonly<Record<LegacyXlsErrorCode, string>> = {
	encrypted:
		'This .xls workbook is password protected. Remove the password in Excel and open it again.',
	'unsupported-version':
		'This workbook uses an Excel format older than Excel 97 (BIFF5 or earlier), which is not supported.',
	corrupt: 'This file is not a readable Excel 97-2003 (.xls) workbook.',
};

const toBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
	input instanceof Uint8Array ? input : new Uint8Array(input);

function readSource(bytes: Uint8Array): XlsWorkbook {
	try {
		return readXlsWorkbook(bytes);
	} catch (error) {
		if (error instanceof XlsReadError) {
			const wrapped = new LegacyXlsError(MESSAGES[error.code], error.code);
			wrapped.cause = error;
			throw wrapped;
		}
		const wrapped = new LegacyXlsError(MESSAGES.corrupt);
		wrapped.cause = error;
		throw wrapped;
	}
}

function properties(bytes: Uint8Array): WorkbookProperties {
	let summary: ReturnType<typeof readLegacyOfficeMetadata>;
	try {
		summary = readLegacyOfficeMetadata(bytes);
	} catch {
		return {};
	}
	const out: WorkbookProperties = {};
	if (summary?.title) out.title = summary.title;
	if (summary?.subject) out.subject = summary.subject;
	if (summary?.author) out.creator = summary.author;
	if (summary?.keywords) out.keywords = summary.keywords;
	if (summary?.comments) out.description = summary.comments;
	if (summary?.lastAuthor) out.lastModifiedBy = summary.lastAuthor;
	if (summary?.application) out.application = summary.application;
	return out;
}

const plural = (count: number, word: string): string => `${count} ${word}${count === 1 ? '' : 's'}`;

/** Loads an Excel 97-2003 workbook into the model (`workbook.format` is `'xls'`). */
export async function loadLegacyXls(input: Uint8Array | ArrayBuffer): Promise<Workbook> {
	const bytes = toBytes(input);
	const xls = readSource(bytes);
	const workbook = createWorkbook();
	workbook.format = 'xls';
	workbook.date1904 = xls.date1904;
	workbook.properties = properties(bytes);
	if (xls.structureLocked) workbook.structureLocked = true;
	const warnings = workbook.warnings;
	warnings.push(
		'Opened an Excel 97-2003 (.xls) workbook as a converted copy; saving writes .xlsx. Rich text runs within cells, named cell styles, print settings and the workbook theme are not imported.',
	);

	const defaultXf = defaultXfIndex(xls);
	const defaultSource = xls.xfs[defaultXf];
	if (defaultSource) {
		const style = xfToCellStyle(xls, defaultSource);
		workbook.styles = [style];
		workbook.namedStyles = [{ name: 'Normal', style, builtinId: 0 }];
	}
	const styleIds = new Map<number, number>();
	const resolve = (xf: number): number => {
		const known = styleIds.get(xf);
		if (known !== undefined) return known;
		const source = xls.xfs[xf];
		const id = xf === defaultXf || !source ? 0 : internStyle(workbook, xfToCellStyle(xls, source));
		styleIds.set(xf, id);
		return id;
	};

	const sheetIndex = new Map<number, number>();
	workbook.sheets = [];
	let undecoded = 0;
	let external = false;
	xls.sheets.forEach((source, index) => {
		if (source.kind !== 'worksheet') {
			const kind = source.kind === 'chart' ? 'chart sheet' : `${source.kind} sheet`;
			warnings.push(
				`Skipped the ${kind} "${source.name}": only worksheets are imported from .xls.`,
			);
			return;
		}
		const mapped = mapSheet(source, workbook.sheets.length + 1, resolve);
		sheetIndex.set(index, workbook.sheets.length);
		workbook.sheets.push(mapped.sheet);
		undecoded += mapped.undecodedFormulas;
		external ||= mapped.externalReferences;
		if (source.unsupported.length)
			warnings.push(
				`Sheet "${source.name}" has ${source.unsupported.join(', ')} that the .xls import does not load; they are not shown and are dropped on save.`,
			);
	});
	if (workbook.sheets.length === 0) {
		workbook.sheets.push(createWorksheet('Sheet1', 1));
		warnings.push('The .xls file has no worksheets; an empty sheet was added.');
	}
	workbook.activeSheet = sheetIndex.get(xls.activeSheet) ?? 0;

	workbook.definedNames = [];
	let skippedNames = 0;
	for (const name of xls.names) {
		if (name.isFunction && name.name.startsWith('_xlfn.')) continue;
		const local = name.localSheet === undefined ? undefined : sheetIndex.get(name.localSheet);
		if (name.formula === undefined || (name.localSheet !== undefined && local === undefined)) {
			skippedNames++;
			continue;
		}
		const defined: DefinedName = {
			name: name.builtin ? `_xlnm.${name.name}` : name.name,
			formula: name.formula,
		};
		if (local !== undefined) defined.localSheet = local;
		if (name.hidden) defined.hidden = true;
		workbook.definedNames.push(defined);
	}

	if (undecoded)
		warnings.push(
			`${plural(undecoded, 'formula')} could not be decoded from the .xls file and ${undecoded === 1 ? 'is' : 'are'} kept as cached values.`,
		);
	if (skippedNames)
		warnings.push(
			`${plural(skippedNames, 'defined name')} could not be decoded and ${skippedNames === 1 ? 'was' : 'were'} skipped.`,
		);
	if (external)
		warnings.push(
			'Formulas that reference other workbooks keep their cached values; external links are not imported.',
		);
	for (const item of xls.unsupported)
		warnings.push(
			`The .xls file contains ${item}, which ${item.endsWith('s') ? 'are' : 'is'} not loaded and will not be saved.`,
		);
	return workbook;
}
