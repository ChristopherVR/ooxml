// Reference rewriting for parts the writer copies from the source package without modelling them
// (pivot cache definitions, chart parts behind chartsheets): sheet renames, deleted sheets and row
// or column inserts and deletes made in the editor must reach them, as Excel does on save.
import type { Band } from '../edit/shift-sheet.js';
import type { Workbook, Worksheet } from '../model.js';
import type { SourceIndex } from '../read/package.js';
import { parseWorkbookPart } from '../read/workbook-part.js';
import type { PackageWriter } from './package-writer.js';
import { shiftFormulaInBand } from '../edit/band-formulas.js';
import {
	deleteSheetInFormula,
	renameSheetInFormula,
	shiftFormula,
	type ShiftSpec,
} from '../formula/index.js';

/** A row or column insert (`count` > 0) or delete (`count` < 0) on one sheet. */
export interface CarriedShift {
	axis: 'row' | 'col';
	/** Zero-based first row or column affected. */
	at: number;
	count: number;
	/** Insert/delete cells: only the band of columns (row shift) or rows (column shift) moves. */
	band?: Band;
}

/**
 * The structural edits since load that carried parts must follow, in the sheet names the source
 * package uses. Apply order: shifts (in order), then deleted sheets, then renames (simultaneous).
 */
export interface CarriedRefEdits {
	/** Shifts in edit order, keyed by the source name of the sheet they happened on. */
	shifts: { sheet: string; shift: CarriedShift }[];
	/** Source names of sheets that no longer exist. */
	deleted: string[];
	/** Source name to current name, for sheets whose name changed. */
	renamed: [from: string, to: string][];
}

interface SheetIdentity {
	name: string;
	sheetId: number;
}

/**
 * Derives the edits by comparing the sheets in the source `workbook.xml` with the current ones
 * (matched by the stable `sheetId`). `shiftsOf` returns the shifts recorded for a current sheet.
 */
export function carriedRefEdits<T extends SheetIdentity>(
	sourceSheets: readonly SheetIdentity[],
	currentSheets: readonly T[],
	shiftsOf: (sheet: T) => readonly CarriedShift[] = () => [],
): CarriedRefEdits {
	const edits: CarriedRefEdits = { shifts: [], deleted: [], renamed: [] };
	const byId = new Map(currentSheets.map((sheet) => [sheet.sheetId, sheet]));
	for (const source of sourceSheets) {
		const current = byId.get(source.sheetId);
		if (!current) {
			edits.deleted.push(source.name);
			continue;
		}
		for (const shift of shiftsOf(current)) edits.shifts.push({ sheet: source.name, shift });
		if (current.name !== source.name) edits.renamed.push([source.name, current.name]);
	}
	return edits;
}

export const hasCarriedRefEdits = (edits: CarriedRefEdits): boolean =>
	edits.shifts.length > 0 || edits.deleted.length > 0 || edits.renamed.length > 0;

function shiftOne(
	formula: string,
	formulaSheet: string,
	sheet: string,
	shift: CarriedShift,
): string {
	if (shift.band) return shiftFormulaInBand(formula, formulaSheet, sheet, shift, shift.band);
	const spec: ShiftSpec = { sheet, axis: shift.axis, at: shift.at, count: shift.count };
	return shiftFormula(formula, formulaSheet, spec);
}

/** Renames sheets simultaneously (a swap of two names must not collapse them into one). */
function renameAll(
	formula: string,
	renamed: CarriedRefEdits['renamed'],
	taken: Set<string>,
): string {
	if (!renamed.length) return formula;
	let out = formula;
	const temps: [string, string][] = [];
	renamed.forEach(([from, to], i) => {
		let temp = `zzCarriedSheet${i}`;
		while (taken.has(temp.toLowerCase())) temp += '_';
		temps.push([temp, to]);
		out = renameSheetInFormula(out, from, temp);
	});
	for (const [temp, to] of temps) out = renameSheetInFormula(out, temp, to);
	return out;
}

const takenNames = (edits: CarriedRefEdits): Set<string> =>
	new Set(
		[...edits.deleted, ...edits.renamed.flat(), ...edits.shifts.map((s) => s.sheet)].map((n) =>
			n.toLowerCase(),
		),
	);

/**
 * Applies the edits to one formula as it appears in a carried part. `formulaSheet` resolves
 * unqualified references (`''` when every reference is sheet-qualified, as in chart parts).
 */
export function patchCarriedFormula(
	formula: string,
	formulaSheet: string,
	edits: CarriedRefEdits,
): string {
	let out = formula;
	for (const { sheet, shift } of edits.shifts) out = shiftOne(out, formulaSheet, sheet, shift);
	for (const sheet of edits.deleted) out = deleteSheetInFormula(out, sheet);
	return renameAll(out, edits.renamed, takenNames(edits));
}

const unescapeXml = (text: string): string =>
	text.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (whole, code: string) => {
		if (code === 'amp') return '&';
		if (code === 'lt') return '<';
		if (code === 'gt') return '>';
		if (code === 'quot') return '"';
		if (code === 'apos') return "'";
		const n = code[1] === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
		return Number.isFinite(n) ? String.fromCodePoint(n) : whole;
	});
const escapeXmlText = (text: string): string =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeXmlAttr = (text: string): string => escapeXmlText(text).replace(/"/g, '&quot;');

/** Every `<prefix:f ...>formula</prefix:f>` in a chart or chartEx part (`c:f`, `c15:f`, `cx:f`). */
const CHART_FORMULA = /<([A-Za-z][\w.-]*:)?f(\s[^>]*)?>([^<]*)<\/\1?f>/g;

function patchChartFormulas(xml: string, edits: CarriedRefEdits): string {
	return xml.replace(CHART_FORMULA, (whole, prefix = '', attrs = '', text: string) => {
		const formula = unescapeXml(text);
		const next = patchCarriedFormula(formula, '', edits);
		if (next === formula) return whole;
		return `<${prefix}f${attrs}>${escapeXmlText(next)}</${prefix}f>`;
	});
}

const WORKSHEET_SOURCE = /<((?:[A-Za-z][\w.-]*:)?worksheetSource)(\s[^>]*?)(\/?)>/;
const attrPattern = (name: string): RegExp => new RegExp(`(\\s${name}\\s*=\\s*)(["'])(.*?)\\2`);

/**
 * Rewrites `cacheSource/worksheetSource` (`sheet` + `ref`). A `name` source (defined name or
 * table), an `r:id` (external workbook) source or a source on a deleted sheet is left as it is,
 * as Excel does; a ref whose rows or columns were all deleted is kept (the cache keeps its data).
 */
function patchWorksheetSource(xml: string, edits: CarriedRefEdits): string {
	const match = WORKSHEET_SOURCE.exec(xml);
	if (!match) return xml;
	const [whole, tag, attrs = '', close] = match;
	if (/\sr:id\s*=/.test(attrs)) return xml;
	const sheetMatch = attrPattern('sheet').exec(attrs);
	const refMatch = attrPattern('ref').exec(attrs);
	if (!sheetMatch || !refMatch) return xml;
	const sheet = unescapeXml(sheetMatch[3] ?? '');
	const ref = unescapeXml(refMatch[3] ?? '');
	if (edits.deleted.some((name) => name.toLowerCase() === sheet.toLowerCase())) return xml;
	let nextRef = ref;
	for (const shifted of edits.shifts)
		nextRef = shiftOne(nextRef, sheet, shifted.sheet, shifted.shift);
	if (nextRef.includes('#REF!')) nextRef = ref;
	const rename = edits.renamed.find(([from]) => from.toLowerCase() === sheet.toLowerCase());
	const nextSheet = rename ? rename[1] : sheet;
	if (nextRef === ref && nextSheet === sheet) return xml;
	let nextAttrs = attrs;
	if (nextRef !== ref)
		nextAttrs = nextAttrs.replace(attrPattern('ref'), (_m, lead: string, q: string) => {
			return `${lead}${q}${escapeXmlAttr(nextRef)}${q}`;
		});
	if (nextSheet !== sheet)
		nextAttrs = nextAttrs.replace(attrPattern('sheet'), (_m, lead: string, q: string) => {
			return `${lead}${q}${escapeXmlAttr(nextSheet)}${q}`;
		});
	return xml.replace(whole, `<${tag}${nextAttrs}${close}>`);
}

/** Root element local name of a part, from its first start tag. */
function rootName(xml: string): string | undefined {
	const match = /<(?![?!])(?:[A-Za-z][\w.-]*:)?([A-Za-z][\w.-]*)/.exec(xml);
	return match?.[1];
}

/**
 * The carried part rewritten for the structural edits, or `undefined` when it needs no change
 * (or is not a part this function knows: pivot cache definitions, `chartSpace` chart and chartEx
 * parts). Everything except the touched references keeps its source text.
 */
export function patchCarriedPart(
	partName: string,
	xml: string,
	edits: CarriedRefEdits,
): string | undefined {
	if (!hasCarriedRefEdits(edits) || !/\.xml$/i.test(partName)) return undefined;
	const root = rootName(xml);
	let next: string;
	if (root === 'pivotCacheDefinition') next = patchWorksheetSource(xml, edits);
	else if (root === 'chartSpace') next = patchChartFormulas(xml, edits);
	else return undefined;
	return next === xml ? undefined : next;
}

/**
 * The edits between the source package and the current workbook: renames and deletions come from
 * the sheets of the source `workbook.xml` (matched by `sheetId`), shifts from `shiftsOf`.
 */
export function workbookRefEdits(
	workbook: Workbook,
	source: SourceIndex | undefined,
	shiftsOf: (sheet: Worksheet) => readonly CarriedShift[] = () => [],
): CarriedRefEdits {
	const bookPart = source?.workbookPart();
	const xml = bookPart ? source?.text(bookPart) : undefined;
	if (!xml) return { shifts: [], deleted: [], renamed: [] };
	let sheets: SheetIdentity[];
	try {
		sheets = parseWorkbookPart(xml).sheets;
	} catch {
		return { shifts: [], deleted: [], renamed: [] };
	}
	return carriedRefEdits(sheets, workbook.sheets, shiftsOf);
}

/**
 * Rewrites every part the writer has so far copied byte for byte (and not regenerated) for the
 * edits. Call it after all parts are written and before `writer.build()`.
 */
export function patchCarriedParts(writer: PackageWriter, edits: CarriedRefEdits): void {
	const source = writer.source;
	if (!source || !hasCarriedRefEdits(edits)) return;
	for (const [output, sourceName] of writer.carriedParts()) {
		const xml = source.text(sourceName);
		if (xml === undefined) continue;
		const patched = patchCarriedPart(output, xml, edits);
		if (patched !== undefined) writer.replace(output, patched);
	}
}
