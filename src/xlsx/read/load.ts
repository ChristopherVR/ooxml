import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import { parseRange } from '../address.js';
import type { Workbook, Worksheet } from '../model.js';
import { createWorksheet } from '../workbook.js';
import { parsePersons } from './comments.js';
import { CONTENT_TYPES, SourceIndex, readZipParts } from './package.js';
import { parseDynamicArrayMetadata } from './metadata.js';
import { parseSharedStrings } from './shared-strings.js';
import { parseStyles } from './styles.js';
import { parseTheme } from './theme.js';
import { parseDocProps, parseWorkbookPart } from './workbook-part.js';
import { parseWorksheet, type SheetContext } from './worksheet.js';

const MACRO_TYPES = new Set<string>([CONTENT_TYPES.workbookMacro, CONTENT_TYPES.templateMacro]);
const PRINT_AREA = '_xlnm.Print_Area';

/** Moves single-range `_xlnm.Print_Area` names onto their sheet's page setup. */
function liftPrintAreas(workbook: Workbook): void {
	workbook.definedNames = workbook.definedNames.filter((name) => {
		if (name.name !== PRINT_AREA || name.localSheet === undefined) return true;
		const sheet = workbook.sheets[name.localSheet];
		const match = /^(?:'(?:[^']|'')+'|[^!,]+)!([$A-Z0-9:]+)$/i.exec(name.formula.trim());
		const range = match ? parseRange(match[1] ?? '') : undefined;
		if (!sheet || !range) return true;
		sheet.pageSetup = { ...sheet.pageSetup, printArea: range };
		return false;
	});
}

/**
 * Reads an `.xlsx`, `.xlsm` or `.xltx` package into the workbook model. Parts the model does not
 * represent stay in `workbook.source.parts` and are written back by `saveXlsx`.
 */
export async function loadXlsx(input: Uint8Array | ArrayBuffer): Promise<Workbook> {
	const parts = await readZipParts(input);
	const source = new SourceIndex(parts);
	const workbookPart = source.workbookPart();
	if (!workbookPart) throw new Error('XLSX package has no workbook part (xl/workbook.xml)');
	const workbookXml = source.text(workbookPart) ?? '';
	const book = parseWorkbookPart(workbookXml);
	const warnings: string[] = [];
	const warn = (message: string) => {
		if (!warnings.includes(message)) warnings.push(message);
	};
	const partOf = (type: string) => source.targetOfType(workbookPart, type);
	const stylesPart = partOf(RELATIONSHIP_TYPES.styles);
	const styles = parseStyles(stylesPart ? source.text(stylesPart) : undefined, warnings);
	const themePart = partOf(RELATIONSHIP_TYPES.theme);
	const stringsPart = partOf(RELATIONSHIP_TYPES.sharedStrings);
	const personsPart = partOf(RELATIONSHIP_TYPES.person);
	const metadataPart = partOf(RELATIONSHIP_TYPES.sheetMetadata);
	const ctx: SheetContext = {
		source,
		sharedStrings: parseSharedStrings(
			stringsPart ? source.text(stringsPart) : undefined,
			styles.palette,
		),
		xfMap: styles.xfMap,
		date1904: book.date1904,
		palette: styles.palette,
		dxfs: styles.dxfs,
		persons: parsePersons(personsPart ? source.text(personsPart) : undefined),
		dynamicCells: parseDynamicArrayMetadata(metadataPart ? source.text(metadataPart) : undefined),
		warn,
	};
	const rels = source.rels(workbookPart);
	const sheets: Worksheet[] = book.sheets.map((entry) => {
		const rel = rels.get(entry.relId);
		const target = rel ? source.target(workbookPart, rel) : undefined;
		let sheet: Worksheet;
		if (rel?.type === RELATIONSHIP_TYPES.worksheet && target) {
			sheet = parseWorksheet(ctx, target, entry.name, entry.sheetId);
		} else {
			sheet = createWorksheet(entry.name, entry.sheetId);
			if (target) sheet.partName = target;
			const kind = rel?.type.split('/').pop() ?? 'sheet';
			warn(`Sheet "${entry.name}" is a ${kind}; it is kept on save but cannot be shown or edited.`);
		}
		sheet.state = entry.state;
		return sheet;
	});
	if (source.has('xl/vbaProject.bin') || partOf(RELATIONSHIP_TYPES.vbaProject))
		warn('The workbook contains VBA macros; they are kept on save but never run.');
	if (partOf(RELATIONSHIP_TYPES.pivotCacheDefinition))
		warn('Pivot tables are kept on save but are not refreshed or shown as pivot tables.');
	if (partOf(RELATIONSHIP_TYPES.externalLink))
		warn('External workbook links are kept but not updated.');
	const contentType = source.contentType(workbookPart) ?? CONTENT_TYPES.workbook;
	const corePart = source.targetOfType('', RELATIONSHIP_TYPES.coreProperties);
	const appPart = source.targetOfType('', RELATIONSHIP_TYPES.extendedProperties);
	const workbook: Workbook = {
		sheets,
		styles: styles.styles,
		namedStyles: styles.namedStyles,
		definedNames: book.definedNames,
		theme: parseTheme(themePart ? source.text(themePart) : undefined),
		activeSheet: Math.min(Math.max(0, book.activeTab), Math.max(0, sheets.length - 1)),
		date1904: book.date1904,
		properties: parseDocProps(
			corePart ? source.text(corePart) : undefined,
			appPart ? source.text(appPart) : undefined,
		),
		source: { parts },
		format: MACRO_TYPES.has(contentType) ? 'xlsm' : 'xlsx',
		warnings,
	};
	if (book.fullCalcOnLoad) workbook.fullCalcOnLoad = true;
	if (book.structureLocked) workbook.structureLocked = true;
	if (book.workbookPasswordHash) workbook.workbookPasswordHash = book.workbookPasswordHash;
	if (book.workbookModernHash) workbook.workbookModernHash = book.workbookModernHash;
	if (book.calcMode) workbook.calcMode = book.calcMode;
	if (!sheets.length) throw new Error('XLSX workbook has no sheets');
	liftPrintAreas(workbook);
	return workbook;
}
