import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import { elements, parseXml, relAttr, type XmlElement } from '../../xml/index.js';
import { parseRange } from '../address.js';
import type { DifferentialStyle, Hyperlink, Worksheet } from '../model.js';
import { createWorksheet } from '../workbook.js';
import { mergeComments, parseLegacyComments, parseThreadedComments } from './comments.js';
import { readConditionalFormats, readDataValidations } from './conditional.js';
import { parseDrawing } from './drawing.js';
import type { SourceIndex } from './package.js';
import { readSheetData, type CellContext } from './sheet-data.js';
import {
	readAutoFilter,
	readColumns,
	readFitToPage,
	readPageSetup,
	readProtection,
	readSheetFormat,
	readSheetView,
	readTabColor,
} from './sheet-props.js';
import { readPrintOptions } from './print-options.js';
import { parseTable } from './tables.js';
import { att, outerXml, selfContainedXml, xChildren } from './xml-util.js';

export interface SheetContext extends CellContext {
	source: SourceIndex;
	dxfs: readonly DifferentialStyle[];
	persons: ReadonlyMap<string, string>;
}

/**
 * Elements the model represents but whose source XML is kept under `source:<name>` in
 * {@link Worksheet.preserved}, so a save can write the original back when the modelled part of
 * it did not change (keeping attributes the model does not carry).
 */
export const SNAPSHOT_ELEMENTS = [
	'sheetPr',
	'sheetViews',
	'sheetFormatPr',
	'sheetProtection',
	'autoFilter',
	'pageMargins',
	'pageSetup',
	'headerFooter',
] as const;
export const snapshotKey = (local: string) => `source:${local}`;

/** Elements written from the model; never kept verbatim. */
const MODELLED = new Set([
	'dimension',
	'cols',
	'sheetData',
	'mergeCells',
	'conditionalFormatting',
	'dataValidations',
	'hyperlinks',
	'printOptions',
	'drawing',
	'legacyDrawing',
	'tableParts',
]);

function keep(sheet: Worksheet, key: string, xml: string): void {
	const list = sheet.preserved.get(key);
	if (list) list.push(xml);
	else sheet.preserved.set(key, [xml]);
}

/** The local name an `mc:AlternateContent` stands in for (its first choice's content). */
function alternateContentName(node: XmlElement): string | undefined {
	const choice = elements(node)[0];
	return choice ? (elements(choice)[0]?.localName ?? undefined) : undefined;
}

function readHyperlinks(
	ctx: SheetContext,
	partName: string,
	node: XmlElement | undefined,
): Hyperlink[] {
	if (!node) return [];
	const rels = ctx.source.rels(partName);
	const out: Hyperlink[] = [];
	for (const link of xChildren(node, 'hyperlink')) {
		const range = parseRange(att(link, 'ref') ?? '');
		if (!range) continue;
		const hyperlink: Hyperlink = { range };
		const rel = rels.get(relAttr(link, 'id') ?? '');
		if (rel) hyperlink.target = rel.target;
		const location = att(link, 'location');
		if (location) hyperlink.location = location;
		const tooltip = att(link, 'tooltip');
		if (tooltip) hyperlink.tooltip = tooltip;
		const display = att(link, 'display');
		if (display) hyperlink.display = display;
		out.push(hyperlink);
	}
	return out;
}

function readTables(
	ctx: SheetContext,
	partName: string,
	node: XmlElement | undefined,
	sheet: Worksheet,
): void {
	if (!node) return;
	const rels = ctx.source.rels(partName);
	for (const part of xChildren(node, 'tablePart')) {
		const rel = rels.get(relAttr(part, 'id') ?? '');
		const target = rel ? ctx.source.target(partName, rel) : undefined;
		const xml = target ? ctx.source.text(target) : undefined;
		const table = target && xml ? parseTable(xml, target) : undefined;
		if (table) sheet.tables.push(table);
		else ctx.warn('A table part could not be read and was dropped.');
	}
}

function readComments(ctx: SheetContext, partName: string, sheet: Worksheet): void {
	const legacyPart = ctx.source.targetOfType(partName, RELATIONSHIP_TYPES.comments);
	const threadedPart = ctx.source.targetOfType(partName, RELATIONSHIP_TYPES.threadedComment);
	const legacy = parseLegacyComments(legacyPart ? ctx.source.text(legacyPart) : undefined);
	const threaded = parseThreadedComments(
		threadedPart ? ctx.source.text(threadedPart) : undefined,
		ctx.persons,
	);
	sheet.comments = mergeComments(legacy, threaded);
}

/** Reads one worksheet part. */
export function parseWorksheet(
	ctx: SheetContext,
	partName: string,
	name: string,
	sheetId: number,
): Worksheet {
	const sheet = createWorksheet(name, sheetId);
	sheet.partName = partName;
	const xml = ctx.source.text(partName);
	if (!xml) {
		ctx.warn(`Sheet "${name}" has no part ${partName}; it was loaded empty.`);
		return sheet;
	}
	const root = parseXml(xml, { label: 'XLSX worksheet' }).documentElement;
	const byName = new Map<string, XmlElement>();
	const cfBlocks: XmlElement[] = [];
	for (const node of elements(root)) {
		const local = node.localName ?? '';
		if (local === 'conditionalFormatting') cfBlocks.push(node);
		else if (!byName.has(local)) byName.set(local, node);
		if ((SNAPSHOT_ELEMENTS as readonly string[]).includes(local))
			keep(sheet, snapshotKey(local), outerXml(node));
		else if (local === 'AlternateContent')
			keep(sheet, alternateContentName(node) ?? local, selfContainedXml(node));
		else if (!MODELLED.has(local)) keep(sheet, local, selfContainedXml(node));
	}
	const sheetPr = byName.get('sheetPr');
	const tabColor = readTabColor(sheetPr, ctx.palette);
	if (tabColor) sheet.tabColor = tabColor;
	sheet.view = readSheetView(byName.get('sheetViews'));
	const format = readSheetFormat(byName.get('sheetFormatPr'));
	sheet.defaultRowHeight = format.defaultRowHeight;
	if (format.defaultColWidth !== undefined) sheet.defaultColWidth = format.defaultColWidth;
	if (format.format) sheet.format = format.format;
	sheet.columns = readColumns(byName.get('cols'), ctx.xfMap);
	const data = byName.get('sheetData');
	if (data) readSheetData(ctx, data, sheet);
	const protection = readProtection(byName.get('sheetProtection'));
	if (protection) sheet.protection = protection;
	const autoFilter = readAutoFilter(byName.get('autoFilter'));
	if (autoFilter) sheet.autoFilter = autoFilter;
	sheet.merges = xChildren(byName.get('mergeCells') ?? root, 'mergeCell')
		.map((node) => parseRange(att(node, 'ref') ?? ''))
		.filter((range) => range !== undefined);
	sheet.conditionalFormats = readConditionalFormats(cfBlocks, ctx.dxfs, ctx.palette, (m) =>
		ctx.warn(m),
	);
	sheet.dataValidations = readDataValidations(byName.get('dataValidations'));
	sheet.hyperlinks = readHyperlinks(ctx, partName, byName.get('hyperlinks'));
	const page = readPageSetup(
		byName.get('pageMargins'),
		byName.get('pageSetup'),
		byName.get('headerFooter'),
		readFitToPage(sheetPr),
	);
	if (page) sheet.pageSetup = page;
	const printOptions = readPrintOptions(byName.get('printOptions'));
	if (printOptions) sheet.printOptions = printOptions;
	const drawing = byName.get('drawing');
	if (drawing) {
		const rel = ctx.source.rels(partName).get(relAttr(drawing, 'id') ?? '');
		const target = rel ? ctx.source.target(partName, rel) : undefined;
		if (target) sheet.drawings = parseDrawing(ctx.source, target, (m) => ctx.warn(m));
	}
	readTables(ctx, partName, byName.get('tableParts'), sheet);
	readComments(ctx, partName, sheet);
	return sheet;
}
