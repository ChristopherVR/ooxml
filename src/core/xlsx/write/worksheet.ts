import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import { NS } from '../../xml/index.js';
import { formatRange } from '../address.js';
import type { Worksheet } from '../model.js';
import { CONTENT_TYPES } from '../read/package.js';
import { printOptionsXml } from '../read/print-options.js';
import { conditionalFormatsXml, dataValidationsXml } from './conditional.js';
import { RelationshipSet } from './package-writer.js';
import { pageXml } from './page-setup.js';
import { sheetDataXml } from './sheet-data.js';
import {
	sourceRelIds,
	writeComments,
	writeSheetDrawing,
	writeTables,
	type SaveContext,
} from './sheet-parts.js';
import {
	autoFilterXml,
	colsXml,
	dimensionXml,
	mergeCellsXml,
	protectionXml,
	sheetFormatXml,
	sheetPrXml,
	sheetViewsXml,
} from './sheet-props.js';
import { XML_HEADER, attrs, inlineFragment } from './xml-out.js';

export type { SaveContext } from './sheet-parts.js';

/** CT_Worksheet child order; kept elements are slotted in by local name. */
const ORDER = [
	'sheetPr',
	'dimension',
	'sheetViews',
	'sheetFormatPr',
	'cols',
	'sheetData',
	'sheetCalcPr',
	'sheetProtection',
	'protectedRanges',
	'scenarios',
	'autoFilter',
	'sortState',
	'dataConsolidate',
	'customSheetViews',
	'mergeCells',
	'phoneticPr',
	'conditionalFormatting',
	'dataValidations',
	'hyperlinks',
	'printOptions',
	'pageMargins',
	'pageSetup',
	'headerFooter',
	'rowBreaks',
	'colBreaks',
	'customProperties',
	'cellWatches',
	'ignoredErrors',
	'smartTags',
	'drawing',
	'legacyDrawing',
	'legacyDrawingHF',
	'drawingHF',
	'picture',
	'oleObjects',
	'controls',
	'webPublishItems',
	'tableParts',
	'extLst',
];

/** Relationship types the writer regenerates from the model instead of keeping. */
const REGENERATED = new Set<string>([
	RELATIONSHIP_TYPES.hyperlink,
	RELATIONSHIP_TYPES.comments,
	RELATIONSHIP_TYPES.threadedComment,
	RELATIONSHIP_TYPES.table,
	RELATIONSHIP_TYPES.drawing,
]);

const NS_XR2 = 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2';
const NS_XR3 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3';

/** Writes one worksheet part with its relationships and dependent parts. */
export function writeWorksheet(
	ctx: SaveContext,
	sheet: Worksheet,
	index: number,
	partName: string,
): void {
	const { writer, source } = ctx;
	const rels = new RelationshipSet();
	let sourceDrawing: string | undefined;
	if (source && sheet.partName && source.has(sheet.partName)) {
		const ids = sourceRelIds(source, sheet.partName);
		for (const [id, rel] of source.rels(sheet.partName)) {
			const target = source.target(sheet.partName, rel);
			if (rel.type === RELATIONSHIP_TYPES.drawing && id === ids.drawing) sourceDrawing = target;
			if (
				REGENERATED.has(rel.type) ||
				(rel.type === RELATIONSHIP_TYPES.vmlDrawing && id === ids.legacy)
			)
				continue;
			rels.keep(id, rel);
			if (target) writer.carry(target);
		}
	}
	const blocks = new Map<string, string>();
	const headerText = new Map<string, string>();
	const formulas = new Map<string, string>();
	const tableParts = writeTables(ctx, sheet, partName, rels, headerText, formulas);
	blocks.set('sheetPr', sheetPrXml(sheet));
	blocks.set('dimension', dimensionXml(sheet));
	blocks.set('sheetViews', sheetViewsXml(sheet, index === ctx.workbook.activeSheet));
	blocks.set('sheetFormatPr', sheetFormatXml(sheet));
	blocks.set('cols', colsXml(sheet, ctx.workbook.styles.length));
	blocks.set(
		'sheetData',
		sheetDataXml(
			{
				strings: ctx.strings,
				styleCount: ctx.workbook.styles.length,
				headerText,
				formulas,
				...(ctx.dynamicArrays ? { dynamicArrays: ctx.dynamicArrays } : {}),
				...(ctx.metadata ? { metadata: ctx.metadata } : {}),
			},
			sheet,
		),
	);
	blocks.set('sheetProtection', protectionXml(sheet));
	blocks.set('autoFilter', autoFilterXml(sheet));
	blocks.set('mergeCells', mergeCellsXml(sheet));
	blocks.set('conditionalFormatting', conditionalFormatsXml(sheet.conditionalFormats, ctx.styles));
	blocks.set('dataValidations', dataValidationsXml(sheet.dataValidations));
	const links = sheet.hyperlinks.map((link) => {
		const id = link.target ? rels.add(RELATIONSHIP_TYPES.hyperlink, link.target, true) : undefined;
		return `<hyperlink${attrs({ ref: formatRange(link.range), 'r:id': id, location: link.location, tooltip: link.tooltip, display: link.display })}/>`;
	});
	blocks.set('hyperlinks', links.length ? `<hyperlinks>${links.join('')}</hyperlinks>` : '');
	const page = pageXml(sheet);
	blocks.set('printOptions', printOptionsXml(sheet.printOptions));
	blocks.set('pageMargins', page.margins);
	// A copied sheet (or one whose source rels are gone) must not point at relationships it lacks.
	blocks.set(
		'pageSetup',
		page.setup.replace(/\sr:id="([^"]*)"/, (attr, id: string) =>
			rels.entries.has(id) ? attr : '',
		),
	);
	blocks.set('headerFooter', page.headerFooter);
	blocks.set('drawing', writeSheetDrawing(ctx, sheet, partName, rels, sourceDrawing));
	blocks.set('legacyDrawing', writeComments(ctx, sheet, index, partName, rels));
	blocks.set('tableParts', tableParts);
	for (const [key, list] of sheet.preserved) {
		if (key.startsWith('source:') || !ORDER.includes(key) || (blocks.get(key) ?? '') !== '')
			continue;
		const valid = list.filter((xml) =>
			[...xml.matchAll(/\br:(?:id|embed|link|pict)="([^"]*)"/g)].every((m) =>
				rels.entries.has(m[1] ?? ''),
			),
		);
		blocks.set(key, valid.map(inlineFragment).join(''));
	}
	const body = ORDER.map((name) => inlineFragment(blocks.get(name) ?? '')).join('');
	const root = `<worksheet xmlns="${NS.x}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="${NS.x14ac}" xmlns:xr="${NS.xr}" xmlns:xr2="${NS_XR2}" xmlns:xr3="${NS_XR3}">`;
	writer.add(partName, `${XML_HEADER}${root}${body}</worksheet>`, CONTENT_TYPES.worksheet);
	writer.rels(partName, rels);
}
