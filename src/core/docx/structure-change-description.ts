/**
 * Describes tracked table, row, cell and section property changes (`w:tblPrChange`,
 * `w:trPrChange`, `w:tcPrChange`, `w:sectPrChange`) the way Word's markup balloons word them.
 */
import { parseXml } from '../xml/index';
import type { SectionProperties, Table, TableCell } from './model';
import type { TableRowProperties } from './table-model';
import { first, getW, type XmlElement } from './xml';
import { onOffElement, parseSignedTwips, parseTwips, parseUnsignedInteger } from './simple-types';
import { parseShadingFill, parseTableBorders } from './table-borders';
import { parseCellMargins } from './table-margins';
import { wordColorName } from './format-change-description';

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const number = (value: number) => String(Number(value.toFixed(2)));
const inches = (twips: number) => `${number(twips / 1440)}"`;
const title = (value: string) =>
	value.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/^./, (c) => c.toUpperCase());

function parsePrior(xml: string | undefined): XmlElement | undefined {
	if (!xml) return undefined;
	try {
		return parseXml(xml).documentElement;
	} catch {
		return undefined;
	}
}

/** The items of a table's property change (`w:tblPrChange`), or undefined when it has none. */
export function describeTableFormatChange(table: Table): string[] | undefined {
	if (table.formatRevision?.kind !== 'tableChange') return undefined;
	const prior = parsePrior(table.formatRevision.previousPropertiesXml);
	const items: string[] = [];
	if (prior) {
		const priorW = first(prior, 'tblW');
		const priorWidth =
			(getW(priorW, 'type') ?? 'dxa') === 'dxa' ? parseTwips(getW(priorW, 'w')) : undefined;
		if (priorWidth !== table.widthTwips)
			items.push(
				table.widthTwips !== undefined
					? `Table width: ${inches(table.widthTwips)}`
					: 'Table width: Auto',
			);
		const priorJc = getW(first(prior, 'jc'), 'val');
		if (priorJc !== table.justification)
			items.push(`Table alignment: ${title(table.justification ?? 'left')}`);
		const priorInd = parseSignedTwips(getW(first(prior, 'tblInd'), 'w'));
		if (priorInd !== table.indentTwips)
			items.push(`Table indent: ${inches(table.indentTwips ?? 0)}`);
		const priorLayout = getW(first(prior, 'tblLayout'), 'type') === 'fixed' ? 'fixed' : 'autofit';
		if (priorLayout !== (table.layout ?? 'autofit'))
			items.push(`Table layout: ${table.layout === 'fixed' ? 'Fixed' : 'AutoFit'}`);
		const priorStyle = getW(first(prior, 'tblStyle'), 'val');
		if (priorStyle !== table.style) items.push(`Table style: ${table.style ?? 'None'}`);
		const priorBorders = parseTableBorders(first(prior, 'tblBorders'));
		if (!same(priorBorders, table.borders))
			items.push(table.borders ? 'Table borders' : 'No table borders');
		const priorMargins = parseCellMargins(first(prior, 'tblCellMar'));
		if (!same(priorMargins, table.cellMargins)) items.push('Table cell margins');
	}
	return items.length ? items : ['Table properties'];
}

/** The items of a row's property change (`w:trPrChange`), or undefined when it has none. */
export function describeRowFormatChange(row: TableRowProperties): string[] | undefined {
	if (row.formatRevision?.kind !== 'rowChange') return undefined;
	const prior = parsePrior(row.formatRevision.previousPropertiesXml);
	const items: string[] = [];
	if (prior) {
		const priorHeight = first(prior, 'trHeight');
		const priorTwips = parseTwips(getW(priorHeight, 'val'));
		const priorRule =
			priorTwips !== undefined
				? getW(priorHeight, 'hRule') === 'exact'
					? 'exact'
					: 'atLeast'
				: undefined;
		if (priorTwips !== row.heightTwips || priorRule !== row.heightRule)
			items.push(
				row.heightTwips !== undefined
					? `Row height: ${row.heightRule === 'exact' ? 'Exactly ' : 'At least '}${inches(row.heightTwips)}`
					: 'Row height: Auto',
			);
		const priorCantSplit = Boolean(onOffElement(first(prior, 'cantSplit')));
		if (priorCantSplit !== Boolean(row.cantSplit))
			items.push(row.cantSplit ? 'Prevent row from splitting' : 'Allow row to split');
		const priorHeader = Boolean(onOffElement(first(prior, 'tblHeader')));
		if (priorHeader !== Boolean(row.header))
			items.push(row.header ? 'Repeat as header row' : 'Not header row');
	}
	return items.length ? items : ['Row properties'];
}

/** The items of a cell's property change (`w:tcPrChange`), or undefined when it has none. */
export function describeCellFormatChange(cell: TableCell): string[] | undefined {
	if (cell.formatRevision?.kind !== 'cellChange') return undefined;
	const prior = parsePrior(cell.formatRevision.previousPropertiesXml);
	const items: string[] = [];
	if (prior) {
		const priorW = first(prior, 'tcW');
		const priorWidth =
			(getW(priorW, 'type') ?? 'dxa') === 'dxa' ? parseTwips(getW(priorW, 'w')) : undefined;
		if (priorWidth !== cell.widthTwips)
			items.push(
				cell.widthTwips !== undefined
					? `Cell width: ${inches(cell.widthTwips)}`
					: 'Cell width: Auto',
			);
		const priorVAlign = getW(first(prior, 'vAlign'), 'val');
		if (priorVAlign !== cell.verticalAlign)
			items.push(`Vertical alignment: ${title(cell.verticalAlign ?? 'top')}`);
		const priorShd = parseShadingFill(first(prior, 'shd'));
		if (priorShd !== cell.shadingFill)
			items.push(
				cell.shadingFill
					? `Shading: ${wordColorName(cell.shadingFill, cell.shadingThemeFill)}`
					: 'No shading',
			);
		const priorBorders = parseTableBorders(first(prior, 'tcBorders'));
		if (!same(priorBorders, cell.borders))
			items.push(cell.borders ? 'Cell borders' : 'No cell borders');
		const priorMargins = parseCellMargins(first(prior, 'tcMar'));
		if (!same(priorMargins, cell.margins)) items.push('Cell margins');
	}
	return items.length ? items : ['Cell properties'];
}

/** The items of a section's property change (`w:sectPrChange`), or undefined when it has none. */
export function describeSectionFormatChange(section: SectionProperties): string[] | undefined {
	if (section.formatRevision?.kind !== 'sectionChange') return undefined;
	const prior = parsePrior(section.formatRevision.previousPropertiesXml);
	const items: string[] = [];
	if (prior) {
		const size = first(prior, 'pgSz');
		const priorW = parseTwips(getW(size, 'w')) ?? 12240;
		const priorH = parseTwips(getW(size, 'h')) ?? 15840;
		if (priorW !== section.pageWidthTwips || priorH !== section.pageHeightTwips)
			items.push(
				`Page size: ${inches(section.pageWidthTwips)} x ${inches(section.pageHeightTwips)}`,
			);
		const priorOrient = getW(size, 'orient') === 'landscape' ? 'landscape' : 'portrait';
		if (priorOrient !== section.orientation)
			items.push(`Orientation: ${title(section.orientation)}`);
		const mar = first(prior, 'pgMar');
		const priorTop = parseSignedTwips(getW(mar, 'top')) ?? 1440;
		const priorBottom = parseSignedTwips(getW(mar, 'bottom')) ?? 1440;
		const priorLeft = parseTwips(getW(mar, 'left')) ?? 1440;
		const priorRight = parseTwips(getW(mar, 'right')) ?? 1440;
		const margins: string[] = [];
		if (priorTop !== section.marginTopTwips) margins.push(`Top: ${inches(section.marginTopTwips)}`);
		if (priorBottom !== section.marginBottomTwips)
			margins.push(`Bottom: ${inches(section.marginBottomTwips)}`);
		if (priorLeft !== section.marginLeftTwips)
			margins.push(`Left: ${inches(section.marginLeftTwips)}`);
		if (priorRight !== section.marginRightTwips)
			margins.push(`Right: ${inches(section.marginRightTwips)}`);
		if (margins.length) items.push(`Margins: ${margins.join(', ')}`);
		const cols = first(prior, 'cols');
		const priorCols = parseUnsignedInteger(getW(cols, 'num')) ?? 1;
		if (priorCols !== section.columns.count) items.push(`Columns: ${section.columns.count}`);
		const priorVAlign = getW(first(prior, 'vAlign'), 'val');
		if (priorVAlign !== section.verticalAlign)
			items.push(`Vertical alignment: ${title(section.verticalAlign ?? 'top')}`);
		if (Boolean(onOffElement(first(prior, 'titlePg'))) !== Boolean(section.titlePage))
			items.push(section.titlePage ? 'Different first page' : 'Same first page');
		if (Boolean(first(prior, 'lnNumType')) !== Boolean(section.lineNumbering))
			items.push(section.lineNumbering ? 'Line numbers' : 'No line numbers');
	}
	return items.length ? items : ['Section properties'];
}
