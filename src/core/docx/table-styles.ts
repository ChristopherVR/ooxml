// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	TableConditionalRegion,
	TableStyleCatalog,
	TableStyleConditionalFormatting,
	TableStyleDefinition,
} from './table-model';
import { children, first, getW, named, parseXml, type XmlDocument, type XmlElement } from './xml';
import { parseRunProperties } from './run-properties';
import { parseOnOff } from './simple-types';
import { inheritanceCycleEntries } from './style-inheritance';
import { parseCellMargins } from './table-margins';
import { parseShadingFill, parseShadingThemeFill, parseTableBorders } from './table-borders';

const REGIONS: TableConditionalRegion[] = [
	'wholeTable',
	'firstRow',
	'lastRow',
	'firstCol',
	'lastCol',
	'band1Horz',
	'band2Horz',
	'band1Vert',
	'band2Vert',
];

function tableStyleElements(document: XmlDocument): XmlElement[] {
	return Array.from(document.getElementsByTagName('*')).filter(
		(element): element is XmlElement =>
			named(element, 'style') && getW(element, 'type') === 'table',
	);
}

function conditionalFormatting(block: XmlElement): TableStyleConditionalFormatting {
	const result: TableStyleConditionalFormatting = {};
	const tcPr = first(block, 'tcPr');
	const borders = parseTableBorders(first(tcPr, 'tcBorders'));
	if (borders) result.borders = borders;
	const shd = first(tcPr, 'shd');
	const fill = parseShadingFill(shd);
	if (fill) result.shadingFill = fill;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const run = parseRunProperties(first(block, 'rPr'));
	if (Object.keys(run).length) result.run = run;
	return result;
}

/** Parses `styles.xml` table styles, including `w:tblStylePr` conditional formatting regions. */
export function parseTableStyleCatalog(xml: string): TableStyleCatalog {
	const document = parseXml(xml);
	const styles = Object.create(null) as Record<string, TableStyleDefinition>;
	for (const element of tableStyleElements(document)) {
		const id = getW(element, 'styleId');
		if (!id) continue;
		const basedOn = getW(first(element, 'basedOn'), 'val');
		const isDefault = parseOnOff(getW(element, 'default'));
		const name = getW(first(element, 'name'), 'val');
		const tblPr = first(element, 'tblPr');
		const borders = parseTableBorders(first(tblPr, 'tblBorders'));
		const shadingFill = parseShadingFill(first(tblPr, 'shd'));
		const cellMargins = parseCellMargins(first(tblPr, 'tblCellMar'));
		const conditional: Partial<Record<TableConditionalRegion, TableStyleConditionalFormatting>> =
			{};
		for (const block of children(element, 'tblStylePr')) {
			const type = getW(block, 'type');
			if (type && (REGIONS as string[]).includes(type))
				conditional[type as TableConditionalRegion] = conditionalFormatting(block);
		}
		styles[id] = {
			id,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
			...(borders ? { borders } : {}),
			...(shadingFill ? { shadingFill } : {}),
			...(cellMargins ? { cellMargins } : {}),
			conditional,
		};
	}
	const warnings = inheritanceCycleEntries(styles).map(
		(id) => `Table style inheritance cycle detected at "${id}"; cyclic inheritance is ignored.`,
	);
	return { styles, warnings };
}
