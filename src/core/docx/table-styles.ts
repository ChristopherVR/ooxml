import { parseFormatting } from './paragraph-styles';
import { readStyleMetadata } from './style-metadata';
import { parsePercentageShading } from './percentage-shading';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	TableConditionalRegion,
	TableStyleCatalog,
	TableStyleConditionalFormatting,
	TableStyleDefinition,
} from './table-model';
import { children, first, getW, parseXml, type XmlElement } from './xml';
import { parseRunProperties } from './run-properties';
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

function conditionalFormatting(block: XmlElement): TableStyleConditionalFormatting {
	const result: TableStyleConditionalFormatting = {};
	const tcPr = first(block, 'tcPr');
	const borders = parseTableBorders(first(tcPr, 'tcBorders'));
	if (borders) result.borders = borders;
	const shd = first(tcPr, 'shd');
	const fill = parseShadingFill(shd);
	const pattern = parsePercentageShading(shd);
	if (pattern) result.percentageShading = pattern;
	if (fill) result.shadingFill = fill;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const run = parseRunProperties(first(block, 'rPr'));
	if (Object.keys(run).length) result.run = run;
	return result;
}

/** Parses `styles.xml` table styles, including `w:tblStylePr` conditional formatting regions. */
export function parseTableStyleCatalog(xml: string | XmlElement): TableStyleCatalog {
	const root = typeof xml === 'string' ? parseXml(xml).documentElement : xml;
	const styles = Object.create(null) as Record<string, TableStyleDefinition>;
	for (const element of children(root, 'style')) {
		if (getW(element, 'type') !== 'table') continue;
		const id = getW(element, 'styleId');
		if (!id) continue;
		const tblPr = first(element, 'tblPr');
		const borders = parseTableBorders(first(tblPr, 'tblBorders'));
		const pPr = first(element, 'pPr');
		const rPr = first(element, 'rPr');
		const shd = first(tblPr, 'shd');
		const shadingThemeFill = parseShadingThemeFill(shd);
		const shadingFill = parseShadingFill(shd);
		const pattern = parsePercentageShading(shd);
		const cellMargins = parseCellMargins(first(tblPr, 'tblCellMar'));
		const conditional: Partial<Record<TableConditionalRegion, TableStyleConditionalFormatting>> =
			{};
		for (const block of children(element, 'tblStylePr')) {
			const type = getW(block, 'type');
			if (type && (REGIONS as string[]).includes(type))
				conditional[type as TableConditionalRegion] = conditionalFormatting(block);
		}
		styles[id] = {
			...readStyleMetadata(element, id),
			...(borders ? { borders } : {}),
			...(shadingFill ? { shadingFill } : {}),
			...(pattern ? { percentageShading: pattern } : {}),
			...(pPr && { paragraphFormatting: parseFormatting(pPr) }),
			...(rPr && { runFormatting: parseRunProperties(rPr) }),
			...(shadingThemeFill && { shadingThemeFill }),
			...(cellMargins ? { cellMargins } : {}),
			conditional,
		};
	}
	const warnings = inheritanceCycleEntries(styles).map(
		(id) => `Table style inheritance cycle detected at "${id}"; cyclic inheritance is ignored.`,
	);
	return { styles, warnings };
}
