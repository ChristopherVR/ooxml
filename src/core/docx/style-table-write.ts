import type { Table } from './model';
import type { TableConditionalRegion, TableStyleDefinition } from './table-model';
import { buildCellProperties } from './table-cell-write';
import { orderChildren } from './element-order';
import {
	patchStyleParagraph,
	patchStyleRun,
	setStyleAttribute,
	styleChild,
} from './style-formatting-package';
import { patchTableProperties } from './write-table-level-properties';
import {
	CELL_ORDER,
	patchCellBorders,
	patchTableMargins,
	TABLE_ORDER,
} from './write-table-properties';
import { children, first, getW, type XmlElement } from './xml';

function patchShading(
	owner: XmlElement,
	property: 'tblPr' | 'tcPr',
	next: Partial<TableStyleDefinition>,
	prior: Partial<TableStyleDefinition>,
) {
	if (
		next.shadingFill === prior.shadingFill &&
		JSON.stringify(next.shadingThemeFill) === JSON.stringify(prior.shadingThemeFill) &&
		JSON.stringify(next.percentageShading) === JSON.stringify(prior.percentageShading)
	)
		return;
	const props = styleChild(owner, property);
	for (const old of children(props, 'shd')) props.removeChild(old);
	const shade = first(buildCellProperties(owner.ownerDocument, { paragraphs: [], ...next }), 'shd');
	if (shade) props.appendChild(shade);
	orderChildren(props, property === 'tblPr' ? TABLE_ORDER : CELL_ORDER);
}

export function patchTableStyle(
	style: XmlElement,
	next: TableStyleDefinition,
	prior?: TableStyleDefinition,
): void {
	const doc = style.ownerDocument;
	const nextTable: Table = { type: 'table', rows: [], ...next };
	const baseTable: Table | undefined = prior && { type: 'table', rows: [], ...prior };
	patchTableProperties(doc, nextTable, style, baseTable);
	patchTableMargins(doc, nextTable, style, baseTable);
	patchShading(style, 'tblPr', next, prior ?? {});
	patchStyleParagraph(style, next.paragraphFormatting, prior?.paragraphFormatting);
	patchStyleRun(style, next.runFormatting, prior?.runFormatting);
	const regions = new Set<TableConditionalRegion>([
		...(Object.keys(next.conditional) as TableConditionalRegion[]),
		...(Object.keys(prior?.conditional ?? {}) as TableConditionalRegion[]),
	]);
	for (const region of regions) {
		const after = next.conditional[region];
		const before = prior?.conditional[region];
		if (JSON.stringify(after) === JSON.stringify(before)) continue;
		let block = children(style, 'tblStylePr').find((node) => getW(node, 'type') === region);
		if (!after) {
			if (block) style.removeChild(block);
			continue;
		}
		if (!block) {
			block = doc.createElementNS(style.namespaceURI, 'w:tblStylePr');
			setStyleAttribute(block, 'type', region);
			style.appendChild(block);
		}
		patchStyleRun(block, after.run, before?.run);
		patchCellBorders(doc, block, after.borders, before?.borders);
		patchShading(block, 'tcPr', after, before ?? {});
		orderChildren(block, ['pPr', 'rPr', 'tblPr', 'trPr', 'tcPr']);
	}
}
