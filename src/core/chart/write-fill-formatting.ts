import { NS, elements, first, type XmlElement } from '../xml';
import { parseDrawingFill } from '../diagram/drawing-fill';
import { drawingFillXml, setDrawingFillXml } from '../diagram/write-fill';
import { chartFormattingNodes } from './formatting-nodes';
import type { ChartStyleDefinition } from './style-definition';

/** Patch direct chart-element fills using the shared DrawingML codec. */
export function writeChartFillFormatting(
	root: XmlElement,
	formatting: ChartStyleDefinition | undefined,
): boolean {
	if (!formatting) return false;
	let changed = false;
	for (const [part, node] of Object.entries(chartFormattingNodes(root))) {
		const fill = formatting.entries[part]?.fill;
		if (!node || !fill) continue;
		let properties = first(node, 'spPr', NS.c);
		if (JSON.stringify(parseDrawingFill(properties)) === JSON.stringify(fill)) continue;
		const xml = drawingFillXml(fill);
		if (xml === undefined) continue;
		if (!properties) {
			properties = root.ownerDocument.createElementNS(NS.c, 'c:spPr');
			const successors = [
				'txPr',
				'crossAx',
				'crosses',
				'crossesAt',
				'crossBetween',
				'auto',
				'lblAlgn',
				'lblOffset',
				'tickLblSkip',
				'tickMarkSkip',
				'noMultiLvlLbl',
				'externalData',
				'printSettings',
				'userShapes',
				'extLst',
			];
			node.insertBefore(
				properties,
				elements(node).find(
					(child) => child.namespaceURI === NS.c && successors.includes(child.localName),
				) ?? null,
			);
		}
		setDrawingFillXml(properties, xml);
		changed = true;
	}
	return changed;
}
