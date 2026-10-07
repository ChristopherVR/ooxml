import { NS, buildXml, first, type XmlElement } from '../xml/index';
import { readChartAppearance } from './read-appearance';
import type { ChartStyleDefinition, ChartStyleEntry } from './style-definition';
import { chartFormattingNodes } from './formatting-nodes';
import { readBuiltInChartStyle } from './built-in-text-style';

const c = (parent: XmlElement | undefined, name: string) => first(parent, name, NS.c);
const a = (parent: XmlElement | undefined, name: string) => first(parent, name, NS.a);

function entry(node: XmlElement): ChartStyleEntry {
	const body = c(node, 'txPr');
	const paragraph = a(body, 'p');
	const defaults = a(a(paragraph, 'pPr'), 'defRPr');
	const rich = c(c(node, 'tx'), 'rich');
	const richParagraph = a(rich, 'p');
	// Whole-element appearance uses the first rich run. Mixed-run styling remains in source XML.
	const out: ChartStyleEntry = {
		sourceXml: buildXml(node),
		...readChartAppearance(defaults, c(node, 'spPr')),
		...readChartAppearance(a(a(richParagraph, 'pPr'), 'defRPr'), undefined),
		...readChartAppearance(a(a(richParagraph, 'r'), 'rPr'), undefined),
	};
	if (['catAx', 'dateAx', 'valAx'].includes(node.localName)) {
		const deleted = c(node, 'delete')?.getAttribute('val');
		if (deleted !== undefined) out.axisVisible = deleted !== '1' && deleted !== 'true';
		out.labelsVisible =
			out.axisVisible !== false && c(node, 'tickLblPos')?.getAttribute('val') !== 'none';
	}
	return out;
}

/** Direct element formatting overrides the external style without flattening theme colors. */
export function readChartFormatting(root: XmlElement): ChartStyleDefinition | undefined {
	const nodes = chartFormattingNodes(root);
	const entries = Object.fromEntries(
		Object.entries(nodes).flatMap(([name, node]) => (node ? [[name, entry(node)]] : [])),
	);
	const builtInStyle = readBuiltInChartStyle(root);
	return Object.keys(entries).length
		? {
				entries,
				sourceXml: buildXml(root),
				...(builtInStyle === undefined ? {} : { builtInStyle }),
			}
		: undefined;
}
