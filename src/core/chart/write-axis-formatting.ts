import { NS, elements, first, type XmlElement } from '../xml';
import type { ChartStyleDefinition, ChartStyleEntry } from './style-definition';
import { chartFormattingNodes } from './formatting-nodes';
import { readChartFormatting } from './read-formatting';

/** Patch explicit axis flags without rebuilding axes, styling or extension data. */
export function writeChartAxisFormatting(
	root: XmlElement,
	formatting: ChartStyleDefinition | undefined,
): boolean {
	if (!formatting) return false;
	const nodes = chartFormattingNodes(root);
	const previous = readChartFormatting(root);
	let changed = false;
	for (const part of ['categoryAxis', 'valueAxis'] as const) {
		const axis = nodes[part];
		const entry: ChartStyleEntry | undefined = formatting.entries[part];
		if (!axis || !entry) continue;
		for (const [name, flag, value, successors] of [
			[
				'delete',
				entry.axisVisible,
				entry.axisVisible ? '0' : '1',
				[
					'axPos',
					'majorGridlines',
					'minorGridlines',
					'title',
					'numFmt',
					'majorTickMark',
					'minorTickMark',
					'tickLblPos',
					'spPr',
					'txPr',
					'crossAx',
					'crosses',
					'crossesAt',
					'extLst',
				],
			],
			[
				'tickLblPos',
				entry.labelsVisible,
				entry.labelsVisible ? 'nextTo' : 'none',
				[
					'spPr',
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
					'extLst',
				],
			],
		] as const) {
			if (flag === undefined) continue;
			const key = name === 'delete' ? 'axisVisible' : 'labelsVisible';
			const revealingAxis =
				name === 'tickLblPos' &&
				entry.axisVisible === true &&
				previous?.entries[part]?.axisVisible === false;
			if (flag === previous?.entries[part]?.[key] && !revealingAxis) continue;
			let node = first(axis, name, NS.c);
			const old = node?.getAttribute('val');
			// Retain equivalent boolean spellings and high/low label positions when visible.
			if (
				name === 'delete' &&
				old !== undefined &&
				old !== null &&
				((flag && (old === '0' || old === 'false')) || (!flag && (old === '1' || old === 'true')))
			)
				continue;
			if (name === 'tickLblPos' && node && (flag ? old !== 'none' : old === 'none')) continue;
			if (!node) {
				node = root.ownerDocument.createElementNS(NS.c, 'c:' + name);
				const before = elements(axis).find(
					(child) =>
						child.namespaceURI === NS.c &&
						(successors as readonly string[]).includes(child.localName),
				);
				axis.insertBefore(node, before ?? null);
			}
			node.setAttribute('val', value);
			changed = true;
		}
	}
	return changed;
}
