import { NS, elements, first, parseXml, type XmlElement } from '../xml';
import { chartFormattingNodes } from './formatting-nodes';
import type { ChartStyleDefinition } from './style-definition';

/** Preserve whole layout/overlay subtrees when chart topology is regenerated. */
export function writeChartLayoutFormatting(
	root: XmlElement,
	formatting: ChartStyleDefinition,
): boolean {
	const nodes = chartFormattingNodes(root);
	let changed = false;
	for (const part of ['title', 'plotArea', 'legend'] as const) {
		const target = nodes[part];
		const xml = formatting.entries[part]?.sourceXml;
		if (!target || !xml?.trimStart().startsWith('<')) continue;
		const source = parseXml(xml).documentElement;
		for (const name of ['layout', 'overlay']) {
			const property = first(source, name, NS.c);
			if (!property) continue;
			const clone = root.ownerDocument.importNode(property, true);
			const current = first(target, name, NS.c);
			if (current) target.replaceChild(clone, current);
			else
				target.insertBefore(
					clone,
					elements(target).find(
						(child) =>
							child.namespaceURI === NS.c &&
							![
								'tx',
								'legendPos',
								'legendEntry',
								...(name === 'overlay' ? ['layout'] : []),
							].includes(child.localName),
					) ?? null,
				);
			changed = true;
		}
	}
	return changed;
}
