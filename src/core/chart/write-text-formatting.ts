import { NS, elements, first, parseXml, type XmlElement } from '../xml';
import { chartFormattingNodes } from './formatting-nodes';
import type { ChartStyleDefinition } from './style-definition';

const rich = (node: XmlElement) => first(first(node, 'tx', NS.c), 'rich', NS.c);

/** Preserve imported text bodies/properties when regenerating chart topology. */
export function writeChartTextFormatting(
	root: XmlElement,
	formatting: ChartStyleDefinition,
): boolean {
	const doc = root.ownerDocument;
	let changed = false;
	const copy = (
		source: XmlElement | undefined,
		target: XmlElement,
		name: string,
		namespace = NS.a,
	) => {
		const property = first(source, name, namespace);
		if (!property) return;
		changed = true;
		const clone = doc.importNode(property, true);
		const current = first(target, name, namespace);
		if (current) target.replaceChild(clone, current);
		else target.insertBefore(clone, elements(target)[0] ?? null);
	};
	for (const [part, node] of Object.entries(chartFormattingNodes(root))) {
		const xml = formatting.entries[part]?.sourceXml;
		if (!node || !xml?.trimStart().startsWith('<')) continue;
		const source = parseXml(xml).documentElement;
		const text = first(source, 'txPr', NS.c);
		if (text) {
			changed = true;
			const current = first(node, 'txPr', NS.c);
			const clone = doc.importNode(text, true);
			if (current) node.replaceChild(clone, current);
			else
				node.insertBefore(
					clone,
					elements(node).find(
						(child) =>
							child.namespaceURI === NS.c &&
							['crossAx', 'externalData', 'printSettings', 'userShapes', 'extLst'].includes(
								child.localName,
							),
					) ?? null,
				);
		}
		if (part !== 'title') continue;
		const sourceRich = rich(source);
		const targetRich = rich(node);
		if (!sourceRich || !targetRich) continue;
		copy(sourceRich, targetRich, 'bodyPr');
		copy(sourceRich, targetRich, 'lstStyle');
		const sourceParagraph = first(sourceRich, 'p', NS.a);
		const targetParagraph = first(targetRich, 'p', NS.a);
		if (!targetParagraph) continue;
		copy(sourceParagraph, targetParagraph, 'pPr');
		const targetRun = first(targetParagraph, 'r', NS.a);
		if (targetRun) copy(first(sourceParagraph, 'r', NS.a), targetRun, 'rPr');
	}
	return changed;
}
