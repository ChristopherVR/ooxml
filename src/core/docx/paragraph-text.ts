import { runContentText } from './special-characters';
import { isElement, named, type XmlElement } from './xml';

/** Visible text leaves in a paragraph, retaining the source nodes for text-only patches. */
export function paragraphTextNodes(paragraph: XmlElement): XmlElement[] {
	const nodes: XmlElement[] = [];
	const visit = (parent: XmlElement) => {
		for (const child of Array.from(parent.childNodes).filter(isElement)) {
			if (
				['pPr', 'rPr', 'p', 'drawing', 'pict', 'del', 'moveFrom', 'delText'].some((name) =>
					named(child, name),
				)
			)
				continue;
			if (runContentText(child) !== undefined) nodes.push(child);
			else visit(child);
		}
	};
	visit(paragraph);
	return nodes;
}

export const paragraphText = (paragraph: XmlElement): string =>
	paragraphTextNodes(paragraph)
		.map((node) => runContentText(node)!)
		.join('');
