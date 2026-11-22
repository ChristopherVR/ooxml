import { makeNS, parseXml, type XmlElement } from '../xml/index';

export const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';

export function createSvg(): XmlElement {
	return parseXml(`<svg xmlns="${SVG_NAMESPACE}"/>`).documentElement;
}

export function setSvgAttributes(
	element: XmlElement,
	attributes: Record<string, string | number>,
): void {
	for (const [name, value] of Object.entries(attributes)) element.setAttribute(name, String(value));
}

export function appendSvg(
	parent: XmlElement,
	name: string,
	attributes: Record<string, string | number> = {},
): XmlElement {
	const node = makeNS(parent.ownerDocument, SVG_NAMESPACE, name);
	setSvgAttributes(node, attributes);
	return parent.appendChild(node);
}
