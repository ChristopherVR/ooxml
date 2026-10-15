import { elements, parseXml, type XmlElement } from '../xml/index';
import { mergeSiblings } from './latex-omml-siblings';
import type { OmmlNode } from './omml-node';

const PREFIXES: Readonly<Record<string, string>> = {
	'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
	'http://purl.oclc.org/ooxml/officeDocument/math': 'm',
	'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
	'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
};

/** Deepest element nesting converted; real equations stay far below this. */
const MAX_DEPTH = 512;

/** Adapt an ordered, namespace-aware OOXML DOM without a product parser. */
export function ommlFromElement(element: XmlElement): OmmlNode {
	const visit = (node: XmlElement, depth: number): OmmlNode => {
		if (depth > MAX_DEPTH) throw new Error(`OMML nesting exceeds ${MAX_DEPTH} levels`);
		const prefix = PREFIXES[node.namespaceURI ?? ''];
		const name = prefix ? `${prefix}:${node.localName}` : node.tagName;
		const body = mergeSiblings(elements(node).map((child) => visit(child, depth + 1)));
		for (const attribute of Array.from(node.attributes)) {
			if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
			body[`@_${attribute.localName}`] = attribute.value;
		}
		const text = Array.from(node.childNodes)
			.filter((child) => child.nodeType === 3 || child.nodeType === 4)
			.map((child) => child.nodeValue ?? '')
			.join('');
		if (text && (!elements(node).length || text.trim())) body['#text'] = text;
		return { [name]: body };
	};
	return visit(element, 1);
}

/** Parse standalone OMML using the shared XML parser's validation and DTD rejection. */
export function parseOmml(xml: string): OmmlNode {
	return ommlFromElement(parseXml(xml, { label: 'OMML' }).documentElement);
}
