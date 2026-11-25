import { elements, parseXml, type XmlElement } from '../xml/index';
import { mergeSiblings } from './latex-omml-siblings';
import type { OmmlNode } from './omml-node';
import { checkEquationTree, MAX_OMML_LENGTH, MAX_OMML_NODES } from './limits';

const PREFIXES: Readonly<Record<string, string>> = {
	'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
	'http://purl.oclc.org/ooxml/officeDocument/math': 'm',
	'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
	'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
};

/** Adapt an ordered, namespace-aware OOXML DOM without a product parser. */
export function ommlFromElement(element: XmlElement): OmmlNode {
	checkEquationTree(element);
	const visit = (node: XmlElement): OmmlNode => {
		const prefix = PREFIXES[node.namespaceURI ?? ''];
		const name = prefix ? `${prefix}:${node.localName}` : node.tagName;
		const children = elements(node);
		const body = mergeSiblings(children.map(visit));
		for (const attribute of Array.from(node.attributes)) {
			if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
			body[`@_${attribute.localName}`] = attribute.value;
		}
		const text = Array.from(node.childNodes)
			.filter((child) => child.nodeType === 3 || child.nodeType === 4)
			.map((child) => child.nodeValue ?? '')
			.join('');
		if (text && (!children.length || text.trim())) body['#text'] = text;
		return { [name]: body };
	};
	return visit(element);
}

/** Parse standalone OMML using the shared XML parser's validation and DTD rejection. */
export function parseOmml(xml: string): OmmlNode {
	if (xml.length > MAX_OMML_LENGTH) throw new Error('OMML exceeds the size limit');
	return ommlFromElement(
		parseXml(xml, { label: 'OMML', limits: { maxNodes: MAX_OMML_NODES } }).documentElement,
	);
}
