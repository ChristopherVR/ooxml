import { parseXml, selectXmlNodes, type XmlElement, type XmlNode } from '../../../xml/index';
import { parseNumberText } from '../text-number';
import { chargeArrayCells, ERR, Matrix, type Scalar } from '../values';
import { spec, str } from './helpers';

const XML = 'http://www.w3.org/XML/1998/namespace';
const XMLNS = 'http://www.w3.org/2000/xmlns/';

function preservesSpace(node: XmlNode | null): boolean {
	for (let current = node; current; current = current.parentNode) {
		if (current.nodeType !== 1) continue;
		const value = (current as XmlElement).getAttributeNS(XML, 'space');
		if (value) return value === 'preserve';
	}
	return false;
}

/** Excel ignores whitespace-only text nodes, but keeps text for XPath predicates unchanged. */
function prepare(root: XmlElement): Record<string, string> {
	const namespaces: Record<string, string> = Object.create(null);
	const visit = (element: XmlElement): void => {
		for (const attr of Array.from(element.attributes)) {
			if (attr.namespaceURI !== XMLNS || attr.localName === 'xmlns') continue;
			const prior = namespaces[attr.localName];
			if (prior !== undefined && prior !== attr.value) throw new Error('Conflicting namespaces');
			namespaces[attr.localName] = attr.value;
		}
		for (const child of Array.from(element.childNodes)) {
			if (child.nodeType === 1) visit(child as XmlElement);
			else if (child.nodeType === 3 && !child.nodeValue?.trim() && !preservesSpace(element))
				element.removeChild(child);
		}
	};
	visit(root);
	return namespaces;
}

function nodeValue(node: XmlNode): Scalar {
	let text: string;
	if (node.nodeType === 2) text = (node.nodeValue ?? '').trim();
	else if (node.nodeType === 1) {
		const preserve = preservesSpace(node);
		text = Array.from(node.childNodes)
			.filter((child) => child.nodeType === 3 || child.nodeType === 4)
			.map((child) => {
				const value = child.nodeValue ?? '';
				return preserve || child.nodeType === 4 ? value : value.trim();
			})
			.join('');
	} else return ERR.VALUE;
	if (!text) return ERR.VALUE;
	if (/^(true|false)$/i.test(text)) return text.toLowerCase() === 'true';
	// FILTERXML reads numeric dates in year/month/day order, including two-digit years.
	const date = /^(\d{1,4})[-/](\d{1,2})[-/](\d{1,4})(.*)$/.exec(text.trim());
	let numberText = text;
	if (date) {
		let year = Number(date[1]);
		if (year < 100) year += year < 30 ? 2000 : 1900;
		numberText = `${year}-${date[2]}-${date[3]}${date[4]?.replace(/^T/, ' ') ?? ''}`;
	}
	return parseNumberText(numberText) ?? text;
}

export const FILTERXML = spec(
	'FILTERXML',
	'Web',
	'FILTERXML(xml, xpath)',
	'Returns XML element or attribute values selected by XPath, as a column.',
	2,
	2,
	(args) => {
		const xml = str(args[0]);
		const expression = str(args[1]);
		if (xml.length > 32_767 || expression.length > 32_767) return ERR.VALUE;
		let nodes: XmlNode[];
		try {
			const doc = parseXml(xml, { warningsAsErrors: true, label: 'FILTERXML' });
			nodes = selectXmlNodes(expression, doc, {
				namespaces: prepare(doc.documentElement),
				matchDefaultNamespace: true,
			});
		} catch {
			return ERR.VALUE;
		}
		if (!nodes.length) return ERR.VALUE;
		chargeArrayCells(nodes.length);
		return new Matrix(nodes.map((node) => [nodeValue(node)]));
	},
);
