import { elements, WORD_NS, type XmlDocument, type XmlElement } from './xml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const XMLNS = 'http://www.w3.org/2000/xmlns/';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
function prefix(doc: XmlDocument, preferred: string, namespace: string): string {
	let name = preferred;
	for (
		let suffix = 1;
		doc.documentElement.lookupNamespaceURI(name) &&
		doc.documentElement.lookupNamespaceURI(name) !== namespace;
		suffix++
	)
		name = preferred + suffix;
	doc.documentElement.setAttributeNS(XMLNS, 'xmlns:' + name, namespace);
	return name;
}
export function writeFootnoteColumns(
	doc: XmlDocument,
	section: XmlElement,
	value: number | undefined,
): void {
	if (value !== undefined && (!Number.isInteger(value) || value < 0 || value > 2147483647))
		throw new Error('Footnote column count must be an integer from 0 to 2147483647.');
	const old = elements(section).filter(
		(element) => element.namespaceURI === W15 && element.localName === 'footnoteColumns',
	);
	for (const element of old) section.removeChild(element);
	if (value === undefined) return;
	const w15 = prefix(doc, 'w15', W15),
		mc = prefix(doc, 'mc', MC);
	const ignored = new Set(
		(doc.documentElement.getAttributeNS(MC, 'Ignorable') ?? '').split(/\s+/).filter(Boolean),
	);
	ignored.add(w15);
	doc.documentElement.setAttributeNS(MC, mc + ':Ignorable', [...ignored].join(' '));
	const element = doc.createElementNS(W15, w15 + ':footnoteColumns');
	element.setAttributeNS(WORD_NS, 'w:val', String(value));
	section.appendChild(element);
}
