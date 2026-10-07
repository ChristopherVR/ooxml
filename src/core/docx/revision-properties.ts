import { named, parseXml, type XmlElement } from './xml';

/** Namespace-aware comparison ignores serialization prefixes and attribute ordering. */
export function propertiesSignature(element: XmlElement): string {
	return JSON.stringify([
		element.namespaceURI,
		element.localName,
		Array.from(element.attributes)
			.filter((attribute) => attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
			.map((attribute) => [attribute.namespaceURI, attribute.localName, attribute.value])
			.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
		Array.from(element.childNodes).flatMap((child) =>
			child.nodeType === 1
				? [propertiesSignature(child as XmlElement)]
				: child.textContent?.trim()
					? [child.textContent]
					: [],
		),
	]);
}

/** Validate an immutable prior-properties subtree using the shared XML parser. */
export function parsePropertiesSnapshot(xml: string, local: 'rPr' | 'pPr'): XmlElement {
	const element = parseXml(xml).documentElement;
	if (!named(element, local))
		throw new Error(`A formatting revision snapshot must contain Word ${local} properties.`);
	return element;
}
