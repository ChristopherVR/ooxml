import { named, parseXml, type XmlElement } from './xml';

/** Validate an immutable prior-properties subtree using the shared XML parser. */
export function parsePropertiesSnapshot(xml: string, local: 'rPr' | 'pPr'): XmlElement {
	const element = parseXml(xml).documentElement;
	if (!named(element, local))
		throw new Error(`A formatting revision snapshot must contain Word ${local} properties.`);
	return element;
}
