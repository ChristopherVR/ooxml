import { NS, buildXml, first, parseXml, type XmlElement } from '../../xml/index.js';
import type { CoreProperties } from './types.js';

const XSI = 'http://www.w3.org/2001/XMLSchema-instance';

type CoreField = keyof CoreProperties;

/** Every core field: element prefix, namespace, local name and whether it is a W3CDTF date. */
const FIELDS: readonly (readonly [CoreField, 'dc' | 'cp' | 'dcterms', boolean])[] = [
	['title', 'dc', false],
	['subject', 'dc', false],
	['creator', 'dc', false],
	['keywords', 'cp', false],
	['description', 'dc', false],
	['lastModifiedBy', 'cp', false],
	['created', 'dcterms', true],
	['modified', 'dcterms', true],
	['category', 'cp', false],
	['contentStatus', 'cp', false],
	['revision', 'cp', false],
	['language', 'dc', false],
	['identifier', 'dc', false],
	['lastPrinted', 'cp', false],
	['version', 'cp', false],
];

/** The core property names, in the order they are written. */
export const CORE_PROPERTY_FIELDS: readonly CoreField[] = FIELDS.map(([field]) => field);

const EMPTY_CORE =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	`<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="${XSI}"/>`;

/** Reads `docProps/core.xml`. Empty elements count as unset; text is not trimmed. */
export function parseCoreProperties(xml: string | undefined): CoreProperties {
	const out: CoreProperties = {};
	if (!xml) return out;
	const root = parseXml(xml, { label: 'core properties' }).documentElement;
	for (const [field, prefix] of FIELDS) {
		const text = first(root, field, NS[prefix])?.textContent ?? '';
		if (text !== '') out[field] = text;
	}
	return out;
}

/**
 * Writes `docProps/core.xml`. With `sourceXml` the source part is patched: known fields are
 * set, replaced or removed to match `props`, and every element the model does not know is kept.
 */
export function writeCoreProperties(props: CoreProperties, sourceXml?: string): string {
	const doc = parseXml(sourceXml ?? EMPTY_CORE, { label: 'core properties' });
	const root = doc.documentElement;
	for (const [field, prefix, isDate] of FIELDS) {
		const ns = NS[prefix];
		const value = props[field];
		let element: XmlElement | undefined = first(root, field, ns);
		if (value === undefined || value === '') {
			if (element && (element.textContent ?? '') !== '') root.removeChild(element);
			continue;
		}
		if (!element) {
			element = doc.createElementNS(ns, `${prefix}:${field}`);
			if (isDate) element.setAttributeNS(XSI, 'xsi:type', 'dcterms:W3CDTF');
			root.appendChild(element);
		}
		if (element.textContent !== value) element.textContent = value;
	}
	const out = buildXml(doc);
	return out.startsWith('<?xml')
		? out
		: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
}
