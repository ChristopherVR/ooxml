import { NS, first, parseXml, type XmlElement } from '../../xml/index';
import { PartPatch, escapeText } from './patch';
import type { CoreProperties } from './types';

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

/** `date` as W3CDTF at second precision in UTC (`2026-10-08T09:30:00Z`), as Office writes it. */
export const formatW3cdtf = (date: Date): string => date.toISOString().replace(/\.\d{3}Z$/, 'Z');

/**
 * Writes `docProps/core.xml`. With `sourceXml` the source part is patched: known fields are
 * set, replaced or removed to match `props`, and everything else is kept byte for byte (see
 * `PartPatch`). A date that is written carries `xsi:type="dcterms:W3CDTF"` as ECMA-376 Part 2
 * requires, and any `xsi` or `dcterms` declaration that needs is added to the root.
 */
export function writeCoreProperties(props: CoreProperties, sourceXml?: string): string {
	const xml = sourceXml ?? EMPTY_CORE;
	const root = parseXml(xml, { label: 'core properties' }).documentElement;
	const patch = new PartPatch(xml, root);
	for (const [field, prefix, isDate] of FIELDS) {
		const ns = NS[prefix];
		const value = props[field];
		const element: XmlElement | undefined = first(root, field, ns);
		if (value === undefined || value === '') {
			if (element && (element.textContent ?? '') !== '') patch.remove(element);
			continue;
		}
		if (element && element.textContent === value) continue;
		const typed = () =>
			` ${patch.prefix(XSI, 'xsi')}:type="${patch.prefix(NS.dcterms, 'dcterms')}:W3CDTF"`;
		if (element) {
			if (isDate && !element.getAttributeNS(XSI, 'type')) patch.addAttribute(element, typed());
			patch.setContent(element, escapeText(value));
			continue;
		}
		const name = patch.qualify(ns, field, prefix);
		patch.insert(`<${name}${isDate ? typed() : ''}>${escapeText(value)}</${name}>`);
	}
	return patch.toString();
}
