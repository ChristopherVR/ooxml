import { saveEnvironmentOf } from '../opc/reproducible';
import type JSZip from 'jszip';
import { createOfficeGuid } from '../crypto/uuid';
import {
	declaredEntrySize,
	ensureContentTypeOverride,
	ensureRelationship,
	relationshipsPartFor,
} from '../opc/index';
import { parseXml, type XmlElement } from './xml';

export const BIBLIOGRAPHY_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
export const BIBLIOGRAPHY_XML_LIMIT = 2_000_000;

/** The XML is retained until a References command needs to interpret or edit it. */
export interface BibliographyPart {
	part?: string;
	xml: string;
}

const CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * The `<Sources>` root of a bibliography part, or undefined for any other custom XML. The parsed
 * root's namespace decides; text that merely mentions the namespace never does.
 */
export function bibliographySourcesRoot(xml: string): XmlElement | undefined {
	// A cheap skip for parts that cannot hold a `<Sources>` root.
	if (xml.length > BIBLIOGRAPHY_XML_LIMIT || !xml.includes('Sources')) return undefined;
	try {
		const root = parseXml(xml).documentElement;
		if (root?.namespaceURI === BIBLIOGRAPHY_NS && root.localName === 'Sources') return root;
	} catch {
		// Other custom XML parts stay opaque.
	}
	return undefined;
}

export async function readBibliographyPart(
	zip: JSZip,
	parts: string[],
): Promise<BibliographyPart | undefined> {
	for (const part of parts) {
		const file = zip.file(part);
		if (!file || declaredEntrySize(file) > BIBLIOGRAPHY_XML_LIMIT) continue;
		const xml = await file.async('string');
		if (bibliographySourcesRoot(xml)) return { part, xml };
	}
	return undefined;
}

export async function writeBibliographyPart(
	zip: JSZip,
	part: BibliographyPart | undefined,
	prior: BibliographyPart | undefined,
): Promise<void> {
	if (!part || (prior?.part === part.part && prior?.xml === part.xml)) return;
	if (part.xml.length > BIBLIOGRAPHY_XML_LIMIT)
		throw new Error('Bibliography source list is too large');
	if (prior?.part && zip.file(prior.part)) {
		zip.file(prior.part, part.xml);
		return;
	}
	let number = 1;
	while (
		zip.file(`customXml/item${number}.xml`) ||
		zip.file(`customXml/itemProps${number}.xml`) ||
		zip.file(`customXml/_rels/item${number}.xml.rels`)
	)
		number++;
	const item = `customXml/item${number}.xml`;
	const props = `itemProps${number}.xml`;
	const id = createOfficeGuid(saveEnvironmentOf(zip).ids);
	zip.file(item, part.xml);
	zip.file(
		`customXml/${props}`,
		`<?xml version="1.0" encoding="UTF-8"?><ds:datastoreItem xmlns:ds="${CUSTOM_XML}" ds:itemID="${id}"><ds:schemaRefs><ds:schemaRef ds:uri="${BIBLIOGRAPHY_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
	);
	await ensureRelationship(zip, 'word/_rels/document.xml.rels', `${REL}/customXml`, `../${item}`);
	await ensureRelationship(zip, relationshipsPartFor(item), `${REL}/customXmlProps`, props);
	await ensureContentTypeOverride(
		zip,
		`customXml/${props}`,
		'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
	);
}
