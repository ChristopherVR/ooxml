import type JSZip from 'jszip';

import { buildXml, parseXml, type XmlDocument, type XmlElement } from '../../../xml/index';
import type { PptxCustomerData, XmlObject } from '../types';
import {
	custDataContainer,
	ensureCustDataList,
	listChildren,
	removeIfEmpty,
} from './cust-data-list-dom';
import { safeResolveZipPath } from './safe-path';

export const CUSTOM_XML_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const DEFAULT_CONTENT_TYPE = 'application/xml';

interface XmlCodec {
	parse(xml: string): XmlObject;
	build(data: XmlObject): string;
}

export interface CustomerDataScope {
	sourcePartPath: string;
	location: 'presentation' | 'slide';
	entries: PptxCustomerData[];
}

export async function writeCustomerDataScopes(
	zip: JSZip,
	scopes: CustomerDataScope[],
	codec: XmlCodec,
): Promise<void> {
	let nextIndex = maxCustomXmlIndex(zip) + 1;
	const contentTypes = new Map<string, string>();
	for (const scope of scopes) {
		const sourceXml = await zip.file(scope.sourcePartPath)?.async('string');
		if (!sourceXml) {
			continue;
		}
		// Edited as a DOM: rebuilding the written part through the grouping
		// object model would regroup its interleaved siblings (see cust-data-list-dom).
		const sourceDoc = parseXml(sourceXml, { label: scope.sourcePartPath });
		const container = custDataContainer(sourceDoc);
		if (!container || !matchesLocation(container, scope.location)) {
			continue;
		}
		const relsPath = relsForSourcePart(scope.sourcePartPath);
		const relsXml = await zip.file(relsPath)?.async('string');
		const relsData = relsXml
			? codec.parse(relsXml)
			: {
					Relationships: {
						'@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
					},
				};
		const relsRoot = (relsData['Relationships'] ??= {}) as XmlObject;
		const relationships = ensureArray(relsRoot['Relationship']);
		const authoredEntries: XmlObject[] = [];
		for (const entry of scope.entries) {
			const suppliedPath = entry.id ? safeResolveZipPath('', entry.id) : null;
			const path = suppliedPath || `customXml/item${nextIndex++}.xml`;
			if (entry.data === undefined && !zip.file(path)) {
				continue;
			}
			entry.id = path;
			entry.contentType ??= DEFAULT_CONTENT_TYPE;
			if (entry.data !== undefined) {
				zip.file(path, entry.data);
			}
			const relationship = upsertRelationship(
				relationships,
				entry,
				relativeTarget(scope.sourcePartPath, path),
			);
			entry.relId = String(relationship['@_Id']);
			const rawEntry = { ...(entry.rawXml ?? {}), '@_r:id': entry.relId };
			entry.rawXml = rawEntry;
			authoredEntries.push(rawEntry);
			contentTypes.set(path, entry.contentType);
		}
		replaceCustomerDataEntries(sourceDoc, container, authoredEntries, codec);
		relsRoot['Relationship'] = relationships;
		zip.file(relsPath, codec.build(relsData));
		zip.file(scope.sourcePartPath, buildXml(sourceDoc));
	}
	await upsertContentTypes(zip, codec, contentTypes);
}

export async function resolveContentType(
	zip: JSZip,
	codec: Pick<XmlCodec, 'parse'>,
	path: string,
): Promise<string | undefined> {
	const xml = await zip.file('[Content_Types].xml')?.async('string');
	if (!xml) {
		return undefined;
	}
	const root = codec.parse(xml)['Types'] as XmlObject | undefined;
	const override = ensureArray(root?.['Override']).find(
		(entry) => entry['@_PartName'] === `/${path}`,
	);
	if (override?.['@_ContentType']) {
		return String(override['@_ContentType']);
	}
	const extension = path.split('.').pop()?.toLowerCase();
	const fallback = ensureArray(root?.['Default']).find(
		(entry) => String(entry['@_Extension'] ?? '').toLowerCase() === extension,
	);
	return fallback?.['@_ContentType'] ? String(fallback['@_ContentType']) : undefined;
}

/** Whether the part's `p:custDataLst` container is the one `location` names. */
function matchesLocation(container: XmlElement, location: CustomerDataScope['location']): boolean {
	const rootName = container.ownerDocument?.documentElement?.localName;
	return rootName === (location === 'presentation' ? 'presentation' : 'sld');
}

function upsertRelationship(
	relationships: XmlObject[],
	entry: PptxCustomerData,
	target: string,
): XmlObject {
	let relationship = relationships.find(
		(item) => String(item['@_Id'] ?? '') === entry.relId && isCustomXmlRelationship(item['@_Type']),
	);
	if (!relationship) {
		entry.relId = nextRelationshipId(relationships);
		relationship = {};
		relationships.push(relationship);
	}
	relationship['@_Id'] = entry.relId!;
	relationship['@_Type'] = CUSTOM_XML_RELATIONSHIP_TYPE;
	relationship['@_Target'] = target;
	delete relationship['@_TargetMode'];
	return relationship;
}

function isCustomXmlRelationship(value: unknown): boolean {
	return String(value ?? '').endsWith('/relationships/customXml');
}

/**
 * Make `entries` the container list's `p:custData` children, ahead of any
 * `p:tags` (CT_CustomerDataList is `custData*, tags?`), leaving everything else
 * in the part untouched. Each entry is serialized through `codec` so its
 * preserved unknown attributes and children survive, then imported into the
 * part under the part's own namespace declarations.
 */
function replaceCustomerDataEntries(
	doc: XmlDocument,
	container: XmlElement,
	entries: XmlObject[],
	codec: XmlCodec,
): void {
	const list = ensureCustDataList(doc, container);
	for (const stale of listChildren(list, 'custData')) {
		list.removeChild(stale);
	}
	const anchor = listChildren(list, 'tags')[0] ?? null;
	for (const entry of entries) {
		list.insertBefore(customerDataElement(doc, entry, codec), anchor);
	}
	removeIfEmpty(list);
}

function customerDataElement(doc: XmlDocument, entry: XmlObject, codec: XmlCodec): XmlElement {
	const root = doc.documentElement!;
	const declarations = Array.from(root.attributes)
		.filter((attribute) => attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:'))
		.map((attribute) => `${attribute.name}="${attribute.value}"`)
		.join(' ');
	const fragment = parseXml(
		`<wrap ${declarations}>${codec.build({ 'p:custData': entry })}</wrap>`,
		{
			label: 'p:custData',
		},
	);
	return doc.importNode(fragment.documentElement!.firstChild!, true) as XmlElement;
}

async function upsertContentTypes(
	zip: JSZip,
	codec: XmlCodec,
	contentTypes: Map<string, string>,
): Promise<void> {
	const xml = await zip.file('[Content_Types].xml')?.async('string');
	if (!xml || contentTypes.size === 0) {
		return;
	}
	const data = codec.parse(xml);
	const root = data['Types'] as XmlObject;
	const overrides = ensureArray(root['Override']);
	for (const [path, contentType] of contentTypes) {
		const partName = `/${path}`;
		const existing = overrides.find((entry) => entry['@_PartName'] === partName);
		if (existing) {
			existing['@_ContentType'] = contentType;
		} else {
			overrides.push({ '@_PartName': partName, '@_ContentType': contentType });
		}
	}
	root['Override'] = overrides;
	zip.file('[Content_Types].xml', codec.build(data));
}

function relsForSourcePart(source: string): string {
	const slash = source.lastIndexOf('/');
	return `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

function relativeTarget(source: string, target: string): string {
	const from = source.split('/').slice(0, -1);
	const to = target.split('/');
	while (from.length > 0 && to.length > 0 && from[0] === to[0]) {
		from.shift();
		to.shift();
	}
	return [...from.map(() => '..'), ...to].join('/');
}

function nextRelationshipId(relationships: XmlObject[]): string {
	const used = new Set(relationships.map((entry) => String(entry['@_Id'] ?? '')));
	let index = 1;
	while (used.has(`rId${index}`)) {
		index += 1;
	}
	return `rId${index}`;
}

function maxCustomXmlIndex(zip: JSZip): number {
	return zip
		.file(/^customXml\/item\d+\.xml$/u)
		.reduce((max, file) => Math.max(max, Number(file.name.match(/item(\d+)\.xml$/u)?.[1] ?? 0)), 0);
}

function ensureArray(value: unknown): XmlObject[] {
	if (value === undefined || value === null) {
		return [];
	}
	return (Array.isArray(value) ? value : [value]) as XmlObject[];
}
