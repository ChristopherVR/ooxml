import { NS, elements, parseXml, relAttr, type XmlElement } from '../xml/index';
import { partNameKey } from './part-names';

export interface Relationship {
	id: string;
	type: string;
	target: string;
	/** `External` for `TargetMode="External"` (hyperlinks, linked media), else `Internal`. */
	mode: 'Internal' | 'External';
}

/** Parses a `.rels` part into a map keyed by relationship id; an absent part gives an empty map. */
export function parseRelationships(xml: string | undefined): Map<string, Relationship> {
	const map = new Map<string, Relationship>();
	if (!xml) return map;
	const root = parseXml(xml).documentElement;
	for (const node of elements(root)) {
		if (node.localName !== 'Relationship') continue;
		const id = node.getAttribute('Id');
		const target = node.getAttribute('Target');
		if (!id || target == null) continue;
		map.set(id, {
			id,
			type: node.getAttribute('Type') ?? '',
			target,
			mode: node.getAttribute('TargetMode') === 'External' ? 'External' : 'Internal',
		});
	}
	return map;
}

const RELATIONSHIP_ATTRIBUTES: ReadonlySet<string> = new Set([
	'Id',
	'Type',
	'Target',
	'TargetMode',
]);
const NO_ATTRIBUTES: ReadonlySet<string> = new Set();

const isNamespaceDeclaration = (name: string): boolean =>
	name === 'xmlns' || name.startsWith('xmlns:');

function assertOnlyAttributes(element: XmlElement, allowed: ReadonlySet<string>): void {
	for (const { name } of Array.from(element.attributes)) {
		if (!allowed.has(name) && !isNamespaceDeclaration(name)) {
			throw new Error(`Relationships part: unexpected attribute ${name}.`);
		}
	}
}

function readStrictRelationship(element: XmlElement): Relationship {
	if (element.namespaceURI !== NS.rels || element.localName !== 'Relationship') {
		throw new Error(`Relationships part: unexpected element ${element.nodeName}.`);
	}
	assertOnlyAttributes(element, RELATIONSHIP_ATTRIBUTES);
	const id = element.getAttribute('Id');
	const type = element.getAttribute('Type');
	const target = element.getAttribute('Target');
	if (!id || !type || target === null) {
		throw new Error('Relationships part: Id, Type and Target are required.');
	}
	const mode = element.getAttribute('TargetMode') ?? 'Internal';
	if (mode !== 'Internal' && mode !== 'External') {
		throw new Error(`Relationships part: invalid TargetMode ${mode}.`);
	}
	return { id, type, target, mode };
}

/**
 * Reads a relationships part for a security decision, in document order. Anything the schema
 * (ECMA-376 Part 2, 6.5.3) does not allow throws, so every reader of an accepted part sees the same
 * relationships. As in the schema, `Target` may be empty (`xsd:anyURI`) and `Id`s are unique after
 * `xsd:ID` whitespace collapse: `rId1` and `RID1` are two relationships, `rId1` and `rId1 ` are one.
 * {@link parseRelationships} is the tolerant reader for opening documents.
 */
export function parseRelationshipsPart(xml: string): Relationship[] {
	const root = parseXml(xml).documentElement;
	if (root.namespaceURI !== NS.rels || root.localName !== 'Relationships') {
		throw new Error('Not a relationships part.');
	}
	assertOnlyAttributes(root, NO_ATTRIBUTES);
	const relationships = elements(root).map(readStrictRelationship);
	const ids = new Set(relationships.map(({ id }) => id.replace(/[\t\n\r ]+/gu, ' ').trim()));
	if (ids.size !== relationships.length) {
		throw new Error('Relationships part: duplicate Id.');
	}
	return relationships;
}

/** Whether a stored item is read as a relationships part. */
export const isRelationshipsItem = (name: string): boolean => partNameKey(name).endsWith('.rels');

const escapeAttr = (value: string) =>
	value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** A relationship to serialize; its id is the key of the map it is stored in. */
export type RelationshipInput = Omit<Relationship, 'id'> & { id?: string };

/** A relationships part rewritten to exactly what {@link parseRelationshipsPart} accepted. */
export function canonicalRelationshipsXml(xml: string): string {
	const relationships = parseRelationshipsPart(xml);
	return buildRelationshipsXml(new Map(relationships.map((item) => [item.id, item])));
}

/** Serializes relationships into a `.rels` part, using each map key as the relationship id. */
export function buildRelationshipsXml(
	relationships: ReadonlyMap<string, RelationshipInput>,
): string {
	const items = Array.from(relationships)
		.map(([id, rel]) => {
			const mode = rel.mode === 'External' ? ' TargetMode="External"' : '';
			return `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(rel.type)}" Target="${escapeAttr(rel.target)}"${mode}/>`;
		})
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rels}">${items}</Relationships>`;
}

/**
 * Resolves a relationship `Target` against the part that declared it. Absolute targets (leading
 * `/`) are package-rooted; relative ones resolve against the declaring part's folder, handling
 * `.` and `..` segments. Returns a package part name without a leading slash.
 */
export function resolvePartPath(basePart: string, target: string): string {
	if (target.startsWith('/')) return target.slice(1);
	const baseDir = basePart.slice(0, basePart.lastIndexOf('/') + 1);
	const resolved: string[] = [];
	for (const segment of `${baseDir}${target}`.split('/')) {
		if (segment === '.' || segment === '') continue;
		if (segment === '..') resolved.pop();
		else resolved.push(segment);
	}
	return resolved.join('/');
}

/** The `.rels` part name that holds the relationships of `partName` (`word/document.xml` -> `word/_rels/document.xml.rels`). */
export function relationshipsPartFor(partName: string): string {
	const slash = partName.lastIndexOf('/');
	return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
}

/** The first `rIdN` that is not in `used`; the caller records the result. */
export function nextRelationshipId(used: ReadonlySet<string>): string {
	let index = 1;
	while (used.has(`rId${index}`)) index++;
	return `rId${index}`;
}

/** The `r:id` relationship reference on an element such as `w:headerReference`. */
export const getRelationshipId = (element: XmlElement): string | undefined =>
	relAttr(element, 'id') || undefined;
