import { NS, elements, parseXml, relAttr, type XmlElement } from '@christophervr/ooxml-xml';

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

const escapeAttr = (value: string) =>
	value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** A relationship to serialize; its id is the key of the map it is stored in. */
export type RelationshipInput = Omit<Relationship, 'id'> & { id?: string };

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
