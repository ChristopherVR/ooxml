import type JSZip from 'jszip';
import { NS, parseXml, type XmlElement } from '../xml/index';
import { parseRelationships, resolvePartPath, type Relationship } from './relationships';

const STRICT_RELATIONSHIPS_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

/**
 * Relationship types whose target belongs to the one source that declares it: a chart, an
 * embedded package, an OLE object, a diagram's parts, a slide's notes, a master's theme and
 * layouts. Each is the last path segment of the type URI. Images, media, fonts and hyperlinks are
 * absent: any number of sources may point at one of those.
 */
const OWNED_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
	'chart',
	'chartEx',
	'chartUserShapes',
	'chartColorStyle',
	'chartStyle',
	'package',
	'oleObject',
	'diagramData',
	'diagramLayout',
	'diagramQuickStyle',
	'diagramColors',
	'diagramDrawing',
	'notesSlide',
	'comments',
	'themeOverride',
	'theme',
	'slideLayout',
]);

/**
 * Sources whose relationship of a type names a part another source owns. The presentation part
 * points at the first master's theme, and every slide at the layout its master owns.
 */
const REFERENCING_SOURCES: Readonly<Record<string, RegExp>> = {
	theme: /^ppt\/presentation\.xml$/u,
	slideLayout: /^ppt\/slides\/[^/]+$/u,
};

/**
 * Owned types the declaring part's markup never cites by id: the relationship itself is the claim.
 * Every other owned relationship claims once per citing element, and not at all when nothing
 * cites it (a relationship left behind by a deleted or moved element, which PowerPoint ignores).
 */
const UNCITED_TYPES: ReadonlySet<string> = new Set([
	'notesSlide',
	'comments',
	'theme',
	'themeOverride',
	'diagramDrawing',
]);

const typeName = (relationship: Relationship): string =>
	relationship.type.slice(relationship.type.lastIndexOf('/') + 1);

/** Whether `relationship`, declared by `source`, makes `source` the owner of its target. */
export function ownsRelationshipTarget(source: string, relationship: Relationship): boolean {
	if (relationship.mode === 'External') return false;
	const type = typeName(relationship);
	return OWNED_RELATIONSHIP_TYPES.has(type) && !REFERENCING_SOURCES[type]?.test(source);
}

/** One claim on an owned part: the declaring part and relationship id. */
export interface PartOwnerClaim {
	source: string;
	relationshipId: string;
}

/** An owned part that more than one owner claims. */
export interface SharedOwnedPart {
	part: string;
	claims: PartOwnerClaim[];
}

/** `mc:Fallback` is another rendering of the element its `mc:Choice` holds, not a second owner. */
function countRelationshipReferences(root: XmlElement, counts: Map<string, number>): void {
	for (const attribute of Array.from(root.attributes)) {
		if (attribute.namespaceURI === NS.r || attribute.namespaceURI === STRICT_RELATIONSHIPS_NS)
			counts.set(attribute.value, (counts.get(attribute.value) ?? 0) + 1);
	}
	for (const child of Array.from(root.children)) {
		if (child.namespaceURI === NS.mc && child.localName === 'Fallback') continue;
		countRelationshipReferences(child, counts);
	}
}

async function relationshipReferenceCounts(
	zip: JSZip,
	source: string,
): Promise<Map<string, number>> {
	const counts = new Map<string, number>();
	if (!source.endsWith('.xml')) return counts;
	const xml = await zip.file(source)?.async('string');
	if (xml) countRelationshipReferences(parseXml(xml).documentElement, counts);
	return counts;
}

/**
 * Every owned part (see {@link ownsRelationshipTarget}) with more than one owner. Each element of
 * the source that names the relationship is an owner, so two chart frames on one slide sharing a
 * relationship are two claims on the chart. PowerPoint refuses such a package (0x80070570), and
 * an edit through one owner would silently change the other's content.
 */
export async function findSharedOwnedParts(zip: JSZip): Promise<SharedOwnedPart[]> {
	const claims = new Map<string, PartOwnerClaim[]>();
	for (const name of Object.keys(zip.files)) {
		const match = /^(?<dir>(?:.*\/)?)_rels\/(?<file>[^/]+)\.rels$/u.exec(name);
		if (!match?.groups || zip.files[name]?.dir) continue;
		const source = `${match.groups['dir']}${match.groups['file']}`;
		const relationships = parseRelationships(await zip.file(name)?.async('string'));
		const owned = [...relationships.values()].filter((relationship) =>
			ownsRelationshipTarget(source, relationship),
		);
		if (owned.length === 0) continue;
		const references = await relationshipReferenceCounts(zip, source);
		for (const relationship of owned) {
			const part = resolvePartPath(source, relationship.target);
			const list = claims.get(part) ?? [];
			const owners = UNCITED_TYPES.has(typeName(relationship))
				? 1
				: (references.get(relationship.id) ?? 0);
			for (let index = 0; index < owners; index++)
				list.push({ source, relationshipId: relationship.id });
			claims.set(part, list);
		}
	}
	return [...claims]
		.filter(([, list]) => list.length > 1)
		.map(([part, list]) => ({ part, claims: list }));
}
