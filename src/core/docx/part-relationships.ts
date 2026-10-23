// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes relationships added while applying edits to one package part (the document, a header,
// a footer or a notes part), plus the bytes and content types of newly inserted media.
import type JSZip from 'jszip';
import type { PendingMediaPart } from './model';
import {
	buildRelationshipsXml,
	ensureContentTypeDefault,
	parseRelationships,
} from './package-parts';
import {
	RelationshipAllocator,
	scanUsedRelationshipIds,
	type NewRelationship,
} from './relationship-allocator';
import type { DocPrIdAllocator } from './docpr-ids';
import type { XmlDocument } from './xml';
import { writeNewChartPart } from './chart-save';

const CONTENT_TYPES_PART = '[Content_Types].xml';
const extensionOf = (partName: string): string => partName.split('.').pop()?.toLowerCase() ?? '';

/** The relationships part for `partName`, e.g. word/_rels/header1.xml.rels. */
export function relationshipsPartFor(partName: string): string {
	const slash = partName.lastIndexOf('/');
	return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}

/** `target` as a relationship target of `source`: relative when it is in the same folder or below. */
function relativeTarget(source: string, target: string): string {
	const folder = source.slice(0, source.lastIndexOf('/') + 1);
	return target.startsWith(folder) ? target.slice(folder.length) : `/${target}`;
}

/** An allocator that avoids every id already used by the part's XML or declared in its .rels. */
export async function allocatorForPart(
	zip: JSZip,
	partName: string,
	doc: XmlDocument,
	docPrIds?: DocPrIdAllocator,
): Promise<RelationshipAllocator> {
	const declared = parseRelationships(
		await zip.file(relationshipsPartFor(partName))?.async('string'),
	);
	const allocator = new RelationshipAllocator(
		[...scanUsedRelationshipIds(doc), ...declared.keys()],
		docPrIds,
	);
	allocator.docPrIds.reserveFromDocument(doc);
	return allocator;
}

/** Adds `created` relationships to the part's .rels and writes any newly inserted media. */
export async function writeNewRelationships(
	zip: JSZip,
	partName: string,
	created: readonly NewRelationship[],
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
): Promise<void> {
	if (!created.length) return;
	const relsPath = relationshipsPartFor(partName);
	const relationships = parseRelationships(await zip.file(relsPath)?.async('string'));
	// New charts write their own parts and content types first, so the content types read below include them.
	const chartTargets = new Map<string, string>();
	for (const relationship of created)
		if (relationship.chart)
			chartTargets.set(
				relationship.id,
				relativeTarget(partName, await writeNewChartPart(zip, relationship.chart)),
			);
	let contentTypesXml = await zip.file(CONTENT_TYPES_PART)?.async('string');
	for (const relationship of created) {
		relationships.set(relationship.id, {
			id: relationship.id,
			target: chartTargets.get(relationship.id) ?? relationship.target,
			mode: relationship.mode,
			type: relationship.type,
		});
		if (!relationship.partName) continue;
		const pending = pendingMedia?.get(relationship.partName);
		if (!pending)
			throw new Error(`Missing bytes for newly inserted media part: ${relationship.partName}`);
		zip.file(relationship.partName, pending.bytes);
		contentTypesXml = ensureContentTypeDefault(
			contentTypesXml,
			extensionOf(relationship.partName),
			pending.contentType,
		);
	}
	zip.file(relsPath, buildRelationshipsXml(relationships));
	if (contentTypesXml) zip.file(CONTENT_TYPES_PART, contentTypesXml);
}
