// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes relationships added while applying edits to one package part (the document, a header,
// a footer or a notes part), plus the bytes and content types of newly inserted media.
import { bindSaveEnvironment, saveEnvironmentOf } from '../opc/reproducible';
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
import { ensureContentTypeOverride } from './zip-parts';
import type { XmlDocument } from './xml';
import { writeNewChartExPart, writeNewChartPart } from './chart-save';
import { writeNewInkPart } from './ink-save';
import { INK_RELATIONSHIP } from './ink';

const CONTENT_TYPES_PART = '[Content_Types].xml';
const extensionOf = (partName: string): string => partName.split('.').pop()?.toLowerCase() ?? '';

/** The relationships part for `partName`, e.g. word/_rels/header1.xml.rels. */
export function relationshipsPartFor(partName: string): string {
	const slash = partName.lastIndexOf('/');
	return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}

/** The part a relationship of `source` names (relative to its folder, or absolute). */
export function resolvePartTarget(source: string, target: string): string {
	if (target.startsWith('/')) return target.slice(1);
	const parts = source.split('/').slice(0, -1);
	for (const segment of target.split('/'))
		if (segment === '..') parts.pop();
		else if (segment !== '.') parts.push(segment);
	return parts.join('/');
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
	// The story's writers draw their random ids from the save's source (`saveEnvironmentOf`).
	bindSaveEnvironment(doc, saveEnvironmentOf(zip));
	const declared = parseRelationships(
		await zip.file(relationshipsPartFor(partName))?.async('string'),
	);
	const allocator = new RelationshipAllocator(
		[...scanUsedRelationshipIds(doc), ...declared.keys()],
		docPrIds,
	);
	allocator.docPrIds.reserveFromDocument(doc);
	for (const [id, rel] of declared)
		if (rel.type === INK_RELATIONSHIP && rel.mode !== 'External')
			allocator.inkParts.set(id, resolvePartTarget(partName, rel.target));
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
		else if (relationship.chartEx)
			chartTargets.set(
				relationship.id,
				relativeTarget(partName, await writeNewChartExPart(zip, relationship.chartEx)),
			);
		else if (relationship.ink)
			chartTargets.set(
				relationship.id,
				relativeTarget(partName, await writeNewInkPart(zip, relationship.ink)),
			);
	let contentTypesXml = await zip.file(CONTENT_TYPES_PART)?.async('string');
	const overrides: NewRelationship[] = [];
	for (const relationship of created) {
		relationships.set(relationship.id, {
			id: relationship.id,
			target: chartTargets.get(relationship.id) ?? relationship.target,
			mode: relationship.mode,
			type: relationship.type,
		});
		if (!relationship.partName) continue;
		if (relationship.overrideContentType) {
			const bytes =
				relationship.content ?? pendingMedia?.get(relationship.partName)?.bytes ?? undefined;
			if (bytes === undefined)
				throw new Error(`Missing content for newly inserted part: ${relationship.partName}`);
			zip.file(relationship.partName, bytes);
			overrides.push(relationship);
			continue;
		}
		const pending = pendingMedia?.get(relationship.partName);
		// A picture already in the package (a note moved to the other notes part) only needs
		// the relationship; its bytes and content type are there already.
		if (!pending && zip.file(relationship.partName)) continue;
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
	for (const { partName, overrideContentType } of overrides)
		if (partName && overrideContentType)
			await ensureContentTypeOverride(zip, partName, overrideContentType);
}
