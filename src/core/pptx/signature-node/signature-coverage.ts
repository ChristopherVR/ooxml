/**
 * Decides whether verified references cover a package's content.
 *
 * OPC signatures may sign a subset of a package, and Office leaves document properties, thumbnails
 * and a few settings parts out. Coverage therefore walks from the package relationships, following
 * only relationships the signature protects, and requires every part reached, and every
 * relationships part beside it, to be covered. It then closes the package: any other stored item
 * is a gap unless it sits where those unsigned parts live, because a reader may find parts by name
 * without following a relationship.
 *
 * The exemption is about the part and does not depend on whether its relationship is signed.
 * ECMA-376 Part 2, 10.3 leaves what to sign to a signature policy and nothing requires the target
 * of a signed relationship to be signed, so a signed relationship to unsigned custom XML is not a
 * gap. Such parts are always listed in `uncoveredParts`.
 *
 * Node-only - reads parts through `jszip`.
 */

import { MEDIA_TYPES_ITEM, partNameKey } from '../../opc/part-names';
import { relationshipsPartFor, resolvePartPath } from '../../opc/relationships';
import { isSignaturePart } from '../../opc/signature/detection';
import type { PackageRelationship, SignatureCoverageGap } from '../core/utils/signature-types';
import { resolvePartName } from './reference-checks';
import type { CoveredPart, SignaturePackage } from './reference-checks';
import { parseRelationshipsPart } from './reference-transforms';

const PACKAGE_RELATIONSHIPS_PART = '_rels/.rels';
const PACKAGE_TYPES = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_TYPES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_OFFICE_TYPES = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

/** Relationship types Office leaves unsigned: metadata and viewing state, not document content. */
const UNSIGNED_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
	`${PACKAGE_TYPES}/metadata/core-properties`,
	`${PACKAGE_TYPES}/metadata/thumbnail`,
	`${PACKAGE_TYPES}/digital-signature/origin`,
	`${OFFICE_TYPES}/metadata/core-properties`,
	`${OFFICE_TYPES}/metadata/thumbnail`,
	`${OFFICE_TYPES}/extended-properties`,
	`${OFFICE_TYPES}/custom-properties`,
	`${OFFICE_TYPES}/customXml`,
	`${OFFICE_TYPES}/presProps`,
	`${OFFICE_TYPES}/viewProps`,
	`${STRICT_OFFICE_TYPES}/metadata/core-properties`,
	`${STRICT_OFFICE_TYPES}/metadata/thumbnail`,
	`${STRICT_OFFICE_TYPES}/extendedProperties`,
	`${STRICT_OFFICE_TYPES}/customProperties`,
	`${STRICT_OFFICE_TYPES}/customXml`,
	`${STRICT_OFFICE_TYPES}/presProps`,
	`${STRICT_OFFICE_TYPES}/viewProps`,
]);

/** Where the targets of those relationships live, as part name keys. */
const UNSIGNED_PART_LOCATION =
	/^(?:docprops\/|customxml\/|ppt\/(?:presprops|viewprops)\.xml$|ppt\/_rels\/(?:presprops|viewprops)\.xml\.rels$)/;

const MEDIA_TYPES_KEY = partNameKey(MEDIA_TYPES_ITEM);

/** ZIP items OPC does not map to parts: the Media Types stream and deleted-part leftovers. */
const isNotAPart = (key: string): boolean => key === MEDIA_TYPES_KEY || key.startsWith('[trash]/');

/** The signature machinery itself, which a signature cannot cover, and items that are not parts. */
export const isOutsideCoverage = (name: string): boolean =>
	isSignaturePart(name) || isNotAPart(partNameKey(name));

export interface SignatureCoverage {
	uncoveredParts: string[];
	coverageGaps: SignatureCoverageGap[];
}

/** The part whose relationships `relationshipsPart` holds; empty for the package itself. */
function sourcePartOf(relationshipsPart: string): string {
	const match = /^(.*?)_rels\/([^/]*)\.rels$/i.exec(relationshipsPart);
	return match ? `${match[1]}${match[2]}` : '';
}

/**
 * The part name a relationship targets. Percent-encoding is part of the name, and a reference with
 * an empty path, which the schema allows, names its source (RFC 3986, 5.2.2).
 */
function targetPartOf(sourcePart: string, relationship: PackageRelationship): string {
	const path = relationship.target.split(/[?#]/)[0];
	return path ? resolvePartPath(sourcePart, path) : sourcePart;
}

function mayStayUnsigned(sourcePart: string, relationship: PackageRelationship): boolean {
	if (relationship.mode !== 'Internal' || !UNSIGNED_RELATIONSHIP_TYPES.has(relationship.type)) {
		return false;
	}
	const target = targetPartOf(sourcePart, relationship);
	return isSignaturePart(target) || UNSIGNED_PART_LOCATION.test(partNameKey(target));
}

/** The relationships in a part, or `undefined` when it is not a valid relationships part. */
async function readRelationships(
	pkg: SignaturePackage,
	relationshipsPart: string,
): Promise<PackageRelationship[] | undefined> {
	try {
		const xml = await pkg.zip.file(relationshipsPart)?.async('string');
		return xml === undefined ? undefined : parseRelationshipsPart(xml);
	} catch {
		return undefined;
	}
}

/** Report what the verified references leave uncovered. */
export async function evaluateSignatureCoverage(
	pkg: SignaturePackage,
	coveredParts: ReadonlyMap<string, CoveredPart>,
): Promise<SignatureCoverage> {
	const coverageGaps: SignatureCoverageGap[] = [];
	const pending = [resolvePartName(pkg, PACKAGE_RELATIONSHIPS_PART) ?? PACKAGE_RELATIONSHIPS_PART];
	const visited = new Set<string>();
	let reachedSignedContent = false;

	while (pending.length > 0) {
		const relationshipsPart = pending.pop()!;
		if (visited.has(relationshipsPart)) continue;
		visited.add(relationshipsPart);

		const covered = coveredParts.get(relationshipsPart);
		const all = covered && (await readRelationships(pkg, relationshipsPart));
		if (!covered || !all) {
			coverageGaps.push({ kind: 'unsigned-relationships-part', part: relationshipsPart });
			continue;
		}
		const sourcePart = sourcePartOf(relationshipsPart);
		const selected = covered.selectedRelationships ?? all;
		const selectedIds = new Set(selected.map((relationship) => relationship.id));
		for (const relationship of all) {
			if (!selectedIds.has(relationship.id) && !mayStayUnsigned(sourcePart, relationship)) {
				coverageGaps.push({
					kind: 'unsigned-relationship',
					part: relationshipsPart,
					relationshipId: relationship.id,
				});
			}
		}

		for (const relationship of selected) {
			if (relationship.mode === 'External') continue;
			const target = resolvePartName(pkg, targetPartOf(sourcePart, relationship));
			if (target === undefined || isOutsideCoverage(target)) continue;
			if (!coveredParts.has(target)) {
				if (!mayStayUnsigned(sourcePart, relationship)) {
					coverageGaps.push({
						kind: 'unsigned-part',
						part: target,
						relationshipId: relationship.id,
					});
				}
				continue;
			}
			reachedSignedContent = true;
			const targetRelationships = resolvePartName(pkg, relationshipsPartFor(target));
			if (targetRelationships !== undefined) pending.push(targetRelationships);
		}
	}

	if (!reachedSignedContent) {
		coverageGaps.push({ kind: 'no-signed-content', part: PACKAGE_RELATIONSHIPS_PART });
	}
	const uncoveredParts = [...pkg.partNames]
		.filter((part) => !coveredParts.has(part) && !isOutsideCoverage(part))
		.sort();
	const reported = new Set(coverageGaps.map((gap) => gap.part));
	for (const part of uncoveredParts) {
		if (!reported.has(part) && !UNSIGNED_PART_LOCATION.test(partNameKey(part))) {
			coverageGaps.push({ kind: 'unsigned-part', part });
		}
	}
	return { uncoveredParts, coverageGaps };
}
