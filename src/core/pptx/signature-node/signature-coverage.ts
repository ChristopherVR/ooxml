/**
 * Decides whether verified references cover a package's content.
 *
 * OPC signatures may sign a subset of a package, and Office leaves document properties, thumbnails
 * and a few settings parts out. So coverage is a walk: from the package relationships, follow only
 * relationships the signature protects, and require every part reached, and every relationships
 * part beside it, to be covered. A relationship added beside the signed ones is a gap unless its
 * type is one that signers routinely leave out.
 *
 * Node-only - reads parts through `jszip`.
 */

import { relationshipsPartFor, resolvePartPath } from '../../opc/relationships';
import type { PackageRelationship, SignatureCoverageGap } from '../core/utils/signature-types';
import type { CoveredPart, SignaturePackage } from './reference-checks';
import { parseRelationshipsPart } from './reference-transforms';

const PACKAGE_RELATIONSHIPS_PART = '_rels/.rels';
const CONTENT_TYPES_PART = '[Content_Types].xml';
const SIGNATURE_ORIGIN_PART = '_xmlsignatures/origin.sigs';

/** Relationship types Office leaves unsigned: metadata and viewing state, not document content. */
const UNSIGNED_RELATIONSHIP_TYPE_SUFFIXES = [
	'/metadata/core-properties',
	'/metadata/thumbnail',
	'/extended-properties',
	'/custom-properties',
	'/digital-signature/origin',
	'/customXml',
	'/presProps',
	'/viewProps',
];

export interface SignatureCoverage {
	uncoveredParts: string[];
	coverageGaps: SignatureCoverageGap[];
}

const mayStayUnsigned = (relationship: PackageRelationship): boolean =>
	UNSIGNED_RELATIONSHIP_TYPE_SUFFIXES.some((suffix) => relationship.type.endsWith(suffix));

/** The part whose relationships `relationshipsPart` holds; empty for the package itself. */
function sourcePartOf(relationshipsPart: string): string {
	const match = /^(.*?)_rels\/([^/]*)\.rels$/.exec(relationshipsPart);
	return match ? `${match[1]}${match[2]}` : '';
}

function targetPartOf(sourcePart: string, relationship: PackageRelationship): string {
	const target = relationship.target.split(/[?#]/)[0] ?? '';
	let decoded = target;
	try {
		decoded = decodeURIComponent(target);
	} catch {
		// keep the raw target: it then matches no part
	}
	return resolvePartPath(sourcePart, decoded);
}

/** The signature machinery itself, which a signature cannot cover. */
function signatureInfrastructure(signaturePaths: readonly string[]): Set<string> {
	const parts = new Set([
		CONTENT_TYPES_PART,
		SIGNATURE_ORIGIN_PART,
		relationshipsPartFor(SIGNATURE_ORIGIN_PART),
	]);
	for (const path of signaturePaths) {
		parts.add(path);
		parts.add(relationshipsPartFor(path));
	}
	return parts;
}

/** The relationships in a part, or `undefined` when it is not a readable relationships part. */
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
	signaturePaths: readonly string[],
): Promise<SignatureCoverage> {
	const infrastructure = signatureInfrastructure(signaturePaths);
	const coverageGaps: SignatureCoverageGap[] = [];
	const pending = [PACKAGE_RELATIONSHIPS_PART];
	const visited = new Set<string>();
	let reachedSignedContent = false;

	while (pending.length > 0) {
		const relationshipsPart = pending.pop()!;
		if (visited.has(relationshipsPart)) continue;
		visited.add(relationshipsPart);

		const covered = coveredParts.get(relationshipsPart);
		if (!covered) {
			coverageGaps.push({ kind: 'unsigned-relationships-part', part: relationshipsPart });
			continue;
		}
		const all = await readRelationships(pkg, relationshipsPart);
		if (!all) {
			coverageGaps.push({ kind: 'unsigned-relationships-part', part: relationshipsPart });
			continue;
		}
		const selected = covered.selectedRelationships ?? all;
		const selectedSet = new Set(selected.map((relationship) => relationship.id));
		for (const relationship of all) {
			if (!selectedSet.has(relationship.id) && !mayStayUnsigned(relationship)) {
				coverageGaps.push({
					kind: 'unsigned-relationship',
					part: relationshipsPart,
					relationshipId: relationship.id,
				});
			}
		}

		const sourcePart = sourcePartOf(relationshipsPart);
		for (const relationship of selected) {
			if (relationship.mode === 'External') continue;
			const target = targetPartOf(sourcePart, relationship);
			if (infrastructure.has(target) || !pkg.partNames.has(target)) continue;
			if (!coveredParts.has(target)) {
				coverageGaps.push({
					kind: 'unsigned-part',
					part: target,
					relationshipId: relationship.id,
				});
				continue;
			}
			reachedSignedContent = true;
			const targetRelationships = relationshipsPartFor(target);
			if (pkg.partNames.has(targetRelationships)) pending.push(targetRelationships);
		}
	}

	if (!reachedSignedContent) {
		coverageGaps.push({ kind: 'no-signed-content', part: PACKAGE_RELATIONSHIPS_PART });
	}
	const uncoveredParts = [...pkg.partNames]
		.filter((part) => !coveredParts.has(part) && !infrastructure.has(part))
		.sort();
	return { uncoveredParts, coverageGaps };
}
