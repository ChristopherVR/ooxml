import type JSZip from 'jszip';
import { buildXml, elements, parseXml, type XmlDocument } from '../../xml/index';
import { contentTypeForPart, parseContentTypes } from '../content-types';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../relationships';
import { DIGITAL_SIGNATURE_ORIGIN_REL_TYPE, DIGITAL_SIGNATURE_REL_TYPE } from './constants';
import { isSignaturePart } from './detection';

const signatureTypes = new Set([
	'application/vnd.openxmlformats-package.digital-signature-origin',
	'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml',
	'application/vnd.openxmlformats-package.digital-signature-certificate',
]);
const signatureRelationships = new Set([
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate',
]);

function partPath(base: string, target: string): string {
	return resolvePartPath(base, decodeURIComponent(target));
}

/** Signature metadata only: no cryptographic verification or certificate trust decision. */
export interface PackageSignatures {
	hasSignatures: boolean;
	/** Origin, signature, certificate and relationship parts to remove after an edit. */
	parts: Set<string>;
}

/** Inspect an already bounded package, including signatures outside `_xmlsignatures/`. */
export async function inspectPackageSignatures(zip: JSZip): Promise<PackageSignatures> {
	const parts = new Set<string>();
	const types = parseContentTypes(await zip.file('[Content_Types].xml')?.async('string'));
	let hasSignatures = [...types.defaults.values(), ...types.overrides.values()].some((type) =>
		signatureTypes.has(type),
	);
	for (const path of Object.keys(zip.files)) {
		if (isSignaturePart(path) || signatureTypes.has(contentTypeForPart(types, path) ?? ''))
			parts.add(path);
	}
	// Overrides can also identify missing parts in an incomplete signed package.
	for (const [path, type] of types.overrides)
		if (signatureTypes.has(type)) parts.add(partPath('', path));
	const rootRels = parseRelationships(await zip.file('_rels/.rels')?.async('string'));
	for (const rel of rootRels.values()) {
		if (!signatureRelationships.has(rel.type)) continue;
		hasSignatures = true;
		if (rel.mode === 'Internal') parts.add(partPath('', rel.target));
	}
	// Set iteration visits newly discovered parts too, and terminates on cyclic relationships.
	for (const path of parts) {
		if (path.endsWith('/') || path.endsWith('.rels')) continue;
		const relsPath = relationshipsPartFor(path);
		const file = zip.file(relsPath);
		if (!file) continue;
		parts.add(relsPath);
		for (const rel of parseRelationships(await file.async('string')).values())
			if (rel.mode === 'Internal' && signatureRelationships.has(rel.type))
				parts.add(partPath(path, rel.target));
	}
	return { hasSignatures: hasSignatures || parts.size > 0, parts };
}

/** Remove invalidated package signatures, preserving unrelated XML and package parts. */
export async function stripPackageSignatures(zip: JSZip): Promise<void> {
	const signatures = await inspectPackageSignatures(zip);
	if (!signatures.hasSignatures) return;
	const documents = new Map<string, XmlDocument>();
	for (const path of ['_rels/.rels', '[Content_Types].xml']) {
		const file = zip.file(path);
		if (!file) continue;
		const document = parseXml(await file.async('string'));
		let changed = false;
		for (const node of elements(document.documentElement)) {
			let remove = false;
			if (path === '_rels/.rels' && node.localName === 'Relationship') {
				const target = node.getAttribute('Target');
				remove =
					signatureRelationships.has(node.getAttribute('Type') ?? '') ||
					(node.getAttribute('TargetMode') !== 'External' &&
						target !== null &&
						signatures.parts.has(partPath('', target)));
			} else if (node.localName === 'Default' || node.localName === 'Override') {
				const part = node.getAttribute('PartName');
				remove =
					signatureTypes.has(node.getAttribute('ContentType') ?? '') ||
					(part !== null && (isSignaturePart(part) || signatures.parts.has(partPath('', part))));
			}
			if (remove) {
				document.documentElement.removeChild(node);
				changed = true;
			}
		}
		if (changed) documents.set(path, document);
	}
	for (const path of signatures.parts) zip.remove(path);
	for (const [path, document] of documents) zip.file(path, buildXml(document));
}
