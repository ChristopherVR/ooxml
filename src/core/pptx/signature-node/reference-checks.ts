/**
 * Reference digest verification for digital signatures (XML-DSig, ECMA-376 Part 2).
 *
 * References are read only from text whose integrity is already established: the canonical
 * `SignedInfo` a verified `SignatureValue` covers, then the canonical `Object` a verified reference
 * covers. Anything else in the signature part is ignored.
 *
 * Node-only - uses `node:crypto` for hashing and `jszip` for package access.
 */

import crypto from 'node:crypto';

import type JSZip from 'jszip';

import { contentTypeForPart, parseContentTypes } from '../../opc/content-types';
import { MEDIA_TYPES_ITEM, partNameKey } from '../../opc/part-names';
import type { ContentTypes } from '../../opc/content-types';
import {
	DIGEST_ALGORITHM_TO_HASH,
	OPC_DIGITAL_SIGNATURE_NS,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
} from '../core/utils/signature-constants';
import { parseReferenceUri } from '../core/utils/signature-reference-utils';
import type {
	PackageRelationship,
	ParsedReferenceTransform,
	SignatureReferenceCheck,
} from '../core/utils/signature-types';
import { applyReferenceTransforms, extractReferenceTransforms } from './reference-transforms';
import type { XmlSignature } from './signature-document';
import {
	canonicalizeNode,
	childElements,
	getFirstDescendantElementByLocalName,
	getNodeLocalName,
	INCLUSIVE_C14N,
	parseStrictXml,
} from './xml-canonicalization';
import type { XmlElement } from './xml-canonicalization';

/** The package as the verifier reads it. */
export interface SignaturePackage {
	zip: JSZip;
	partNames: ReadonlySet<string>;
	/** Stored names by {@link partNameKey}, for names that differ from the stored spelling. */
	namesByKey: ReadonlyMap<string, string>;
	contentTypes: ContentTypes;
}

/** The stored name of the part `name` denotes: that exact item, else an equivalent one. */
export function resolvePartName(pkg: SignaturePackage, name: string): string | undefined {
	return pkg.partNames.has(name) ? name : pkg.namesByKey.get(partNameKey(name));
}

/** A part a verified reference covers, keyed by its stored name. */
export interface CoveredPart {
	/** Set when a relationships transform limited the reference to these relationships. */
	selectedRelationships?: PackageRelationship[] | undefined;
}

export interface SignedReferenceResult {
	checks: SignatureReferenceCheck[];
	coveredParts: Map<string, CoveredPart>;
	/** Signing time, read from signed content only. */
	signingTime?: string | undefined;
}

interface SignedReference {
	uri: string;
	digestAlgorithm?: string | undefined;
	digestExpectedBase64?: string | undefined;
	transforms: ParsedReferenceTransform[];
}

/** Compute a Base64-encoded digest using Node.js `crypto` (synchronous). */
export function computeDigestBase64(
	content: Uint8Array,
	digestAlgorithmUri: string,
): string | undefined {
	const hashName = DIGEST_ALGORITHM_TO_HASH[digestAlgorithmUri];
	if (!hashName) {
		return undefined;
	}
	return crypto.createHash(hashName).update(Buffer.from(content)).digest('base64');
}

/** Read what reference checks need from the package once per inspection. */
export async function openSignaturePackage(zip: JSZip): Promise<SignaturePackage> {
	const partNames = new Set(Object.keys(zip.files).filter((name) => !zip.files[name]?.dir));
	const namesByKey = new Map([...partNames].map((name) => [partNameKey(name), name]));
	const pkg: SignaturePackage = {
		zip,
		partNames,
		namesByKey,
		contentTypes: parseContentTypes(undefined),
	};
	const mediaTypes = resolvePartName(pkg, MEDIA_TYPES_ITEM);
	try {
		pkg.contentTypes = parseContentTypes(
			mediaTypes === undefined ? undefined : await zip.file(mediaTypes)?.async('string'),
		);
	} catch {
		// An unreadable content types part declares nothing, so `?ContentType` references fail.
	}
	return pkg;
}

const isDsigElement = (element: XmlElement, localName: string): boolean =>
	element.namespaceURI === XMLDSIG_NS && getNodeLocalName(element) === localName;

function readReference(element: XmlElement): SignedReference {
	const children = childElements(element);
	const onlyChild = (localName: string): XmlElement | undefined => {
		const matches = children.filter((child) => isDsigElement(child, localName));
		return matches.length === 1 ? matches[0] : undefined;
	};
	return {
		uri: element.getAttribute('URI')?.trim() ?? '',
		digestAlgorithm: onlyChild('DigestMethod')?.getAttribute('Algorithm')?.trim(),
		digestExpectedBase64: onlyChild('DigestValue')?.textContent?.replace(/\s+/g, ''),
		transforms: extractReferenceTransforms(element),
	};
}

const readReferences = (parent: XmlElement): SignedReference[] =>
	childElements(parent)
		.filter((child) => isDsigElement(child, 'Reference'))
		.map(readReference);

const digestsEqual = (left: string, right: string): boolean =>
	Buffer.from(left, 'base64').equals(Buffer.from(right, 'base64'));

/** The reference's fields every check reports, with the given outcome. */
function referenceCheck(
	reference: SignedReference,
	digestStatus: SignatureReferenceCheck['digestStatus'],
	extra: Partial<SignatureReferenceCheck> = {},
): SignatureReferenceCheck {
	return {
		uri: reference.uri,
		existsInPackage: true,
		digestAlgorithm: reference.digestAlgorithm,
		digestExpectedBase64: reference.digestExpectedBase64,
		digestStatus,
		transformAlgorithms: reference.transforms.map((transform) => transform.algorithm),
		...extra,
	};
}

function compareDigest(
	reference: SignedReference,
	octets: Uint8Array,
	extra: Partial<SignatureReferenceCheck> = {},
): SignatureReferenceCheck {
	const { digestAlgorithm, digestExpectedBase64 } = reference;
	if (!digestAlgorithm || !digestExpectedBase64) {
		return referenceCheck(reference, 'insufficient-data', extra);
	}
	const digestActualBase64 = computeDigestBase64(octets, digestAlgorithm);
	if (!digestActualBase64) {
		return referenceCheck(reference, 'unsupported-algorithm', extra);
	}
	const matches = digestsEqual(digestActualBase64, digestExpectedBase64);
	return referenceCheck(reference, matches ? 'verified' : 'mismatch', {
		...extra,
		digestActualBase64,
	});
}

async function checkPartReference(
	reference: SignedReference,
	pkg: SignaturePackage,
	coveredParts: Map<string, CoveredPart>,
): Promise<SignatureReferenceCheck> {
	const parsed = parseReferenceUri(reference.uri);
	if (!parsed) {
		return referenceCheck(reference, 'insufficient-data');
	}
	const { partPath, contentType } = parsed;
	const resolvedPartPath = resolvePartName(pkg, partPath);
	const part = resolvedPartPath === undefined ? null : pkg.zip.file(resolvedPartPath);
	if (!part || resolvedPartPath === undefined) {
		return referenceCheck(reference, 'missing-part', {
			resolvedPartPath: partPath,
			existsInPackage: false,
		});
	}
	if (contentType !== undefined) {
		const declared = contentTypeForPart(pkg.contentTypes, resolvedPartPath);
		if (declared?.toLowerCase() !== contentType.toLowerCase()) {
			return referenceCheck(reference, 'content-type-mismatch', { resolvedPartPath });
		}
	}
	const transformed = applyReferenceTransforms(
		await part.async('uint8array'),
		reference.transforms,
	);
	if (transformed.unsupportedAlgorithms.length > 0) {
		return referenceCheck(reference, 'unsupported-transform', { resolvedPartPath });
	}
	const check = compareDigest(reference, transformed.data, { resolvedPartPath });
	if (check.digestStatus === 'verified') {
		coveredParts.set(resolvedPartPath, {
			selectedRelationships: transformed.selectedRelationships,
		});
	}
	return check;
}

/** The canonical octets of a same-document reference, or `undefined` for an unsupported chain. */
function canonicalObjectOctets(
	element: XmlElement,
	transforms: ParsedReferenceTransform[],
): string | undefined {
	const algorithms = transforms.map((transform) => transform.algorithm);
	if (!algorithms.every((algorithm) => SUPPORTED_XML_CANON_TRANSFORMS.has(algorithm))) {
		return undefined;
	}
	let octets = canonicalizeNode(element, algorithms[0] ?? INCLUSIVE_C14N);
	for (const algorithm of algorithms.slice(1)) {
		octets = canonicalizeNode(parseStrictXml(octets), algorithm);
	}
	return octets;
}

function readSigningTime(signedElement: XmlElement): string | undefined {
	const packageTime = Array.from(
		signedElement.getElementsByTagNameNS(OPC_DIGITAL_SIGNATURE_NS, 'SignatureTime'),
	)[0];
	const value = packageTime
		? getFirstDescendantElementByLocalName(packageTime, 'Value')
		: getNodeLocalName(signedElement) === 'SigningTime'
			? signedElement
			: getFirstDescendantElementByLocalName(signedElement, 'SigningTime');
	return value?.textContent?.trim() || undefined;
}

/**
 * Verify every reference of a `SignedInfo` whose signature value has verified, following each
 * Manifest inside an `Object` that a verified reference protects.
 */
export async function checkSignedReferences(
	signature: XmlSignature,
	canonicalSignedInfo: string,
	pkg: SignaturePackage,
): Promise<SignedReferenceResult> {
	const checks: SignatureReferenceCheck[] = [];
	const coveredParts = new Map<string, CoveredPart>();
	let signingTime: string | undefined;

	for (const reference of readReferences(parseStrictXml(canonicalSignedInfo).documentElement!)) {
		if (!reference.uri.startsWith('#')) {
			checks.push(await checkPartReference(reference, pkg, coveredParts));
			continue;
		}
		const target = signature.referenceTargets.get(reference.uri.slice(1));
		if (!target) {
			checks.push(referenceCheck(reference, 'missing-part', { existsInPackage: false }));
			continue;
		}
		const octets = canonicalObjectOctets(target, reference.transforms);
		if (octets === undefined) {
			checks.push(referenceCheck(reference, 'unsupported-transform'));
			continue;
		}
		const check = compareDigest(reference, new Uint8Array(Buffer.from(octets, 'utf8')));
		checks.push(check);
		if (check.digestStatus !== 'verified') {
			continue;
		}
		const signedElement = parseStrictXml(octets).documentElement!;
		signingTime ??= readSigningTime(signedElement);
		const manifests = isDsigElement(signedElement, 'Manifest')
			? [signedElement]
			: childElements(signedElement).filter((child) => isDsigElement(child, 'Manifest'));
		for (const manifestReference of manifests.flatMap(readReferences)) {
			checks.push(await checkPartReference(manifestReference, pkg, coveredParts));
		}
	}
	return { checks, coveredParts, signingTime };
}
