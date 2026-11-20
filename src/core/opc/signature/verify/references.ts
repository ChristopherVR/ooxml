/**
 * Reference digest verification for digital signatures (XML-DSig, ECMA-376 Part 2).
 *
 * References are read only from text whose integrity is already established: the canonical
 * `SignedInfo` a verified `SignatureValue` covers, then the canonical `Object` a verified reference
 * covers. Anything else in the signature part is ignored. Platform-neutral.
 */

import type JSZip from 'jszip';

import { decodeBase64, digestSync, encodeBase64 } from '../../../digest/index';
import { decodeXmlBytes } from '../../../xml/index';
import { contentTypeForPart, parseContentTypes } from '../../content-types';
import type { ContentTypes } from '../../content-types';
import { MEDIA_TYPES_ITEM, partNameKey } from '../../part-names';
import {
	DIGEST_ALGORITHM_TO_WEB_CRYPTO,
	INCLUSIVE_C14N,
	OPC_DIGITAL_SIGNATURE_NS,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
} from '../constants';
import { parseReferenceUri } from '../reference-utils';
import type {
	PackageRelationship,
	ParsedReferenceTransform,
	SignatureReferenceCheck,
} from '../types';
import { canonicalize } from './c14n';
import type { XmlSignature } from './document';
import { childElements, firstDescendantNamed, localNameOf, parseStrictXml } from './dom';
import { SIGNATURE_LIMITS } from './limits';
import { applyReferenceTransforms, extractReferenceTransforms } from './transforms';

/** Where the verifier reads a package's parts: a JSZip archive or a map of part bytes. */
export type SignaturePackageSource = JSZip | ReadonlyMap<string, Uint8Array>;

/** The package as the verifier reads it. */
export interface SignaturePackage {
	partNames: ReadonlySet<string>;
	/** Stored names by {@link partNameKey}, for names that differ from the stored spelling. */
	namesByKey: ReadonlyMap<string, string>;
	contentTypes: ContentTypes;
	/** The bytes of a stored part, by its stored name. */
	read(name: string): Promise<Uint8Array | undefined>;
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

/** The Base64 digest of `content` under an XML-DSig digest algorithm URI, when it is one known. */
export function computeDigestBase64(
	content: Uint8Array,
	digestAlgorithmUri: string,
): string | undefined {
	const hash = DIGEST_ALGORITHM_TO_WEB_CRYPTO[digestAlgorithmUri];
	const digest = hash === undefined ? undefined : digestSync(hash, content);
	return digest && encodeBase64(digest);
}

const isZip = (source: SignaturePackageSource): source is JSZip =>
	typeof (source as JSZip).file === 'function';

/** Read what reference checks need from the package once per inspection. */
export async function openSignaturePackage(
	source: SignaturePackageSource,
): Promise<SignaturePackage> {
	const zip = isZip(source) ? source : undefined;
	const partNames = new Set(
		zip
			? Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
			: [...(source as ReadonlyMap<string, Uint8Array>).keys()],
	);
	const read = async (name: string): Promise<Uint8Array | undefined> =>
		zip
			? await zip.file(name)?.async('uint8array')
			: (source as ReadonlyMap<string, Uint8Array>).get(name);
	const namesByKey = new Map([...partNames].map((name) => [partNameKey(name), name]));
	const pkg: SignaturePackage = {
		partNames,
		namesByKey,
		contentTypes: parseContentTypes(undefined),
		read: (name) => (partNames.has(name) ? read(name) : Promise.resolve(undefined)),
	};
	const mediaTypes = resolvePartName(pkg, MEDIA_TYPES_ITEM);
	try {
		const bytes = mediaTypes === undefined ? undefined : await pkg.read(mediaTypes);
		pkg.contentTypes = parseContentTypes(bytes && decodeXmlBytes(bytes));
	} catch {
		// An unreadable content types part declares nothing, so `?ContentType` references fail.
	}
	return pkg;
}

const isDsigElement = (element: Element, localName: string): boolean =>
	element.namespaceURI === XMLDSIG_NS && localNameOf(element) === localName;

function readReference(element: Element): SignedReference {
	const children = childElements(element);
	const onlyChild = (localName: string): Element | undefined => {
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

const readReferences = (parent: Element): SignedReference[] =>
	childElements(parent)
		.filter((child) => isDsigElement(child, 'Reference'))
		.map(readReference);

function digestsEqual(left: string, right: string): boolean {
	const [a, b] = [decodeBase64(left), decodeBase64(right)];
	return !!a && !!b && a.length === b.length && a.every((byte, index) => byte === b[index]);
}

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

/** What one signature may still read; see {@link SIGNATURE_LIMITS}. */
interface WorkBudget {
	references: number;
	bytes: number;
}

async function checkPartReference(
	reference: SignedReference,
	pkg: SignaturePackage,
	coveredParts: Map<string, CoveredPart>,
	budget: WorkBudget,
): Promise<SignatureReferenceCheck> {
	const parsed = parseReferenceUri(reference.uri);
	if (!parsed) {
		return referenceCheck(reference, 'insufficient-data');
	}
	const { partPath, contentType } = parsed;
	const resolvedPartPath = resolvePartName(pkg, partPath);
	const bytes = resolvedPartPath === undefined ? undefined : await pkg.read(resolvedPartPath);
	budget.bytes -= bytes?.length ?? 0;
	if (!bytes || resolvedPartPath === undefined) {
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
	const transformed = applyReferenceTransforms(bytes, reference.transforms);
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
	element: Element,
	transforms: ParsedReferenceTransform[],
): string | undefined {
	const algorithms = transforms.map((transform) => transform.algorithm);
	if (
		transforms.some((transform) => transform.unsupportedParameters) ||
		!algorithms.every((algorithm) => SUPPORTED_XML_CANON_TRANSFORMS.has(algorithm))
	) {
		return undefined;
	}
	let octets = canonicalize(element, algorithms[0] ?? INCLUSIVE_C14N);
	for (const algorithm of algorithms.slice(1)) {
		octets = canonicalize(parseStrictXml(octets), algorithm);
	}
	return octets;
}

function readSigningTime(signedElement: Element): string | undefined {
	const packageTime = Array.from(
		signedElement.getElementsByTagNameNS(OPC_DIGITAL_SIGNATURE_NS, 'SignatureTime'),
	)[0];
	const value = packageTime
		? firstDescendantNamed(packageTime, 'Value')
		: localNameOf(signedElement) === 'SigningTime'
			? signedElement
			: firstDescendantNamed(signedElement, 'SigningTime');
	return value?.textContent?.trim() || undefined;
}

/**
 * Verify every reference of a `SignedInfo` whose signature value has verified, following each
 * Manifest inside an `Object` that a verified reference protects. Once the signature has read
 * `SIGNATURE_LIMITS.maxReferences` references or `maxDigestedBytes` bytes, the next reference is
 * reported as `limit-exceeded` and no more are read, so the signature is not verified.
 */
export async function checkSignedReferences(
	signature: XmlSignature,
	canonicalSignedInfo: string,
	pkg: SignaturePackage,
): Promise<SignedReferenceResult> {
	const checks: SignatureReferenceCheck[] = [];
	const coveredParts = new Map<string, CoveredPart>();
	let signingTime: string | undefined;
	const budget: WorkBudget = {
		references: SIGNATURE_LIMITS.maxReferences,
		bytes: SIGNATURE_LIMITS.maxDigestedBytes,
	};
	/** Whether `reference` may be read; records it as `limit-exceeded` when not. */
	const admit = (reference: SignedReference): boolean => {
		budget.references -= 1;
		if (budget.references >= 0 && budget.bytes >= 0) return true;
		checks.push(referenceCheck(reference, 'limit-exceeded'));
		return false;
	};

	for (const reference of readReferences(parseStrictXml(canonicalSignedInfo).documentElement!)) {
		if (!admit(reference)) break;
		if (!reference.uri.startsWith('#')) {
			checks.push(await checkPartReference(reference, pkg, coveredParts, budget));
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
		const encoded = new TextEncoder().encode(octets);
		budget.bytes -= encoded.length;
		const check = compareDigest(reference, encoded);
		checks.push(check);
		if (check.digestStatus !== 'verified') {
			continue;
		}
		const signedElement = parseStrictXml(octets).documentElement!;
		signingTime ??= readSigningTime(signedElement);
		const manifests = isDsigElement(signedElement, 'Manifest')
			? [signedElement]
			: childElements(signedElement).filter((child) => isDsigElement(child, 'Manifest'));
		let withinBudget = true;
		for (const manifestReference of manifests.flatMap(readReferences)) {
			withinBudget = admit(manifestReference);
			if (!withinBudget) break;
			checks.push(await checkPartReference(manifestReference, pkg, coveredParts, budget));
		}
		if (!withinBudget) break;
	}
	return { checks, coveredParts, signingTime };
}
