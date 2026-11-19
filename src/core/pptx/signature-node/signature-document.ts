/**
 * Strict reading of an XML-DSig `Signature` document.
 *
 * Every field the verifier relies on comes from one exact place in the XML-DSig schema. A signature
 * with repeated, misplaced or ambiguous structure is rejected here, so a checked element can never
 * differ from the element that was signed.
 *
 * Node-only - depends on `@xmldom/xmldom` through `xml-canonicalization`.
 */

import { SUPPORTED_XML_CANON_TRANSFORMS, XMLDSIG_NS } from '../core/utils/signature-constants';
import { childElements, getNodeLocalName, parseStrictXml } from './xml-canonicalization';
import type { XmlElement } from './xml-canonicalization';

/** The parts of a structurally valid signature. Nothing here is authenticated yet. */
export interface XmlSignature {
	signedInfo: XmlElement;
	canonicalizationMethod: string;
	signatureMethod: string;
	signatureValueBase64: string;
	/** Base64 DER certificates from `KeyInfo/X509Data`, in document order. */
	certificates: string[];
	/** Elements a `#id` reference may name: each `Object` and everything inside it. */
	referenceTargets: ReadonlyMap<string, XmlElement>;
}

export type ParsedXmlSignature =
	| { ok: true; signature: XmlSignature }
	| { ok: false; error: string };

class SignatureStructureError extends Error {}

const ID_ATTRIBUTES = ['Id', 'ID', 'id'];

const isDsigElement = (element: XmlElement, localName: string): boolean =>
	element.namespaceURI === XMLDSIG_NS && getNodeLocalName(element) === localName;

const stripWhitespace = (value: string | null): string => (value ?? '').replace(/\s+/g, '');

function requireAlgorithm(element: XmlElement): string {
	const algorithm = element.getAttribute('Algorithm')?.trim();
	if (!algorithm) {
		throw new SignatureStructureError(`${getNodeLocalName(element)} has no Algorithm.`);
	}
	if (childElements(element).length > 0) {
		throw new SignatureStructureError(`${getNodeLocalName(element)} parameters are unsupported.`);
	}
	return algorithm;
}

function descendantsAndSelf(element: XmlElement): XmlElement[] {
	return [element, ...Array.from(element.getElementsByTagName('*'))];
}

function elementId(element: XmlElement): string | undefined {
	for (const name of ID_ATTRIBUTES) {
		const value = element.getAttribute(name);
		if (value) return value;
	}
	return undefined;
}

/** Rejects a repeated `Id` anywhere in the signature: a `#id` must name exactly one element. */
function assertUniqueIds(root: XmlElement): void {
	const seen = new Set<string>();
	for (const element of descendantsAndSelf(root)) {
		const id = elementId(element);
		if (id === undefined) continue;
		if (seen.has(id)) {
			throw new SignatureStructureError(`Duplicate Id "${id}".`);
		}
		seen.add(id);
	}
}

function readSignedInfo(signedInfo: XmlElement) {
	const [canonicalization, method, ...references] = childElements(signedInfo);
	if (
		!canonicalization ||
		!method ||
		!isDsigElement(canonicalization, 'CanonicalizationMethod') ||
		!isDsigElement(method, 'SignatureMethod')
	) {
		throw new SignatureStructureError('SignedInfo must start with its two method elements.');
	}
	if (references.length === 0 || !references.every((el) => isDsigElement(el, 'Reference'))) {
		throw new SignatureStructureError('SignedInfo must hold one or more References only.');
	}
	const canonicalizationMethod = requireAlgorithm(canonicalization);
	if (!SUPPORTED_XML_CANON_TRANSFORMS.has(canonicalizationMethod)) {
		throw new SignatureStructureError('Unsupported SignedInfo canonicalization method.');
	}
	return { canonicalizationMethod, signatureMethod: requireAlgorithm(method) };
}

function readCertificates(keyInfo: XmlElement | undefined): string[] {
	if (!keyInfo) return [];
	return childElements(keyInfo)
		.filter((element) => isDsigElement(element, 'X509Data'))
		.flatMap((data) => childElements(data))
		.filter((element) => isDsigElement(element, 'X509Certificate'))
		.map((element) => stripWhitespace(element.textContent))
		.filter((value) => value.length > 0);
}

function readSignature(xml: string): XmlSignature {
	const root = parseStrictXml(xml).documentElement!;
	if (!isDsigElement(root, 'Signature')) {
		throw new SignatureStructureError('The document element is not an XML-DSig Signature.');
	}
	assertUniqueIds(root);

	const children = childElements(root);
	const [signedInfo, signatureValue] = children;
	if (
		!signedInfo ||
		!signatureValue ||
		!isDsigElement(signedInfo, 'SignedInfo') ||
		!isDsigElement(signatureValue, 'SignatureValue')
	) {
		throw new SignatureStructureError('Signature must start with SignedInfo and SignatureValue.');
	}
	let rest = children.slice(2);
	const keyInfo = rest[0] && isDsigElement(rest[0], 'KeyInfo') ? rest[0] : undefined;
	if (keyInfo) rest = rest.slice(1);
	if (!rest.every((element) => isDsigElement(element, 'Object'))) {
		throw new SignatureStructureError('Signature holds an unexpected or repeated element.');
	}

	const referenceTargets = new Map<string, XmlElement>();
	for (const element of rest.flatMap(descendantsAndSelf)) {
		const id = elementId(element);
		if (id !== undefined) referenceTargets.set(id, element);
	}

	return {
		signedInfo,
		...readSignedInfo(signedInfo),
		signatureValueBase64: stripWhitespace(signatureValue.textContent),
		certificates: readCertificates(keyInfo),
		referenceTargets,
	};
}

/** Parse a signature part, reporting why it is rejected instead of throwing. */
export function parseXmlSignature(xml: string): ParsedXmlSignature {
	try {
		return { ok: true, signature: readSignature(xml) };
	} catch (error) {
		return { ok: false, error: error instanceof Error ? error.message : String(error) };
	}
}
