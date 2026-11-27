// Signs an OPC package the way Office does (ECMA-376 Part 2 clause 13, W3C XML-DSig, XAdES-BES):
// a `Manifest` of every part (relationship parts through the relationship transform, so a later
// signature's origin relationship does not break this one), the Office object and the XAdES
// signed properties, all protected by one RSA-SHA256 `SignedInfo` that the host's signer signs.
// Signatures already in the package are kept, so several people can sign one document.
// DOM-free and platform-neutral: digests are synchronous, the signature is the host's.
import JSZip from 'jszip';
import { encodeBase64 } from '../../../digest/base64';
import { contentTypeForPart, parseContentTypes } from '../../content-types';
import { isPartName, MEDIA_TYPES_ITEM } from '../../part-names';
import {
	buildRelationshipsXml,
	nextRelationshipId,
	parseRelationshipsPart,
	relationshipsPartFor,
	resolvePartPath,
	type RelationshipInput,
} from '../../relationships';
import { buildXml, decodeXmlBytes, parseXml } from '../../../xml/index';
import {
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
	INCLUSIVE_C14N,
	OPC_DIGITAL_SIGNATURE_NS,
	OPC_RELATIONSHIP_TRANSFORM,
	XMLDSIG_NS,
} from '../constants';
import { escapeXmlAttr } from '../xml-utils';
import { canonicalize } from '../verify/c14n';
import { isOutsideCoverage } from '../verify/coverage';
import { childElements, parseStrictXml } from '../verify/dom';
import { verifySignedPackage } from '../verify/inspect';
import { computeDigestBase64 } from '../verify/references';
import { applyReferenceTransforms } from '../verify/transforms';
import { readCertificate } from '../verify/x509';
import { isVerifiedSignatureState } from '../state';
import type { PackageSigner } from './signer';
import {
	OFFICE_OBJECT_ID,
	PACKAGE_OBJECT_ID,
	SIGNATURE_ID,
	SIGNED_PROPERTIES_ID,
	officeObjectXml,
	qualifyingPropertiesXml,
	type OfficeSignatureDetails,
} from './office-objects';

const SHA256 = 'http://www.w3.org/2001/04/xmlenc#sha256';
const RSA_SHA256 = 'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256';
const OBJECT_TYPE = 'http://www.w3.org/2000/09/xmldsig#Object';
const SIGNED_PROPERTIES_TYPE = 'http://uri.etsi.org/01903#SignedProperties';
const ORIGIN_TYPE = 'application/vnd.openxmlformats-package.digital-signature-origin';
const SIGNATURE_TYPE = 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml';
const PACKAGE_RELS = '_rels/.rels';

export interface SignPackageOptions extends OfficeSignatureDetails {
	/** The signing time; now by default. */
	now?: Date;
}

const sha256 = (bytes: Uint8Array) => computeDigestBase64(bytes, SHA256)!;
const utf8 = (text: string) => new TextEncoder().encode(text);
const storedItems = (zip: JSZip) => Object.keys(zip.files).filter((name) => !zip.files[name]!.dir);
const isRels = (name: string) => /(^|\/)_rels\/[^/]*\.rels$/iu.test(name);

/** The origin part (existing or new) and the free name for this signature's part. */
async function prepareOrigin(zip: JSZip): Promise<{ origin: string; signature: string }> {
	const xml = await zip.file(PACKAGE_RELS)?.async('uint8array');
	const relationships = xml ? parseRelationshipsPart(decodeXmlBytes(xml)) : [];
	const existing = relationships.find((rel) => rel.type === DIGITAL_SIGNATURE_ORIGIN_REL_TYPE);
	const origin = existing ? resolvePartPath('', existing.target) : '_xmlsignatures/origin.sigs';
	if (!existing) {
		const kept = new Map<string, RelationshipInput>(relationships.map((rel) => [rel.id, rel]));
		kept.set(nextRelationshipId(new Set(kept.keys())), {
			type: DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
			target: origin,
			mode: 'Internal',
		});
		zip.file(PACKAGE_RELS, buildRelationshipsXml(kept));
	}
	if (!zip.file(origin)) zip.file(origin, new Uint8Array());
	const folder = origin.slice(0, origin.lastIndexOf('/') + 1);
	const taken = new Set(storedItems(zip).map((name) => name.toLowerCase()));
	let n = 1;
	while (taken.has(`${folder}sig${n}.xml`.toLowerCase())) n++;
	return { origin, signature: `${folder}sig${n}.xml` };
}

async function registerSignature(zip: JSZip, origin: string, signature: string): Promise<void> {
	const relsPath = relationshipsPartFor(origin);
	const xml = await zip.file(relsPath)?.async('uint8array');
	const kept = new Map<string, RelationshipInput>(
		(xml ? parseRelationshipsPart(decodeXmlBytes(xml)) : []).map((rel) => [rel.id, rel]),
	);
	kept.set(nextRelationshipId(new Set(kept.keys())), {
		type: DIGITAL_SIGNATURE_REL_TYPE,
		target: signature.slice(signature.lastIndexOf('/') + 1),
		mode: 'Internal',
	});
	zip.file(relsPath, buildRelationshipsXml(kept));
	const types = await zip.file(MEDIA_TYPES_ITEM)?.async('string');
	if (!types) throw new Error('The package has no content types.');
	const doc = parseXml(types);
	const root = doc.documentElement;
	const ensure = (partName: string, contentType: string) => {
		const found = Array.from(root.getElementsByTagName('Override')).some(
			(node) => node.getAttribute('PartName')?.toLowerCase() === partName.toLowerCase(),
		);
		if (found) return;
		const override = doc.createElementNS(root.namespaceURI, 'Override');
		override.setAttribute('PartName', partName);
		override.setAttribute('ContentType', contentType);
		root.appendChild(override);
	};
	ensure(`/${origin}`, ORIGIN_TYPE);
	ensure(`/${signature}`, SIGNATURE_TYPE);
	zip.file(MEDIA_TYPES_ITEM, buildXml(doc));
}

/** The manifest references: every part, a relationship part through the relationship transform. */
async function manifestXml(zip: JSZip): Promise<string> {
	const types = parseContentTypes(await zip.file(MEDIA_TYPES_ITEM)?.async('string'));
	const references: string[] = [];
	for (const name of storedItems(zip).sort()) {
		if (isOutsideCoverage(name)) continue;
		if (!isPartName(name)) throw new Error(`"${name}" is not a valid part name.`);
		const bytes = await zip.file(name)!.async('uint8array');
		const contentType = contentTypeForPart(types, name);
		const uri = escapeXmlAttr(contentType ? `/${name}?ContentType=${contentType}` : `/${name}`);
		if (!isRels(name)) {
			references.push(
				`<Reference URI="${uri}"><DigestMethod Algorithm="${SHA256}"/>` +
					`<DigestValue>${sha256(bytes)}</DigestValue></Reference>`,
			);
			continue;
		}
		const ids = parseRelationshipsPart(decodeXmlBytes(bytes))
			.filter((rel) => rel.type !== DIGITAL_SIGNATURE_ORIGIN_REL_TYPE)
			.map((rel) => rel.id)
			.sort();
		const transform = {
			algorithm: OPC_RELATIONSHIP_TRANSFORM,
			relationshipReferenceIds: ids,
			relationshipGroupTypes: [],
		};
		const result = applyReferenceTransforms(bytes, [
			transform,
			{ algorithm: INCLUSIVE_C14N, relationshipReferenceIds: [], relationshipGroupTypes: [] },
		]);
		if (result.unsupportedAlgorithms.length)
			throw new Error(`"${name}" cannot be signed through the relationship transform.`);
		const selections = ids
			.map((id) => `<mdssi:RelationshipReference SourceId="${escapeXmlAttr(id)}"/>`)
			.join('');
		references.push(
			`<Reference URI="${uri}"><Transforms>` +
				`<Transform Algorithm="${OPC_RELATIONSHIP_TRANSFORM}">${selections}</Transform>` +
				`<Transform Algorithm="${INCLUSIVE_C14N}"/></Transforms>` +
				`<DigestMethod Algorithm="${SHA256}"/><DigestValue>${sha256(result.data)}</DigestValue>` +
				'</Reference>',
		);
	}
	return `<Manifest>${references.join('')}</Manifest>`;
}

/** `2026-11-27T09:30:00Z`: the signing time, to the second, in the format both objects use. */
const signingTimeOf = (date: Date) => date.toISOString().replace(/\.\d{3}Z$/u, 'Z');

/** A same-document reference to `id`, digested as the verifier reads it (inclusive c14n). */
function objectReference(doc: Document, id: string, type: string, transform: boolean): string {
	const target = Array.from(doc.getElementsByTagName('*')).find(
		(element) => element.getAttribute('Id') === id,
	);
	if (!target) throw new Error(`Missing signature object ${id}.`);
	const digest = sha256(utf8(canonicalize(target, INCLUSIVE_C14N)));
	const transforms = transform
		? `<Transforms><Transform Algorithm="${INCLUSIVE_C14N}"/></Transforms>`
		: '';
	return (
		`<Reference Type="${type}" URI="#${id}">${transforms}` +
		`<DigestMethod Algorithm="${SHA256}"/><DigestValue>${digest}</DigestValue></Reference>`
	);
}

/**
 * `bytes` with one more signature: the host's `signer` signs every part of the package. Fails,
 * with no output, when the package cannot be signed or the result does not verify.
 */
export async function signPackage(
	bytes: Uint8Array,
	signer: PackageSigner,
	options: SignPackageOptions = {},
): Promise<Uint8Array> {
	const certificate = await signer.certificate();
	const fields = readCertificate(certificate);
	if (!fields?.rsaPublicKey) throw new Error('The signing certificate is not an RSA certificate.');
	const zip = await JSZip.loadAsync(bytes);
	const { origin, signature } = await prepareOrigin(zip);
	await registerSignature(zip, origin, signature);
	const signingTime = signingTimeOf(options.now ?? new Date());
	const packageObject =
		`<Object Id="${PACKAGE_OBJECT_ID}" xmlns:mdssi="${OPC_DIGITAL_SIGNATURE_NS}">${await manifestXml(zip)}` +
		'<SignatureProperties><SignatureProperty Id="idSignatureTime" ' +
		`Target="#${SIGNATURE_ID}"><mdssi:SignatureTime><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>` +
		`<mdssi:Value>${signingTime}</mdssi:Value></mdssi:SignatureTime></SignatureProperty>` +
		'</SignatureProperties></Object>';
	const objects =
		packageObject +
		officeObjectXml(options) +
		qualifyingPropertiesXml(
			signingTime,
			{
				digestMethod: SHA256,
				digestValue: sha256(certificate),
				issuer: fields.issuer,
				serialNumber: BigInt(`0x${fields.serialNumber || '0'}`).toString(10),
			},
			options,
		);
	const open = `<Signature xmlns="${XMLDSIG_NS}" Id="${SIGNATURE_ID}">`;
	const context = parseStrictXml(`${open}${objects}</Signature>`);
	const signedInfo =
		`<SignedInfo><CanonicalizationMethod Algorithm="${INCLUSIVE_C14N}"/>` +
		`<SignatureMethod Algorithm="${RSA_SHA256}"/>` +
		objectReference(context, PACKAGE_OBJECT_ID, OBJECT_TYPE, false) +
		objectReference(context, OFFICE_OBJECT_ID, OBJECT_TYPE, false) +
		objectReference(context, SIGNED_PROPERTIES_ID, SIGNED_PROPERTIES_TYPE, true) +
		'</SignedInfo>';
	const signedInfoElement = childElements(
		parseStrictXml(`${open}${signedInfo}</Signature>`).documentElement,
	)[0]!;
	const value = await signer.sign(utf8(canonicalize(signedInfoElement, INCLUSIVE_C14N)));
	const xml =
		`<?xml version="1.0" encoding="UTF-8"?>${open}${signedInfo}` +
		`<SignatureValue>${encodeBase64(value)}</SignatureValue>` +
		`<KeyInfo><X509Data><X509Certificate>${encodeBase64(certificate)}</X509Certificate></X509Data></KeyInfo>` +
		`${objects}</Signature>`;
	zip.file(signature, xml);
	const signed = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	const summary = await verifySignedPackage(signed);
	const mine = summary.signatures.find((entry) => entry.path === signature);
	if (!mine || !isVerifiedSignatureState(mine.state))
		throw new Error(`The signed package does not verify (${mine?.state ?? summary.state}).`);
	return signed;
}
