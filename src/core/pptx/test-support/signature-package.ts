/**
 * Builds small OPC packages and signatures for the digital signature tests, using the layout
 * ECMA-376 Part 2 describes and Office writes: `SignedInfo` protects an `Object`, and a `Manifest`
 * inside that `Object` lists the parts.
 */

import crypto from 'node:crypto';

import JSZip from 'jszip';
import forge from 'node-forge';

import { OPC_RELATIONSHIP_TRANSFORM, XMLDSIG_NS } from '../core/utils/signature-constants';
import { applyReferenceTransforms } from '../signature-node/reference-transforms';
import {
	canonicalizeNode,
	childElements,
	INCLUSIVE_C14N,
	parseStrictXml,
} from '../signature-node/xml-canonicalization';

const SHA256 = 'http://www.w3.org/2001/04/xmlenc#sha256';
const RSA_SHA256 = 'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256';
const MDSSI = 'http://schemas.openxmlformats.org/package/2006/digital-signature';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml';
const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
export const SLIDE_PART = 'ppt/slides/slide1.xml';
export const CORE_PROPERTIES_RELATIONSHIP = `${RELS}/metadata/core-properties`;
export const OFFICE_DOCUMENT_RELATIONSHIP = `${OFFICE_RELS}/officeDocument`;
export const SLIDE_RELATIONSHIP = `${OFFICE_RELS}/slide`;
export const CUSTOM_XML_RELATIONSHIP = `${OFFICE_RELS}/customXml`;
export const SLIDE_CONTENT_TYPE = `${PML}.slide+xml`;

export interface TestParty {
	cert: forge.pki.Certificate;
	key: forge.pki.rsa.PrivateKey;
	keyPem: string;
	certPem: string;
	certBase64: string;
}

export interface TestPartyOptions {
	name: string;
	issuer?: TestParty;
	ca?: boolean;
	/** The `pathLenConstraint` of `basicConstraints`. */
	pathLength?: number;
	extensions?: object[];
}

/** Create an RSA certificate, self-signed unless `issuer` is given. */
export function createTestParty({
	name,
	issuer,
	ca = false,
	pathLength,
	extensions = [],
}: TestPartyOptions) {
	const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
	const keyPem = privateKey.export({ type: 'pkcs1', format: 'pem' }) as string;
	const key = forge.pki.privateKeyFromPem(keyPem) as forge.pki.rsa.PrivateKey;
	const cert = forge.pki.createCertificate();
	cert.publicKey = forge.pki.publicKeyFromPem(
		publicKey.export({ type: 'spki', format: 'pem' }) as string,
	);
	cert.serialNumber = String(Math.floor(Math.random() * 1e9) + 1);
	cert.validity.notBefore = new Date(Date.now() - 24 * 3600 * 1000);
	cert.validity.notAfter = new Date(Date.now() + 365 * 24 * 3600 * 1000);
	const subject = [{ name: 'commonName', value: name }];
	cert.setSubject(subject);
	cert.setIssuer(issuer ? issuer.cert.subject.attributes : subject);
	cert.setExtensions([
		{
			name: 'basicConstraints',
			cA: ca,
			...(pathLength === undefined ? {} : { pathLenConstraint: pathLength }),
		},
		...extensions,
	]);
	cert.sign(issuer ? issuer.key : key, forge.md.sha256.create());
	const der = forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes();
	const party: TestParty = {
		cert,
		key,
		keyPem,
		certPem: forge.pki.certificateToPem(cert),
		certBase64: forge.util.encode64(der),
	};
	return party;
}

const relationship = (id: string, type: string, target: string) =>
	`<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

export const relationshipsXml = (...items: string[]) =>
	`<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="${RELS}">\n  ${items.join('\n  ')}\n</Relationships>`;

export { relationship as relationshipXml };

/** A minimal presentation package with one slide and core properties. */
export function createTestPackage(): JSZip {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="${RELS_CONTENT_TYPE}"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/ppt/presentation.xml" ContentType="${PML}.presentation.main+xml"/>` +
			`<Override PartName="/${SLIDE_PART}" ContentType="${SLIDE_CONTENT_TYPE}"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		relationshipsXml(
			relationship('rId1', OFFICE_DOCUMENT_RELATIONSHIP, 'ppt/presentation.xml'),
			relationship('rId2', CORE_PROPERTIES_RELATIONSHIP, 'docProps/core.xml'),
		),
	);
	zip.file('docProps/core.xml', '<coreProperties/>');
	zip.file('ppt/presentation.xml', '<presentation/>');
	zip.file(
		'ppt/_rels/presentation.xml.rels',
		relationshipsXml(relationship('rId1', SLIDE_RELATIONSHIP, 'slides/slide1.xml')),
	);
	zip.file(SLIDE_PART, '<sld>original</sld>');
	return zip;
}

const digest = (data: Uint8Array | string) =>
	crypto
		.createHash('sha256')
		.update(typeof data === 'string' ? Buffer.from(data, 'utf8') : data)
		.digest('base64');

const referenceXml = (uri: string, digestValue: string, transforms = '') =>
	`<Reference URI="${uri}">${transforms}<DigestMethod Algorithm="${SHA256}"/>` +
	`<DigestValue>${digestValue}</DigestValue></Reference>`;

/** What a relationships transform selects: relationship `Id`s or relationship types. */
export interface RelationshipSelection {
	ids?: string[];
	types?: string[];
}

async function relationshipsReference(zip: JSZip, part: string, selection: RelationshipSelection) {
	const { ids = [], types = [] } = selection;
	const parameters = [
		...ids.map((id) => `<mdssi:RelationshipReference SourceId="${id}"/>`),
		...types.map((type) => `<mdssi:RelationshipsGroupReference SourceType="${type}"/>`),
	];
	const transforms = [
		{
			algorithm: OPC_RELATIONSHIP_TRANSFORM,
			relationshipReferenceIds: ids,
			relationshipGroupTypes: types,
		},
		{ algorithm: INCLUSIVE_C14N, relationshipReferenceIds: [], relationshipGroupTypes: [] },
	];
	const { data } = applyReferenceTransforms(await zip.file(part)!.async('uint8array'), transforms);
	return referenceXml(
		`/${part}?ContentType=${RELS_CONTENT_TYPE}`,
		digest(data),
		`<Transforms><Transform Algorithm="${OPC_RELATIONSHIP_TRANSFORM}">${parameters.join('')}</Transform>` +
			`<Transform Algorithm="${INCLUSIVE_C14N}"/></Transforms>`,
	);
}

/** A `Manifest` reference to a whole part, under `uri` when it differs from the stored name. */
export async function partReference(
	zip: JSZip,
	part: string,
	contentType: string,
	uri = `/${part}`,
) {
	const bytes = await zip.file(part)!.async('uint8array');
	return referenceXml(`${uri}?ContentType=${contentType}`, digest(bytes));
}

/**
 * The `Manifest` references of the test package: everything but core properties. Relationships are
 * selected by `Id`, or by type with `selectBy: 'type'`.
 */
export async function createManifestReferences(
	zip: JSZip,
	{ selectBy = 'id' }: { selectBy?: 'id' | 'type' } = {},
): Promise<string[]> {
	const select = (id: string, type: string): RelationshipSelection =>
		selectBy === 'id' ? { ids: [id] } : { types: [type] };
	return [
		await relationshipsReference(zip, '_rels/.rels', select('rId1', OFFICE_DOCUMENT_RELATIONSHIP)),
		await relationshipsReference(
			zip,
			'ppt/_rels/presentation.xml.rels',
			select('rId1', SLIDE_RELATIONSHIP),
		),
		await partReference(zip, 'ppt/presentation.xml', `${PML}.presentation.main+xml`),
		await partReference(zip, SLIDE_PART, SLIDE_CONTENT_TYPE),
	];
}

export interface OfficeSignatureOptions {
	signer: TestParty;
	manifestReferences: string[];
	/** Extra `KeyInfo` certificates, placed before the signer's. */
	extraCertificates?: string[];
	/** Markup appended inside `Signature`, outside anything the signature protects. */
	unsignedSuffix?: string;
}

/** Build a signature part with the Office layout. */
export function createOfficeSignatureXml(options: OfficeSignatureOptions): string {
	const { signer, manifestReferences, extraCertificates = [], unsignedSuffix = '' } = options;
	const objectXml =
		`<Object Id="idPackageObject"><Manifest>${manifestReferences.join('')}</Manifest>` +
		`<SignatureProperties><SignatureProperty Id="idSignatureTime" Target="#idPackageSignature">` +
		`<mdssi:SignatureTime xmlns:mdssi="${MDSSI}"><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>` +
		`<mdssi:Value>2026-01-02T03:04:05Z</mdssi:Value></mdssi:SignatureTime></SignatureProperty>` +
		`</SignatureProperties></Object>`;
	const open = `<Signature xmlns="${XMLDSIG_NS}" xmlns:mdssi="${MDSSI}" Id="idPackageSignature">`;
	const objectNode = childElements(
		parseStrictXml(`${open}${objectXml}</Signature>`).documentElement!,
	)[0]!;
	const signedInfo =
		`<SignedInfo><CanonicalizationMethod Algorithm="${INCLUSIVE_C14N}"/>` +
		`<SignatureMethod Algorithm="${RSA_SHA256}"/>` +
		`<Reference Type="${XMLDSIG_NS}Object" URI="#idPackageObject">` +
		`<DigestMethod Algorithm="${SHA256}"/>` +
		`<DigestValue>${digest(canonicalizeNode(objectNode, INCLUSIVE_C14N))}</DigestValue>` +
		`</Reference></SignedInfo>`;
	const signedInfoNode = childElements(
		parseStrictXml(`${open}${signedInfo}</Signature>`).documentElement!,
	)[0]!;
	const signatureValue = crypto
		.createSign('RSA-SHA256')
		.update(canonicalizeNode(signedInfoNode, INCLUSIVE_C14N), 'utf8')
		.sign(signer.keyPem, 'base64');
	const certificates = [...extraCertificates, signer.certBase64]
		.map((certificate) => `<X509Certificate>${certificate}</X509Certificate>`)
		.join('');
	return (
		`<?xml version="1.0" encoding="UTF-8"?>\n${open}${signedInfo}` +
		`<SignatureValue>${signatureValue}</SignatureValue>` +
		`<KeyInfo><X509Data>${certificates}</X509Data></KeyInfo>${objectXml}${unsignedSuffix}</Signature>`
	);
}

/** Add a signature part (with its origin part) and serialize the package. */
export async function withSignature(zip: JSZip, signatureXml: string): Promise<Uint8Array> {
	zip.file('_xmlsignatures/sig1.xml', signatureXml);
	zip.file('_xmlsignatures/origin.sigs', '');
	return zip.generateAsync({ type: 'uint8array' });
}
