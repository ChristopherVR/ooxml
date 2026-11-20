/**
 * PPTX digital signature creation.
 *
 * Node-only - signs all content in a PPTX package with a certificate
 * and returns the signed data along with a verification report.
 */

import crypto from 'node:crypto';

import JSZip from 'jszip';
import { contentTypeForPart, parseContentTypes } from '../../opc/content-types';
import { isPartName, MEDIA_TYPES_ITEM } from '../../opc/part-names';
import {
	buildRelationshipsXml,
	isRelationshipsItem,
	nextRelationshipId,
	parseRelationshipsPart,
	type RelationshipInput,
} from '../../opc/relationships';
import { isSignaturePart } from '../../opc/signature/detection';
import { hasVerifiedIntegrity } from '../../opc/signature/inspection-status';
import { buildXml, decodeXmlBytes, parseXml } from '../../xml/index';
import { openPptxZip } from '../core/utils/zip-open';

import {
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
	EXCLUSIVE_C14N,
	INCLUSIVE_C14N,
	XMLDSIG_NS,
} from '../../opc/signature/constants';
import type {
	DigitalSignatureReport,
	OfficeSignatureReference,
	SignOptions,
	SignResult,
} from '../core/utils/signature-types';
import { escapeXmlAttr, escapeXmlText, isValidBase64 } from '../core/utils/signature-xml-utils';
import { loadSigningMaterialFromBuffer, pemCertificateToBase64 } from './certificate-utils';
import { inspectPptxDigitalSignatures } from './inspection';
import { isOutsideCoverage } from '../../opc/signature/verify/coverage';
import { canonicalize } from '../../opc/signature/verify/c14n';
import { childElements, parseStrictXml } from '../../opc/signature/verify/dom';

const PACKAGE_RELATIONSHIPS_PART = '_rels/.rels';

const storedItemNames = (zip: JSZip): string[] =>
	Object.keys(zip.files).filter((name) => !zip.files[name]?.dir);

/**
 * Refuses a package a signature could not describe or the loader would not open once signed: an
 * item whose name is not a valid part name cannot be referenced, and a signed package with a
 * relationships part the strict reader rejects does not load.
 */
async function assertSignable(zip: JSZip): Promise<void> {
	for (const name of storedItemNames(zip)) {
		if (isOutsideCoverage(name)) continue;
		if (!isPartName(name)) {
			throw new Error(`"${name}" is not a valid part name.`);
		}
		if (!isRelationshipsItem(name)) continue;
		try {
			parseRelationshipsPart(decodeXmlBytes(await zip.file(name)!.async('uint8array')));
		} catch (error) {
			throw new Error(`"${name}" is not a valid relationships part.`, { cause: error });
		}
	}
}

/** Points the package's one origin relationship at the origin part this signer writes. */
async function upsertRootOriginRelationship(zip: JSZip): Promise<void> {
	const bytes = await zip.file(PACKAGE_RELATIONSHIPS_PART)?.async('uint8array');
	const xml = bytes && decodeXmlBytes(bytes);
	const relationships = xml === undefined ? [] : parseRelationshipsPart(xml);
	const id =
		relationships.find(({ type }) => type === DIGITAL_SIGNATURE_ORIGIN_REL_TYPE)?.id ??
		nextRelationshipId(new Set(relationships.map((relationship) => relationship.id)));
	const kept = new Map<string, RelationshipInput>(
		relationships
			.filter(({ type }) => type !== DIGITAL_SIGNATURE_ORIGIN_REL_TYPE)
			.map((relationship) => [relationship.id, relationship]),
	);
	kept.set(id, {
		type: DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
		target: '_xmlsignatures/origin.sigs',
		mode: 'Internal',
	});
	zip.file(PACKAGE_RELATIONSHIPS_PART, buildRelationshipsXml(kept));
}

async function upsertContentTypesForSignature(zip: JSZip, signaturePath: string): Promise<void> {
	const xml = await zip.file(MEDIA_TYPES_ITEM)?.async('string');
	if (!xml) {
		return;
	}
	const doc = parseXml(xml);
	const types = doc.documentElement;
	const ensureOverride = (partName: string, contentType: string): void => {
		const existing = Array.from(types.getElementsByTagName('Override')).find(
			(node) => node.getAttribute('PartName') === partName,
		);
		if (existing) {
			existing.setAttribute('ContentType', contentType);
			return;
		}
		const override = doc.createElement('Override');
		override.setAttribute('PartName', partName);
		override.setAttribute('ContentType', contentType);
		types.appendChild(override);
	};
	ensureOverride(
		'/_xmlsignatures/origin.sigs',
		'application/vnd.openxmlformats-package.digital-signature-origin',
	);
	ensureOverride(
		`/${signaturePath}`,
		'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml',
	);
	zip.file(MEDIA_TYPES_ITEM, buildXml(doc));
}

async function buildOfficeReferenceList(zip: JSZip): Promise<OfficeSignatureReference[]> {
	const references: OfficeSignatureReference[] = [];
	const digestMethod = 'http://www.w3.org/2001/04/xmlenc#sha256';
	const contentTypes = parseContentTypes(await zip.file(MEDIA_TYPES_ITEM)?.async('string'));
	for (const entryPath of storedItemNames(zip)) {
		// The Media Types stream is not a part; each reference binds its part's type instead.
		if (isOutsideCoverage(entryPath)) {
			continue;
		}
		const entryBytes = await zip.file(entryPath)!.async('uint8array');
		const digestValue = crypto
			.createHash('sha256')
			.update(Buffer.from(entryBytes))
			.digest('base64');
		const contentType = contentTypeForPart(contentTypes, entryPath);
		references.push({
			uri: contentType ? `/${entryPath}?ContentType=${contentType}` : `/${entryPath}`,
			digestMethod,
			digestValue,
		});
	}
	references.sort((left, right) => left.uri.localeCompare(right.uri));
	return references;
}

const SIGNING_TIME_OBJECT_ID = 'idSigningTime';

/** The `Object` carrying the signing time, and the reference that protects it. */
function buildSigningTimeObject(digestMethod: string): {
	objectXml: string;
	reference: OfficeSignatureReference;
} {
	const objectXml =
		`<Object Id="${SIGNING_TIME_OBJECT_ID}"><SignatureProperties>` +
		`<SignatureProperty Target="#idPackageSignature">` +
		`<SigningTime>${escapeXmlText(new Date().toISOString())}</SigningTime>` +
		'</SignatureProperty></SignatureProperties></Object>';
	const context = parseStrictXml(`<Signature xmlns="${XMLDSIG_NS}">${objectXml}</Signature>`);
	const canonical = canonicalize(childElements(context.documentElement)[0]!, INCLUSIVE_C14N);
	return {
		objectXml,
		reference: {
			uri: `#${SIGNING_TIME_OBJECT_ID}`,
			digestMethod,
			digestValue: crypto.createHash('sha256').update(canonical, 'utf8').digest('base64'),
		},
	};
}

function buildOfficeSignatureXml(
	partReferences: OfficeSignatureReference[],
	privateKeyPem: string,
	certificatePem: string,
): string {
	const { objectXml, reference: signingTimeReference } = buildSigningTimeObject(
		'http://www.w3.org/2001/04/xmlenc#sha256',
	);
	const references = [...partReferences, signingTimeReference];
	const signedInfoReferencesXml = references
		.map(
			(reference) =>
				`<Reference URI="${escapeXmlAttr(reference.uri)}">` +
				`<DigestMethod Algorithm="${escapeXmlAttr(reference.digestMethod)}"/>` +
				`<DigestValue>${escapeXmlText(reference.digestValue)}</DigestValue>` +
				'</Reference>',
		)
		.join('');
	const signedInfoXml =
		`<SignedInfo xmlns="${XMLDSIG_NS}">` +
		'<CanonicalizationMethod Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>' +
		'<SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>' +
		`${signedInfoReferencesXml}` +
		'</SignedInfo>';
	const canonicalSignedInfo = canonicalize(
		parseStrictXml(signedInfoXml).documentElement,
		EXCLUSIVE_C14N,
	);
	const signer = crypto.createSign('RSA-SHA256');
	signer.update(Buffer.from(canonicalSignedInfo, 'utf8'));
	signer.end();
	const signatureValueBase64 = signer.sign(privateKeyPem).toString('base64');
	const certificateBase64 = pemCertificateToBase64(certificatePem);
	if (!isValidBase64(signatureValueBase64)) {
		throw new Error('Generated SignatureValue is not valid base64');
	}
	if (!isValidBase64(certificateBase64)) {
		throw new Error('X.509 certificate is not valid base64');
	}
	return (
		`<Signature xmlns="${XMLDSIG_NS}" Id="idPackageSignature">${signedInfoXml}` +
		`<SignatureValue>${escapeXmlText(signatureValueBase64)}</SignatureValue>` +
		`<KeyInfo><X509Data><X509Certificate>${escapeXmlText(certificateBase64)}</X509Certificate></X509Data></KeyInfo>` +
		`${objectXml}</Signature>`
	);
}

/**
 * Sign all content in a PPTX package with a certificate.
 *
 * Removes any existing signatures, creates a new XML-DSig signature
 * covering all non-signature parts, and returns the signed data
 * along with a post-sign verification report. Fails, with no signed data,
 * for a package that cannot be represented: invalid part names, invalid
 * relationships parts, or output whose signature does not verify and cover
 * its content. Trust and policy outcomes do not fail signing; read them
 * from the report.
 */
export async function signPptxWithCertificate(
	data: Uint8Array,
	certificateBuffer: Uint8Array,
	options: SignOptions,
): Promise<SignResult> {
	try {
		const signing = loadSigningMaterialFromBuffer(
			certificateBuffer,
			options.certificatePath,
			options.certificatePassword,
		);
		const zip = await openPptxZip(Buffer.from(data));

		for (const entryPath of Object.keys(zip.files)) {
			if (isSignaturePart(entryPath)) {
				zip.remove(entryPath);
			}
		}
		await assertSignable(zip);

		const signaturePath = '_xmlsignatures/sig1.xml';
		await upsertRootOriginRelationship(zip);
		await upsertContentTypesForSignature(zip, signaturePath);
		const references = await buildOfficeReferenceList(zip);
		const signatureXml = buildOfficeSignatureXml(
			references,
			signing.privateKeyPem,
			signing.certificatePem,
		);
		zip.file(signaturePath, signatureXml);
		zip.file('_xmlsignatures/origin.sigs', '<SignatureOrigin/>');
		zip.file(
			'_xmlsignatures/_rels/origin.sigs.rels',
			`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSig1" Type="${DIGITAL_SIGNATURE_REL_TYPE}" Target="sig1.xml"/></Relationships>`,
		);

		const signedData = await zip.generateAsync({ type: 'uint8array' });
		const report = await inspectPptxDigitalSignatures(signedData);
		if (!report.details?.every(hasVerifiedIntegrity)) {
			throw new Error(`The signed package does not verify (${report.verificationStatus}).`);
		}
		return { success: true, signedData, report };
	} catch (error) {
		const errorReport: DigitalSignatureReport = {
			supported: true,
			hasSignature: false,
			signatureCount: 0,
			signaturePaths: [],
			verificationStatus: 'error',
			error: `Failed to sign PPTX: ${String(error)}`,
		};
		return {
			success: false,
			report: errorReport,
			error: `Failed to sign PPTX: ${String(error)}`,
		};
	}
}
