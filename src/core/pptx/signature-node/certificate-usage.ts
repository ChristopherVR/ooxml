/**
 * What a certificate's extensions permit (RFC 5280): signing documents and issuing further CAs.
 *
 * Node-only - depends on `node:crypto` and `node-forge` for ASN.1.
 */

import type crypto from 'node:crypto';

import forge from 'node-forge';

const KEY_USAGE_OID = '2.5.29.15';
const BASIC_CONSTRAINTS_OID = '2.5.29.19';
const DIGITAL_SIGNATURE_BIT = 0x80;
const NON_REPUDIATION_BIT = 0x40;

/** Extended key usages that permit document signing. */
const DOCUMENT_SIGNING_EKUS = new Set([
	'2.5.29.37.0', // anyExtendedKeyUsage
	'1.3.6.1.5.5.7.3.36', // id-kp-documentSigning (RFC 9336)
	'1.3.6.1.4.1.311.10.3.12', // Microsoft Document Signing
	'1.3.6.1.5.5.7.3.4', // id-kp-emailProtection, the usage of S/MIME signing certificates
]);

const CONTEXT_SPECIFIC = forge.asn1.Class.CONTEXT_SPECIFIC;

const children = (node: forge.asn1.Asn1 | undefined): forge.asn1.Asn1[] =>
	Array.isArray(node?.value) ? node.value : [];

/** The DER contents of an extension, or `undefined` when the certificate does not have it. */
function extensionValue(certificate: crypto.X509Certificate, extensionOid: string) {
	const der = forge.util.createBuffer(certificate.raw.toString('binary'));
	const tbs = children(forge.asn1.fromDer(der))[0];
	const extensionsField = children(tbs).find(
		(node) => node.tagClass === CONTEXT_SPECIFIC && node.type === 3,
	);
	for (const extension of children(children(extensionsField)[0])) {
		const fields = children(extension);
		const oid = fields[0];
		const value = fields[fields.length - 1];
		if (typeof oid?.value !== 'string' || typeof value?.value !== 'string') continue;
		if (forge.asn1.derToOid(oid.value) === extensionOid) return value.value;
	}
	return undefined;
}

/** `keyUsage` is a DER BIT STRING: tag, length, unused bit count, then the flag bytes. */
function keyUsageAllowsSigning(extension: string): boolean {
	const flags = extension.charCodeAt(3);
	return extension.length >= 4 && (flags & (DIGITAL_SIGNATURE_BIT | NON_REPUDIATION_BIT)) !== 0;
}

/**
 * A certificate may sign documents when its key usage, if stated, includes `digitalSignature` or
 * `nonRepudiation`, and its extended key usage, if stated, includes a document signing purpose.
 * A TLS server certificate fails the second test.
 */
export function isDocumentSigningCertificate(certificate: crypto.X509Certificate): boolean {
	const extendedUsages = certificate.keyUsage;
	if (extendedUsages && !extendedUsages.some((oid) => DOCUMENT_SIGNING_EKUS.has(oid))) {
		return false;
	}
	const keyUsage = extensionValue(certificate, KEY_USAGE_OID);
	return keyUsage === undefined || keyUsageAllowsSigning(keyUsage);
}

/** The `pathLenConstraint` of `basicConstraints`; throws when the extension is malformed. */
export function pathLengthConstraint(certificate: crypto.X509Certificate): number | undefined {
	const extension = extensionValue(certificate, BASIC_CONSTRAINTS_OID);
	if (extension === undefined) return undefined;
	const fields = children(forge.asn1.fromDer(forge.util.createBuffer(extension)));
	const pathLength = fields.find((field) => field.type === forge.asn1.Type.INTEGER);
	return pathLength ? forge.asn1.derToInteger(pathLength.value as string) : undefined;
}
