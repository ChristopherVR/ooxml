/**
 * Certificate handling utilities for digital signature processing.
 *
 * Node-only - depends on `node:crypto` and `node-forge`.
 */

import crypto from 'node:crypto';
import path from 'node:path';

import forge from 'node-forge';

import type {
	LoadedSigningMaterial,
	SignatureCertificateInfo,
} from '../core/utils/signature-types';
import { isDocumentSigningCertificate, pathLengthConstraint } from './certificate-usage';
import { extractFirstPemBlock } from './pem-utils';

const PRIVATE_KEY_PEM_LABELS = [
	'RSA PRIVATE KEY',
	'EC PRIVATE KEY',
	'ENCRYPTED PRIVATE KEY',
	'PRIVATE KEY',
];

/** Extract certificate metadata from a Base64-encoded DER certificate. */
export function certificateInfoFromBase64(
	certBase64: string,
): SignatureCertificateInfo | undefined {
	try {
		const certificate = new crypto.X509Certificate(Buffer.from(certBase64, 'base64'));
		return {
			subject: certificate.subject || undefined,
			issuer: certificate.issuer || undefined,
			serialNumber: certificate.serialNumber || undefined,
			validFrom: certificate.validFrom || undefined,
			validTo: certificate.validTo || undefined,
		};
	} catch {
		return undefined;
	}
}

/** The parseable certificates among `pems`; a malformed one is skipped, never trusted. */
function parseCertificates(pems: readonly string[]): crypto.X509Certificate[] {
	const parsed: crypto.X509Certificate[] = [];
	for (const pem of pems) {
		try {
			parsed.push(new crypto.X509Certificate(pem));
		} catch {
			// not a certificate: it cannot anchor trust
		}
	}
	return parsed;
}

function isWithinValidity(certificate: crypto.X509Certificate, now: number): boolean {
	return now >= Date.parse(certificate.validFrom) && now <= Date.parse(certificate.validTo);
}

/** `issuer` issued `certificate`: the names chain and OpenSSL verifies the signature. */
function isSignedBy(certificate: crypto.X509Certificate, issuer: crypto.X509Certificate): boolean {
	return certificate.checkIssued(issuer) && certificate.verify(issuer.publicKey);
}

/**
 * Each CA's `pathLenConstraint` limits the non-self-issued intermediates below it, the leaf aside
 * (RFC 5280, 4.2.1.9). `path` runs from the leaf to the trust root.
 */
function pathLengthsHold(path: readonly crypto.X509Certificate[]): boolean {
	let intermediatesBelow = 0;
	for (const [index, certificate] of path.entries()) {
		if (index === 0) continue;
		const limit = pathLengthConstraint(certificate);
		if (limit !== undefined && intermediatesBelow > limit) return false;
		if (certificate.subject !== certificate.issuer) intermediatesBelow++;
	}
	return true;
}

/**
 * A leaf-first chain is trusted when every certificate is within its validity period, each is
 * signed by the next (which must be a CA), the last is a trusted root or signed by one, and every
 * path length constraint on the way holds. Signatures are verified by OpenSSL through `node:crypto`.
 */
function chainIsTrusted(
	chain: readonly crypto.X509Certificate[],
	roots: readonly crypto.X509Certificate[],
	now: number,
): boolean {
	const top = chain[chain.length - 1];
	if (!top || !chain.every((certificate) => isWithinValidity(certificate, now))) return false;
	for (let index = 0; index < chain.length - 1; index++) {
		const issuer = chain[index + 1]!;
		if (!issuer.ca || !isSignedBy(chain[index]!, issuer)) return false;
	}
	return roots.some((root) => {
		if (!isWithinValidity(root, now)) return false;
		if (root.fingerprint256 === top.fingerprint256) return pathLengthsHold(chain);
		return root.ca && isSignedBy(top, root) && pathLengthsHold([...chain, root]);
	});
}

/**
 * Order `KeyInfo` certificates leaf first, from the certificate that verified the signature.
 * `KeyInfo` is unsigned: certificates that do not extend the path are dropped, not trusted.
 */
export function orderCertificateChain(
	signerBase64: string,
	certBase64List: readonly string[],
): string[] {
	const parsed = (value: string) => new crypto.X509Certificate(Buffer.from(value, 'base64'));
	const chain = [signerBase64];
	const remaining = certBase64List.filter((value) => value !== signerBase64);
	let current = parsed(signerBase64);
	while (remaining.length > 0) {
		const next = remaining.findIndex((value) => {
			try {
				return isSignedBy(current, parsed(value));
			} catch {
				return false;
			}
		});
		if (next < 0) break;
		const [issuer] = remaining.splice(next, 1);
		chain.push(issuer!);
		current = parsed(issuer!);
	}
	return chain;
}

/**
 * Validate a leaf-first certificate chain for document signing. The leaf must permit document
 * signing and the chain must end at one of `trustRootsPem`. No roots are trusted by default: the
 * platform's TLS roots say who may serve a website, not who may sign a document.
 */
export function validateCertificateChain(
	certBase64List: string[],
	trustRootsPem: string[],
): {
	status: 'trusted' | 'untrusted' | 'not-checked';
	error?: string;
} {
	if (certBase64List.length === 0) {
		return { status: 'not-checked' };
	}
	try {
		const chain = certBase64List.map(
			(value) => new crypto.X509Certificate(Buffer.from(value, 'base64')),
		);
		if (!isDocumentSigningCertificate(chain[0]!)) {
			return {
				status: 'untrusted',
				error: 'The signing certificate does not permit document signing.',
			};
		}
		const verified = chainIsTrusted(chain, parseCertificates(trustRootsPem), Date.now());
		return verified
			? { status: 'trusted' }
			: {
					status: 'untrusted',
					error: 'Certificate chain is not trusted.',
				};
	} catch (error) {
		return {
			status: 'untrusted',
			error: `Certificate trust validation failed: ${String(error)}`,
		};
	}
}

/**
 * Load a private key and certificate from a PKCS#12 (.pfx/.p12) or PEM buffer.
 */
export function loadSigningMaterialFromBuffer(
	certificateBuffer: Uint8Array,
	certificatePath: string,
	certificatePassword?: string,
): LoadedSigningMaterial {
	const extension = path.extname(certificatePath).toLowerCase();
	if (extension === '.pfx' || extension === '.p12') {
		const p12Der = forge.util.createBuffer(Buffer.from(certificateBuffer).toString('binary'));
		const p12Asn1 = forge.asn1.fromDer(p12Der);
		const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certificatePassword || '');
		const keyBagOid = forge.pki.oids.pkcs8ShroudedKeyBag;
		const certBagOid = forge.pki.oids.certBag;
		const keyBag =
			keyBagOid === undefined ? undefined : p12.getBags({ bagType: keyBagOid })[keyBagOid];
		const certBag =
			certBagOid === undefined ? undefined : p12.getBags({ bagType: certBagOid })[certBagOid];
		if (!keyBag || keyBag.length === 0 || !keyBag[0]?.key) {
			throw new Error('No private key found in PKCS#12 certificate.');
		}
		if (!certBag || certBag.length === 0 || !certBag[0]?.cert) {
			throw new Error('No certificate found in PKCS#12 certificate.');
		}
		return {
			privateKeyPem: forge.pki.privateKeyToPem(keyBag[0].key),
			certificatePem: forge.pki.certificateToPem(certBag[0].cert),
		};
	}

	const pem = Buffer.from(certificateBuffer).toString('utf8');
	const privateKeyPem = extractFirstPemBlock(pem, PRIVATE_KEY_PEM_LABELS);
	const certificatePem = extractFirstPemBlock(pem, ['CERTIFICATE']);
	if (!privateKeyPem || !certificatePem) {
		throw new Error('PEM certificate must contain both private key and certificate.');
	}
	return {
		privateKeyPem,
		certificatePem,
	};
}

/** Convert a PEM certificate to Base64-encoded DER (strip armour + whitespace). */
export function pemCertificateToBase64(pem: string): string {
	return pem
		.replace(/-----BEGIN CERTIFICATE-----/g, '')
		.replace(/-----END CERTIFICATE-----/g, '')
		.replace(/\s+/g, '')
		.trim();
}
