/** Verification of `SignatureValue` over the canonical `SignedInfo`. Platform-neutral. */

import { decodeBase64 } from '../../../digest/index';
import { canonicalize } from './c14n';
import type { XmlSignature } from './document';
import { verifyRsaPkcs1 } from './rsa';
import { readCertificateBase64 } from './x509';

/**
 * The outcome of checking `SignatureValue`. `canonicalSignedInfo` is exactly the text the signer
 * signed: it exists only on success, and references are read from it and nowhere else.
 * `not-checked` means the signature uses a method or key type this verifier does not implement, so
 * it says nothing about the signature either way.
 */
export type SignedInfoVerification =
	| { status: 'verified'; canonicalSignedInfo: string; signerCertificate: string }
	| { status: 'invalid' | 'not-checked' };

const SIGNATURE_METHOD_TO_HASH: Readonly<Record<string, string>> = {
	'http://www.w3.org/2000/09/xmldsig#rsa-sha1': 'SHA-1',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256': 'SHA-256',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha384': 'SHA-384',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha512': 'SHA-512',
};

/**
 * Verify `SignatureValue` over the canonical `SignedInfo`. The signer is whichever `KeyInfo`
 * certificate verifies it; `KeyInfo` is unsigned, so its order proves nothing.
 */
export function verifySignedInfo(signature: XmlSignature): SignedInfoVerification {
	const hash = SIGNATURE_METHOD_TO_HASH[signature.signatureMethod];
	const keys = signature.certificates.flatMap((certificate) => {
		const key = readCertificateBase64(certificate)?.rsaPublicKey;
		return key ? [{ certificate, key }] : [];
	});
	if (!hash || keys.length === 0) return { status: 'not-checked' };
	const signatureValue = decodeBase64(signature.signatureValueBase64);
	if (!signatureValue || signatureValue.length === 0) return { status: 'invalid' };
	let canonicalSignedInfo: string;
	try {
		canonicalSignedInfo = canonicalize(signature.signedInfo, signature.canonicalizationMethod);
	} catch {
		return { status: 'invalid' };
	}
	const message = new TextEncoder().encode(canonicalSignedInfo);
	const signer = keys.find(({ key }) => verifyRsaPkcs1(key, hash, message, signatureValue));
	return signer
		? { status: 'verified', canonicalSignedInfo, signerCertificate: signer.certificate }
		: { status: 'invalid' };
}
