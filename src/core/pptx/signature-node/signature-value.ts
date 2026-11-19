/**
 * Cryptographic verification of `SignedInfo` against the signature's certificates.
 *
 * Node-only - depends on `node:crypto`.
 */

import crypto from 'node:crypto';

import type { XmlSignature } from './signature-document';
import { canonicalizeNode } from './xml-canonicalization';

/**
 * The outcome of checking `SignatureValue`. `canonicalSignedInfo` is exactly the text the signer
 * signed: it exists only on success, and references are read from it and nowhere else.
 */
export type SignedInfoVerification =
	| { status: 'verified'; canonicalSignedInfo: string; signerCertificate: string }
	| { status: 'invalid' | 'not-checked' };

const SIGNATURE_METHOD_TO_VERIFY_ALGORITHM: Record<string, string> = {
	'http://www.w3.org/2000/09/xmldsig#rsa-sha1': 'RSA-SHA1',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256': 'RSA-SHA256',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha384': 'RSA-SHA384',
	'http://www.w3.org/2001/04/xmldsig-more#rsa-sha512': 'RSA-SHA512',
};

function verifiesWith(
	certificateBase64: string,
	algorithm: string,
	canonicalSignedInfo: string,
	signatureValue: Buffer,
): boolean {
	try {
		const certificate = new crypto.X509Certificate(Buffer.from(certificateBase64, 'base64'));
		return crypto
			.createVerify(algorithm)
			.update(canonicalSignedInfo, 'utf8')
			.verify(certificate.publicKey, signatureValue);
	} catch {
		return false;
	}
}

/**
 * Verify `SignatureValue` over the canonical `SignedInfo`. The signer is whichever `KeyInfo`
 * certificate verifies it; `KeyInfo` is unsigned, so its order proves nothing.
 */
export function verifySignedInfo(signature: XmlSignature): SignedInfoVerification {
	if (signature.certificates.length === 0) {
		return { status: 'not-checked' };
	}
	const algorithm = SIGNATURE_METHOD_TO_VERIFY_ALGORITHM[signature.signatureMethod];
	if (!algorithm || signature.signatureValueBase64.length === 0) {
		return { status: 'invalid' };
	}
	try {
		const canonicalSignedInfo = canonicalizeNode(
			signature.signedInfo,
			signature.canonicalizationMethod,
		);
		const signatureValue = Buffer.from(signature.signatureValueBase64, 'base64');
		const signerCertificate = signature.certificates.find((certificate) =>
			verifiesWith(certificate, algorithm, canonicalSignedInfo, signatureValue),
		);
		return signerCertificate
			? { status: 'verified', canonicalSignedInfo, signerCertificate }
			: { status: 'invalid' };
	} catch {
		return { status: 'invalid' };
	}
}
