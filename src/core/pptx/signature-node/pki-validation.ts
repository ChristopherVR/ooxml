/**
 * PKI validation utilities: PEM/fingerprint certificate helpers and the
 * timestamp status of OOXML digital signatures.
 *
 * OCSP revocation checking lives in `./ocsp.ts`.
 *
 * Node-only: depends on `node:crypto`.
 */

import crypto from 'node:crypto';

import type { TimestampAuthorityStatus } from '../core/utils/signature-types';
import { findTimestampTagContent } from './timestamp-tag-scanner';

// ---------------------------------------------------------------------------
// Certificate helpers
// ---------------------------------------------------------------------------

/** Wrap a Base64-encoded DER certificate in PEM armour. */
export function certPemFromBase64(certBase64: string): string | undefined {
	try {
		const der = Buffer.from(certBase64, 'base64');
		if (der.length === 0) {
			return undefined;
		}
		const b64 = der.toString('base64');
		const lines = b64.match(/.{1,64}/g);
		if (!lines) {
			return undefined;
		}
		return `-----BEGIN CERTIFICATE-----\n${lines.join('\n')}\n-----END CERTIFICATE-----`;
	} catch {
		return undefined;
	}
}

/** SHA-256 fingerprint of a PEM certificate (lowercase hex, no colons). */
export function certFingerprintSha256(certPem: string): string | undefined {
	try {
		const cert = new crypto.X509Certificate(certPem);
		return cert.fingerprint256.replace(/:/g, '').toLowerCase();
	} catch {
		return undefined;
	}
}

// ---------------------------------------------------------------------------
// Public: timestamp authority evaluation
// ---------------------------------------------------------------------------

/**
 * Timestamp tokens sit in unsigned markup and are not validated: their CMS signature, TSA chain and
 * message imprint are never checked. So a token is reported `not-checked`, never `valid`, and a
 * policy that requires a timestamp is not met by one.
 */
export async function evaluateTimestampAuthority(
	signatureXml: string,
): Promise<{ status: TimestampAuthorityStatus; error?: string }> {
	if (!findTimestampTagContent(signatureXml)?.trim()) {
		return { status: 'not-present' };
	}
	return { status: 'not-checked', error: 'Timestamp tokens are not validated.' };
}
