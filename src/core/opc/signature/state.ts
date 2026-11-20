/**
 * The one answer every product shows for a package's digital signatures. A document is presented
 * as signed only when its signatures verify; a signature part alone is never enough.
 */

import type {
	DigitalSignatureReport,
	DigitalSignatureVerificationStatus,
	SignatureDetailStatus,
} from './types';

/**
 * - `unsigned`: no signature.
 * - `verified`: every signature verifies, covers the content, and its signer is trusted.
 * - `verified-untrusted`: every signature verifies and covers the content, but no signer is
 *   established as trusted (always the case in a browser, which has no trust store of its own).
 * - `invalid`: a signature does not verify, does not cover the content, or its certificate is
 *   revoked: the document may have changed since it was signed.
 * - `unsupported`: a signature uses something this verifier does not implement, or could not be
 *   read, so it was not verified either way.
 */
export type SignatureState =
	| 'unsigned'
	| 'verified'
	| 'verified-untrusted'
	| 'invalid'
	| 'unsupported';

/** Every state, in the order a summary lists them. */
export const SIGNATURE_STATES: readonly SignatureState[] = [
	'unsigned',
	'verified',
	'verified-untrusted',
	'invalid',
	'unsupported',
];

const STATE_OF_STATUS: Readonly<Record<DigitalSignatureVerificationStatus, SignatureState>> = {
	unsigned: 'unsigned',
	'verified-trusted': 'verified',
	'verified-untrusted': 'verified-untrusted',
	'timestamp-untrusted': 'verified-untrusted',
	'certificate-revoked': 'invalid',
	'digest-mismatch': 'invalid',
	'reference-missing': 'invalid',
	'coverage-incomplete': 'invalid',
	'signature-invalid': 'invalid',
	'timestamp-invalid': 'invalid',
	'present-not-verified': 'unsupported',
	'invalid-package': 'unsupported',
	error: 'unsupported',
};

/** The state of a verification report. */
export const signatureStateOf = (
	report: Pick<DigitalSignatureReport, 'verificationStatus'>,
): SignatureState => STATE_OF_STATUS[report.verificationStatus];

/** Whether a state presents the document as signed: its signatures verified. */
export const isVerifiedSignatureState = (state: SignatureState): boolean =>
	state === 'verified' || state === 'verified-untrusted';

const STATE_OF_DETAIL: Readonly<Record<SignatureDetailStatus, SignatureState>> = {
	verified: 'verified',
	'certificate-untrusted': 'verified-untrusted',
	'timestamp-untrusted': 'verified-untrusted',
	'certificate-revoked': 'invalid',
	'digest-mismatch': 'invalid',
	'reference-missing': 'invalid',
	'coverage-incomplete': 'invalid',
	'signature-invalid': 'invalid',
	'timestamp-invalid': 'invalid',
	'structural-only': 'unsupported',
};

/** One signature as a product shows it. */
export interface SignatureEntry {
	/** The signature part, for example `_xmlsignatures/sig1.xml`. */
	path: string;
	state: SignatureState;
	/** The signer certificate's subject (`CN=..., O=...`), when it could be read. */
	signer?: string | undefined;
	issuer?: string | undefined;
	/** ISO date the signer claims, read from signed content only. */
	signingTime?: string | undefined;
	/** ISO dates of the certificate's validity. */
	validFrom?: string | undefined;
	validTo?: string | undefined;
}

/** A package's signatures as a product shows them: the overall state and each signature. */
export interface SignatureSummary {
	state: SignatureState;
	signatures: SignatureEntry[];
}

/** The summary of a verification report. */
export function summarizeSignatures(report: DigitalSignatureReport): SignatureSummary {
	const details = report.details ?? [];
	return {
		state: signatureStateOf(report),
		signatures: details.length
			? details.map((detail) => ({
					path: detail.path,
					state: STATE_OF_DETAIL[detail.status],
					signer: detail.certificate?.subject,
					issuer: detail.certificate?.issuer,
					signingTime: detail.signingTime,
					validFrom: detail.certificate?.validFrom,
					validTo: detail.certificate?.validTo,
				}))
			: report.signaturePaths.map((path) => ({ path, state: signatureStateOf(report) })),
	};
}

/** The common name of a certificate subject (`CN=...`), else the whole subject. */
export function signerDisplayName(subject: string | undefined): string | undefined {
	if (!subject) return undefined;
	const match = /(?:^|,\s*|\n)CN=([^,\n]+)/u.exec(subject);
	return (match?.[1] ?? subject).trim();
}
