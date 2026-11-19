// Moved from src/core/pptx/core/utils/signature-inspection-status.ts (see PROVENANCE.md).
/**
 * Pure status-computation logic for digital signature inspection.
 *
 * These functions are platform-agnostic - they accept data and policy
 * as parameters instead of reading from the environment.
 */

import type {
	DigitalSignatureReport,
	SignatureDetail,
	SignatureDetailStatus,
	SignatureValidationPolicy,
} from './types';

type SignatureIntegrityFields = Pick<
	SignatureDetail,
	'signatureValueStatus' | 'missingPartReferences' | 'referenceChecks' | 'coverageGaps'
>;

/** Why a signature does not protect the package's content, or `undefined` when it does. */
function integrityFailure(detail: SignatureIntegrityFields): SignatureDetailStatus | undefined {
	if (detail.signatureValueStatus === 'invalid') {
		return 'signature-invalid';
	}
	if (detail.missingPartReferences.length > 0) {
		return 'reference-missing';
	}
	const hasDigestMismatch = detail.referenceChecks.some(
		(check) => check.digestStatus === 'mismatch' || check.digestStatus === 'content-type-mismatch',
	);
	if (hasDigestMismatch) {
		return 'digest-mismatch';
	}
	const everyReferenceVerified =
		detail.referenceChecks.length > 0 &&
		detail.referenceChecks.every((check) => check.digestStatus === 'verified');
	if (detail.signatureValueStatus !== 'verified' || !everyReferenceVerified) {
		return 'structural-only';
	}
	return detail.coverageGaps.length > 0 ? 'coverage-incomplete' : undefined;
}

/**
 * Whether the signature value and every signed reference verified and those references cover the
 * content. Certificates, revocation, timestamps and policy play no part in this.
 */
export const hasVerifiedIntegrity = (detail: SignatureIntegrityFields): boolean =>
	integrityFailure(detail) === undefined;

/**
 * Compute the overall status for an individual signature detail.
 *
 * Integrity is decided first, so no certificate, timestamp or policy outcome can replace an
 * integrity failure. Every status those checks produce therefore means integrity held.
 *
 * @param detail - A partial `SignatureDetail` with the fields needed for status computation.
 * @param policy - The validation policy controlling revocation/timestamp strictness.
 * @returns The computed status for this signature detail.
 */
export function computeDetailStatus(
	detail: SignatureIntegrityFields &
		Pick<
			SignatureDetail,
			'certificateTrustStatus' | 'certificateRevocationStatus' | 'timestampAuthorityStatus'
		>,
	policy: SignatureValidationPolicy,
): SignatureDetailStatus {
	const failure = integrityFailure(detail);
	if (failure) {
		return failure;
	}
	const revocationUnknown =
		detail.certificateRevocationStatus === 'unknown' ||
		detail.certificateRevocationStatus === 'not-checked' ||
		detail.certificateRevocationStatus === 'error';

	if (detail.certificateRevocationStatus === 'revoked') {
		return 'certificate-revoked';
	}
	if (
		policy.requireRevocationCheck &&
		(policy.failOnRevocationUnknown
			? revocationUnknown
			: detail.certificateRevocationStatus !== 'good')
	) {
		return 'certificate-untrusted';
	}
	if (detail.timestampAuthorityStatus === 'invalid') {
		return 'timestamp-invalid';
	}
	if (policy.requireTimestamp && detail.timestampAuthorityStatus !== 'valid') {
		return 'timestamp-untrusted';
	}
	if (detail.timestampAuthorityStatus === 'untrusted') {
		return 'timestamp-untrusted';
	}
	if (detail.certificateTrustStatus === 'untrusted') {
		return 'certificate-untrusted';
	}
	return detail.certificateTrustStatus === 'trusted' ? 'verified' : 'structural-only';
}

/** Report status of each failing detail status, most serious first: integrity before the rest. */
const REPORT_STATUS_PRECEDENCE: readonly (readonly [
	SignatureDetailStatus,
	DigitalSignatureReport['verificationStatus'],
])[] = [
	['signature-invalid', 'signature-invalid'],
	['reference-missing', 'reference-missing'],
	['digest-mismatch', 'digest-mismatch'],
	['structural-only', 'present-not-verified'],
	['coverage-incomplete', 'coverage-incomplete'],
	['certificate-revoked', 'certificate-revoked'],
	['timestamp-invalid', 'timestamp-invalid'],
	['timestamp-untrusted', 'timestamp-untrusted'],
];

/**
 * Compute the overall verification status from all signature details.
 *
 * @param details - Array of signature details from all signatures in the package.
 * @returns The overall verification status for the report.
 */
export function computeVerificationStatus(
	details: SignatureDetail[],
): DigitalSignatureReport['verificationStatus'] {
	for (const [status, reportStatus] of REPORT_STATUS_PRECEDENCE) {
		if (details.some((detail) => detail.status === status)) return reportStatus;
	}
	if (details.length > 0 && details.every((detail) => detail.status === 'verified')) {
		return 'verified-trusted';
	}
	const integrityVerified =
		details.length > 0 &&
		details.every(
			(detail) => detail.status === 'verified' || detail.status === 'certificate-untrusted',
		);
	return integrityVerified ? 'verified-untrusted' : 'present-not-verified';
}
