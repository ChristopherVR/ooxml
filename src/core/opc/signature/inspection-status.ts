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

/**
 * Compute the overall status for an individual signature detail.
 *
 * `verified` and `certificate-untrusted` both mean the integrity checks passed: the signature
 * value verified, every signed reference verified, and those references cover the content.
 *
 * @param detail - A partial `SignatureDetail` with the fields needed for status computation.
 * @param policy - The validation policy controlling revocation/timestamp strictness.
 * @returns The computed status for this signature detail.
 */
export function computeDetailStatus(
	detail: Pick<
		SignatureDetail,
		| 'signatureValueStatus'
		| 'missingPartReferences'
		| 'referenceChecks'
		| 'coverageGaps'
		| 'certificateTrustStatus'
		| 'certificateRevocationStatus'
		| 'timestampAuthorityStatus'
	>,
	policy: SignatureValidationPolicy,
): SignatureDetailStatus {
	const hasDigestMismatch = detail.referenceChecks.some(
		(check) => check.digestStatus === 'mismatch' || check.digestStatus === 'content-type-mismatch',
	);
	const hasMissingPart = detail.missingPartReferences.length > 0;
	const everyReferenceVerified =
		detail.referenceChecks.length > 0 &&
		detail.referenceChecks.every((check) => check.digestStatus === 'verified');
	const revocationUnknown =
		detail.certificateRevocationStatus === 'unknown' ||
		detail.certificateRevocationStatus === 'not-checked' ||
		detail.certificateRevocationStatus === 'error';

	if (detail.signatureValueStatus === 'invalid') {
		return 'signature-invalid';
	}
	if (hasMissingPart) {
		return 'reference-missing';
	}
	if (hasDigestMismatch) {
		return 'digest-mismatch';
	}
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
	if (detail.signatureValueStatus !== 'verified' || !everyReferenceVerified) {
		return 'structural-only';
	}
	if (detail.coverageGaps.length > 0) {
		return 'coverage-incomplete';
	}
	if (detail.certificateTrustStatus === 'untrusted') {
		return 'certificate-untrusted';
	}
	return detail.certificateTrustStatus === 'trusted' ? 'verified' : 'structural-only';
}

/**
 * Compute the overall verification status from all signature details.
 *
 * @param details - Array of signature details from all signatures in the package.
 * @returns The overall verification status for the report.
 */
export function computeVerificationStatus(
	details: SignatureDetail[],
): DigitalSignatureReport['verificationStatus'] {
	const has = (status: SignatureDetailStatus) => details.some((detail) => detail.status === status);
	if (has('signature-invalid')) return 'signature-invalid';
	if (has('certificate-revoked')) return 'certificate-revoked';
	if (has('timestamp-invalid')) return 'timestamp-invalid';
	if (has('timestamp-untrusted')) return 'timestamp-untrusted';
	if (has('reference-missing')) return 'reference-missing';
	if (has('digest-mismatch')) return 'digest-mismatch';
	if (has('coverage-incomplete')) return 'coverage-incomplete';
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
