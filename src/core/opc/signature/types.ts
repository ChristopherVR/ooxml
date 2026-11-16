// Moved from src/core/pptx/core/utils/signature-types.ts (see PROVENANCE.md).
/**
 * Rich types for digital signature inspection, signing, and PKI validation.
 *
 * These are platform-agnostic - used by both browser-level detection
 * and Node-only full verification modules.
 */

// ── Certificate revocation & timestamp authority statuses ────────────────

export type CertificateRevocationStatus = 'good' | 'revoked' | 'unknown' | 'not-checked' | 'error';

export type TimestampAuthorityStatus =
	| 'valid'
	| 'invalid'
	| 'not-present'
	| 'not-checked'
	| 'error'
	| 'untrusted';

// ── Reference checking ──────────────────────────────────────────────────

export interface SignatureReferenceCheck {
	uri: string;
	resolvedPartPath?: string | undefined;
	existsInPackage: boolean;
	digestAlgorithm?: string | undefined;
	digestExpectedBase64?: string | undefined;
	digestActualBase64?: string | undefined;
	digestStatus:
		| 'verified'
		| 'mismatch'
		| 'missing-part'
		| 'unsupported-transform'
		| 'unsupported-algorithm'
		| 'insufficient-data';
	transformAlgorithms: string[];
}

// ── Certificate info ────────────────────────────────────────────────────

export interface SignatureCertificateInfo {
	subject?: string | undefined;
	issuer?: string | undefined;
	serialNumber?: string | undefined;
	validFrom?: string | undefined;
	validTo?: string | undefined;
}

// ── Signature detail ────────────────────────────────────────────────────

export type SignatureDetailStatus =
	| 'verified'
	| 'digest-mismatch'
	| 'reference-missing'
	| 'signature-invalid'
	| 'certificate-untrusted'
	| 'certificate-revoked'
	| 'timestamp-invalid'
	| 'timestamp-untrusted'
	| 'structural-only';

export interface SignatureDetail {
	path: string;
	signatureMethod?: string | undefined;
	canonicalizationMethod?: string | undefined;
	signingTime?: string | undefined;
	referenceCount: number;
	missingPartReferences: string[];
	unsupportedTransforms: string[];
	referenceChecks: SignatureReferenceCheck[];
	certificate?: SignatureCertificateInfo | undefined;
	signatureValueStatus: 'verified' | 'invalid' | 'not-checked';
	certificateTrustStatus: 'trusted' | 'untrusted' | 'not-checked';
	certificateTrustError?: string | undefined;
	certificateRevocationStatus: CertificateRevocationStatus;
	certificateRevocationError?: string | undefined;
	timestampAuthorityStatus: TimestampAuthorityStatus;
	timestampAuthorityError?: string | undefined;
	certificateFingerprintSha256?: string | undefined;
	status: SignatureDetailStatus;
}

// ── Report ──────────────────────────────────────────────────────────────

export type DigitalSignatureVerificationStatus =
	| 'unsigned'
	| 'verified-trusted'
	| 'verified-untrusted'
	| 'certificate-revoked'
	| 'digest-mismatch'
	| 'reference-missing'
	| 'signature-invalid'
	| 'timestamp-invalid'
	| 'timestamp-untrusted'
	| 'present-not-verified'
	| 'invalid-package'
	| 'error';

export interface DigitalSignatureReport {
	supported: boolean;
	hasSignature: boolean;
	signatureCount: number;
	signaturePaths: string[];
	verificationStatus: DigitalSignatureVerificationStatus;
	error?: string;
	details?: SignatureDetail[];
	hasOriginRelationship?: boolean;
}

// ── Signing ─────────────────────────────────────────────────────────────

export interface SignOptions {
	certificatePath: string;
	certificatePassword?: string;
}

export interface SignResult {
	success: boolean;
	signedData?: Uint8Array;
	report: DigitalSignatureReport;
	error?: string;
}

export interface LoadedSigningMaterial {
	privateKeyPem: string;
	certificatePem: string;
}

// ── Reference transforms ────────────────────────────────────────────────

export interface ParsedReferenceTransform {
	algorithm: string;
	relationshipReferenceIds: string[];
}

export interface ReferenceTransformResult {
	data: Uint8Array;
	unsupportedAlgorithms: string[];
}

// ── Validation policy ───────────────────────────────────────────────────

export interface SignatureValidationPolicy {
	requireRevocationCheck: boolean;
	failOnRevocationUnknown: boolean;
	requireTimestamp: boolean;
}

// ── Office signature reference ──────────────────────────────────────────

export interface OfficeSignatureReference {
	uri: string;
	digestMethod: string;
	digestValue: string;
}
