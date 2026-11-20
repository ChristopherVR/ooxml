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
		| 'content-type-mismatch'
		| 'missing-part'
		| 'unsupported-transform'
		| 'unsupported-algorithm'
		| 'insufficient-data'
		/** Not read: the signature reached the verifier's work limit before this reference. */
		| 'limit-exceeded';
	transformAlgorithms: string[];
}

/**
 * Why the verified references do not cover the signed content. The walk starts at the package
 * relationships and follows signed relationships only.
 */
export interface SignatureCoverageGap {
	kind:
		| 'unsigned-relationships-part'
		| 'unsigned-relationship'
		| 'unsigned-part'
		| 'no-signed-content';
	/** The part concerned: the relationships part, or the unsigned relationship target. */
	part: string;
	relationshipId?: string | undefined;
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
	| 'coverage-incomplete'
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
	/** Every reference from the verified `SignedInfo` and the Manifests it protects. */
	referenceChecks: SignatureReferenceCheck[];
	/** Package parts no verified reference covers, signature and content-types parts aside. */
	uncoveredParts: string[];
	/** Reasons the signed references do not cover the content; empty when they do. */
	coverageGaps: SignatureCoverageGap[];
	/** Why the signature XML was rejected before any cryptographic check. */
	structureError?: string | undefined;
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
	| 'coverage-incomplete'
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
	/** `SourceType` values of `RelationshipsGroupReference` transform parameters. */
	relationshipGroupTypes: string[];
	/** Set when the transform has parameters the verifier does not apply. */
	unsupportedParameters?: boolean | undefined;
}

/** A `Relationship` element of a relationships part. */
export interface PackageRelationship {
	id: string;
	type: string;
	target: string;
	mode: 'Internal' | 'External';
}

export interface ReferenceTransformResult {
	data: Uint8Array;
	unsupportedAlgorithms: string[];
	/** The relationships the relationship transform selected, when one ran. */
	selectedRelationships?: PackageRelationship[] | undefined;
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
