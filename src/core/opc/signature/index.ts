// OOXML digital signatures (OPC `_xmlsignatures/` parts, W3C XML-DSig): detection, the signature
// parts a writer must strip, XML-DSig parsing and digest checks over Web Crypto, and the report
// types the Node verifier (`ooxml-core/pptx/signature-node`) fills in. Format-neutral: every
// product signs the same package parts.
//
// Moved from the pptx area (`src/core/pptx/core/utils/signature-*.ts`; see PROVENANCE.md).
export {
	SIGNATURE_PART_PREFIX,
	detectDigitalSignatures,
	getSignaturePathsToStrip,
	isSignaturePart,
	parseSignatureXml,
	verifySignatureDigests,
	type ParsedSignature,
	type SignatureCertificateInfo,
	type SignatureDetectionResult,
	type SignatureReference,
	type SignatureDigestStatus,
	type SignatureStatus,
} from './detection';
export {
	DIGEST_ALGORITHM_TO_HASH,
	DIGEST_ALGORITHM_TO_WEB_CRYPTO,
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
	ENTERPRISE_FAIL_ON_REVOCATION_UNKNOWN_ENV,
	ENTERPRISE_REQUIRE_REVOCATION_ENV,
	ENTERPRISE_REQUIRE_TIMESTAMP_ENV,
	ENTERPRISE_TRUST_ROOTS_FILE_ENV,
	ENTERPRISE_TRUST_ROOTS_PEM_ENV,
	OPC_RELATIONSHIP_TRANSFORM,
	OPC_DIGITAL_SIGNATURE_NS,
	OPC_RELATIONSHIPS_NS,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
	XML_TRANSFORM_ENVELOPED_SIGNATURE,
} from './constants';
export type {
	CertificateRevocationStatus,
	DigitalSignatureReport,
	DigitalSignatureVerificationStatus,
	LoadedSigningMaterial,
	OfficeSignatureReference,
	PackageRelationship,
	ParsedReferenceTransform,
	ReferenceTransformResult,
	SignatureCertificateInfo as SignatureNodeCertificateInfo,
	SignatureDetail,
	SignatureCoverageGap,
	SignatureDetailStatus,
	SignatureReferenceCheck,
	SignatureValidationPolicy,
	SignOptions,
	SignResult,
	TimestampAuthorityStatus,
} from './types';
export {
	escapeXmlAttr,
	escapeXmlText,
	extractAllTagText,
	extractFirstTagText,
	extractTagAttribute,
	isValidBase64,
} from './xml-utils';
export {
	parseReferenceUri,
	resolveReferenceUriToPart,
	type ParsedReferenceUri,
} from './reference-utils';
export { computeDigestBase64 } from './digest';
export { computeDetailStatus, computeVerificationStatus } from './inspection-status';
export {
	inspectPackageSignatures,
	stripPackageSignatures,
	type PackageSignatures,
} from './package';
export {
	INCLUSIVE_C14N,
	INCLUSIVE_C14N_WITH_COMMENTS,
	EXCLUSIVE_C14N,
	EXCLUSIVE_C14N_WITH_COMMENTS,
} from './constants';
export {
	SIGNATURE_STATES,
	isVerifiedSignatureState,
	signatureStateOf,
	signerDisplayName,
	summarizeSignatures,
	type SignatureEntry,
	type SignatureState,
	type SignatureSummary,
} from './state';
export { canonicalize } from './verify/c14n';
export { evaluateSignatureCoverage, type SignatureCoverage } from './verify/coverage';
export { parseXmlSignature, type ParsedXmlSignature, type XmlSignature } from './verify/document';
export {
	verifyPackageSignatures,
	verifyPackageSignatureSummary,
	verifySignedPackage,
	type SignerTrust,
	type VerifySignatureOptions,
} from './verify/inspect';
export {
	checkSignedReferences,
	computeDigestBase64 as computeDigestBase64Sync,
	openSignaturePackage,
	type CoveredPart,
	type SignaturePackage,
	type SignaturePackageSource,
	type SignedReferenceResult,
} from './verify/references';
export { verifySignedInfo, type SignedInfoVerification } from './verify/signed-info';
export { applyReferenceTransforms, extractReferenceTransforms } from './verify/transforms';
export { readCertificate, readCertificateBase64, type CertificateFields } from './verify/x509';
