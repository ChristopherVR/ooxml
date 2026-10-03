// OOXML digital signatures (OPC `_xmlsignatures/` parts, W3C XML-DSig): detection, the signature
// parts a writer must strip, XML-DSig parsing and digest checks over Web Crypto, and the report
// types the Node verifier (`ooxml-core/pptx/signature-node`) fills in. Format-neutral: every
// product signs the same package parts.
//
// Moved from the pptx area (`src/pptx/core/utils/signature-*.ts`; see PROVENANCE.md).
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
	type SignatureStatus,
} from './detection.js';
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
	PPTX_VIEWER_MANIFEST_NS,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
	XML_TRANSFORM_ENVELOPED_SIGNATURE,
} from './constants.js';
export type {
	CertificateRevocationStatus,
	DigitalSignatureReport,
	DigitalSignatureVerificationStatus,
	LoadedSigningMaterial,
	OfficeSignatureReference,
	ParsedReferenceTransform,
	ReferenceTransformResult,
	SignatureCertificateInfo as SignatureNodeCertificateInfo,
	SignatureDetail,
	SignatureDetailStatus,
	SignatureReferenceCheck,
	SignatureValidationPolicy,
	SignOptions,
	SignResult,
	TimestampAuthorityStatus,
} from './types.js';
export {
	escapeXmlAttr,
	escapeXmlText,
	extractAllTagText,
	extractFirstTagText,
	extractTagAttribute,
	isValidBase64,
} from './xml-utils.js';
export { normalizePartPath, resolveReferenceUriToPart } from './reference-utils.js';
export { computeDigestBase64 } from './digest.js';
export { computeDetailStatus, computeVerificationStatus } from './inspection-status.js';
