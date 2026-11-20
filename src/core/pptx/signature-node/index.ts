/**
 * Node-only digital signature subpath export.
 *
 * Usage: `import { ... } from 'pptx-viewer-core/signature-node'`
 *
 * Re-exports platform-agnostic types + Node-only implementations.
 */

// ── Platform-agnostic re-exports (types, constants, pure utils) ─────────

export type {
	CertificateRevocationStatus,
	TimestampAuthorityStatus,
	SignatureReferenceCheck,
	SignatureCoverageGap,
	PackageRelationship,
	SignatureCertificateInfo,
	SignatureDetailStatus,
	SignatureDetail,
	DigitalSignatureVerificationStatus,
	DigitalSignatureReport,
	SignOptions,
	SignResult,
	LoadedSigningMaterial,
	ParsedReferenceTransform,
	ReferenceTransformResult,
	SignatureValidationPolicy,
	OfficeSignatureReference,
} from '../core/utils/signature-types';

export {
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
	OPC_DIGITAL_SIGNATURE_NS,
	OPC_RELATIONSHIPS_NS,
	XMLDSIG_NS,
	OPC_RELATIONSHIP_TRANSFORM,
	XML_TRANSFORM_ENVELOPED_SIGNATURE,
	SUPPORTED_XML_CANON_TRANSFORMS,
	ENTERPRISE_TRUST_ROOTS_FILE_ENV,
	ENTERPRISE_TRUST_ROOTS_PEM_ENV,
	ENTERPRISE_REQUIRE_REVOCATION_ENV,
	ENTERPRISE_FAIL_ON_REVOCATION_UNKNOWN_ENV,
	ENTERPRISE_REQUIRE_TIMESTAMP_ENV,
	DIGEST_ALGORITHM_TO_HASH,
	DIGEST_ALGORITHM_TO_WEB_CRYPTO,
} from '../core/utils/signature-constants';

export {
	escapeXmlAttr,
	escapeXmlText,
	isValidBase64,
	extractTagAttribute,
	extractFirstTagText,
	extractAllTagText,
} from '../core/utils/signature-xml-utils';

export {
	parseReferenceUri,
	resolveReferenceUriToPart,
	type ParsedReferenceUri,
} from '../core/utils/signature-reference-utils';

export { computeDigestBase64 as computeDigestBase64WebCrypto } from '../core/utils/signature-digest';

export {
	computeDetailStatus,
	computeVerificationStatus,
} from '../core/utils/signature-inspection-status';

// ── Platform-neutral verifier (also exported by `ooxml-core/opc`) ─────

export {
	canonicalize,
	checkSignedReferences,
	evaluateSignatureCoverage,
	extractReferenceTransforms,
	applyReferenceTransforms,
	openSignaturePackage,
	parseXmlSignature,
	verifyPackageSignatures,
	verifySignedInfo,
	type CoveredPart,
	type ParsedXmlSignature,
	type SignatureCoverage,
	type SignaturePackage,
	type SignaturePackageSource,
	type SignedInfoVerification,
	type SignedReferenceResult,
	type SignerTrust,
	type VerifySignatureOptions,
	type XmlSignature,
} from '../../opc/signature/index';
export { parseRelationshipsPart } from '../../opc/relationships';
export { computeDigestBase64 } from '../../opc/signature/verify/references';

// ── Node-only exports ───────────────────────────────────────────────────

export {
	certificateInfoFromBase64,
	orderCertificateChain,
	validateCertificateChain,
	loadSigningMaterialFromBuffer,
	pemCertificateToBase64,
} from './certificate-utils';

export { isDocumentSigningCertificate } from './certificate-usage';

export {
	certPemFromBase64,
	certFingerprintSha256,
	evaluateTimestampAuthority,
} from './pki-validation';

export {
	extractOcspUrls,
	buildOcspRequestDer,
	parseOcspResponseStatus,
	evaluateCertificateRevocation,
} from './ocsp';

export {
	extractPemCertificatesFromText,
	loadEnterpriseTrustRoots,
	getSignatureValidationPolicy,
} from './environment';

export { signPptxWithCertificate } from './signing';

export { inspectPptxDigitalSignatures, type InspectPptxSignatureOptions } from './inspection';
