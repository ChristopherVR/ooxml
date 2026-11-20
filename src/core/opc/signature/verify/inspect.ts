/**
 * Verification of every digital signature in an OPC package, the same in a browser and in Node.
 *
 * A signature is verified in dependency order. Its structure is checked, its `SignatureValue` is
 * verified over `SignedInfo`, references are read from that verified text only, and coverage is
 * computed from the references that verified. Each step trusts only what the previous one proved.
 *
 * Whether the signer is trusted is a separate question with no answer in a browser, which has no
 * trust store of its own: without an {@link VerifySignatureOptions.evaluateSigner} every intact
 * signature is `verified-untrusted`. The Node verifier (`ooxml-core/pptx/signature-node`) passes
 * one that checks the configured trust roots and revocation.
 */

import { decodeXmlBytes } from '../../../xml/index';
import { openZip } from '../../zip-limits';
import { DIGITAL_SIGNATURE_ORIGIN_REL_TYPE } from '../constants';
import { detectDigitalSignatures } from '../detection';
import { computeDetailStatus, computeVerificationStatus } from '../inspection-status';
import { summarizeSignatures, type SignatureSummary } from '../state';
import type {
	CertificateRevocationStatus,
	DigitalSignatureReport,
	SignatureCertificateInfo,
	SignatureDetail,
	SignatureValidationPolicy,
	TimestampAuthorityStatus,
} from '../types';
import { evaluateSignatureCoverage } from './coverage';
import { parseXmlSignature } from './document';
import {
	checkSignedReferences,
	openSignaturePackage,
	resolvePartName,
	type SignaturePackage,
	type SignaturePackageSource,
} from './references';
import { verifySignedInfo } from './signed-info';
import { readCertificateBase64 } from './x509';

/** What a trust evaluation decides about the certificate that verified a signature. */
export interface SignerTrust {
	certificateTrustStatus: SignatureDetail['certificateTrustStatus'];
	certificateTrustError?: string | undefined;
	certificateRevocationStatus: CertificateRevocationStatus;
	certificateRevocationError?: string | undefined;
	certificateFingerprintSha256?: string | undefined;
	/** The signer's certificate details, when the evaluator reads them itself. */
	certificate?: SignatureCertificateInfo | undefined;
}

/** Options of {@link verifyPackageSignatures}. */
export interface VerifySignatureOptions {
	/**
	 * Decide whether the signer is trusted. `signerCertificate` verified the signature value;
	 * `certificates` are every certificate the signature carries (unsigned, in document order).
	 */
	evaluateSigner?: (signerCertificate: string, certificates: string[]) => Promise<SignerTrust>;
	/** Evaluate the signature's time-stamp. Without it the time-stamp is `not-checked`. */
	evaluateTimestamp?: (
		signatureXml: string,
	) => Promise<{ status: TimestampAuthorityStatus; error?: string | undefined }>;
	/** Revocation and time-stamp strictness. Defaults to requiring neither. */
	policy?: SignatureValidationPolicy;
}

const LENIENT_POLICY: SignatureValidationPolicy = {
	requireRevocationCheck: false,
	failOnRevocationUnknown: false,
	requireTimestamp: false,
};

const UNEVALUATED_SIGNER: SignerTrust = {
	certificateTrustStatus: 'untrusted',
	certificateTrustError: 'No trust roots are available to evaluate the signer.',
	certificateRevocationStatus: 'not-checked',
};

type DetailBase = Omit<SignatureDetail, 'status'>;

/** The report fields of a signature no reference could be read from. */
const unverifiedDetail = (path: string): DetailBase => ({
	path,
	referenceCount: 0,
	missingPartReferences: [],
	unsupportedTransforms: [],
	referenceChecks: [],
	uncoveredParts: [],
	coverageGaps: [],
	signatureValueStatus: 'invalid',
	certificateTrustStatus: 'not-checked',
	certificateRevocationStatus: 'not-checked',
	timestampAuthorityStatus: 'not-checked',
});

function certificateInfo(base64: string | undefined): SignatureCertificateInfo | undefined {
	const fields = base64 === undefined ? undefined : readCertificateBase64(base64);
	return (
		fields && {
			subject: fields.subject,
			issuer: fields.issuer,
			serialNumber: fields.serialNumber,
			validFrom: fields.validFrom,
			validTo: fields.validTo,
		}
	);
}

async function inspectSignature(
	path: string,
	signatureXml: string,
	pkg: SignaturePackage,
	options: VerifySignatureOptions,
): Promise<DetailBase> {
	const parsed = parseXmlSignature(signatureXml);
	if (!parsed.ok) {
		return { ...unverifiedDetail(path), structureError: parsed.error };
	}
	const { signature } = parsed;
	const timestamp = options.evaluateTimestamp
		? await options.evaluateTimestamp(signatureXml)
		: { status: 'not-checked' as const, error: undefined };
	const base: DetailBase = {
		...unverifiedDetail(path),
		signatureMethod: signature.signatureMethod,
		canonicalizationMethod: signature.canonicalizationMethod,
		timestampAuthorityStatus: timestamp.status,
		timestampAuthorityError: timestamp.error,
	};

	const verification = verifySignedInfo(signature);
	if (verification.status !== 'verified') {
		return {
			...base,
			signatureValueStatus: verification.status,
			certificate: certificateInfo(signature.certificates[0]),
		};
	}

	const { checks, coveredParts, signingTime } = await checkSignedReferences(
		signature,
		verification.canonicalSignedInfo,
		pkg,
	);
	const coverage = await evaluateSignatureCoverage(pkg, coveredParts);
	const trust = options.evaluateSigner
		? await options.evaluateSigner(verification.signerCertificate, signature.certificates)
		: UNEVALUATED_SIGNER;

	return {
		...base,
		signingTime,
		referenceCount: checks.length,
		missingPartReferences: checks
			.filter((check) => check.digestStatus === 'missing-part')
			.map((check) => check.uri),
		unsupportedTransforms: [
			...new Set(
				checks.flatMap((check) =>
					check.digestStatus === 'unsupported-transform' ? check.transformAlgorithms : [],
				),
			),
		],
		referenceChecks: checks,
		uncoveredParts: coverage.uncoveredParts,
		coverageGaps: coverage.coverageGaps,
		signatureValueStatus: 'verified',
		...trust,
		certificate: trust.certificate ?? certificateInfo(verification.signerCertificate),
	};
}

const isPackageSource = (input: unknown): input is SignaturePackageSource =>
	input instanceof Map || typeof (input as { file?: unknown }).file === 'function';

/**
 * Verify every digital signature of an OPC package (`.docx`, `.xlsx`, `.pptx`, `.vsdx` and their
 * macro and template variants). Never throws: a package that cannot be read reports
 * `invalid-package`.
 *
 * @param input - The package bytes, an opened JSZip archive or a map of part bytes by part name.
 */
export async function verifyPackageSignatures(
	input: SignaturePackageSource | Uint8Array | ArrayBuffer,
	options: VerifySignatureOptions = {},
): Promise<DigitalSignatureReport> {
	try {
		const source = isPackageSource(input) ? input : await openZip(input, { label: 'Package' });
		const pkg = await openSignaturePackage(source);
		const { signaturePaths } = detectDigitalSignatures([...pkg.partNames]);
		const rootRelationships = resolvePartName(pkg, '_rels/.rels');
		const rootBytes =
			rootRelationships === undefined ? undefined : await pkg.read(rootRelationships);
		const hasOriginRelationship = Boolean(
			rootBytes && decodeXmlBytes(rootBytes).includes(DIGITAL_SIGNATURE_ORIGIN_REL_TYPE),
		);
		if (signaturePaths.length === 0) {
			return {
				supported: true,
				hasSignature: false,
				signatureCount: 0,
				signaturePaths: [],
				verificationStatus: 'unsigned',
				hasOriginRelationship,
			};
		}
		const policy = options.policy ?? LENIENT_POLICY;
		const details: SignatureDetail[] = [];
		for (const path of signaturePaths) {
			const bytes = await pkg.read(path);
			let xml = '';
			try {
				xml = bytes ? decodeXmlBytes(bytes) : '';
			} catch {
				// Malformed text is rejected below as a structure error.
			}
			const detail = await inspectSignature(path, xml, pkg, options);
			details.push({ ...detail, status: computeDetailStatus(detail, policy) });
		}
		return {
			supported: true,
			hasSignature: true,
			signatureCount: signaturePaths.length,
			signaturePaths,
			verificationStatus: computeVerificationStatus(details),
			details,
			hasOriginRelationship,
		};
	} catch (error) {
		return {
			supported: true,
			hasSignature: false,
			signatureCount: 0,
			signaturePaths: [],
			verificationStatus: 'invalid-package',
			error: `Unable to inspect digital signature parts: ${String(error)}`,
		};
	}
}

/** {@link verifyPackageSignatures} as the summary a product shows. */
export async function verifyPackageSignatureSummary(
	input: SignaturePackageSource | Uint8Array | ArrayBuffer,
	options: VerifySignatureOptions = {},
): Promise<SignatureSummary> {
	return summarizeSignatures(await verifyPackageSignatures(input, options));
}

/**
 * The summary of a package that signature metadata (content types, the origin relationship) marks
 * as signed. Signatures stored where this verifier does not look are `unsupported`, never
 * `unsigned`, so a signed package is never shown as an unsigned one either.
 */
export async function verifySignedPackage(
	input: SignaturePackageSource | Uint8Array | ArrayBuffer,
	options: VerifySignatureOptions = {},
): Promise<SignatureSummary> {
	const summary = await verifyPackageSignatureSummary(input, options);
	return summary.state === 'unsigned' ? { state: 'unsupported', signatures: [] } : summary;
}
