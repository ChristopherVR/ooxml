/**
 * Full PPTX digital signature inspection (Node-only).
 *
 * A signature is verified in dependency order. Its structure is checked, its `SignatureValue` is
 * verified over `SignedInfo`, references are read from that verified text only, and coverage is
 * computed from the references that verified. Each step trusts only what the previous one proved.
 */

import type JSZip from 'jszip';

import { openPptxZip } from '../core/utils/zip-open';

import { detectDigitalSignatures } from '../../opc/signature/detection';
import { DIGITAL_SIGNATURE_ORIGIN_REL_TYPE } from '../core/utils/signature-constants';
import {
	computeDetailStatus,
	computeVerificationStatus,
} from '../core/utils/signature-inspection-status';
import type {
	DigitalSignatureReport,
	SignatureDetail,
	SignatureValidationPolicy,
} from '../core/utils/signature-types';
import {
	certificateInfoFromBase64,
	orderCertificateChain,
	validateCertificateChain,
} from './certificate-utils';
import { getSignatureValidationPolicy, loadEnterpriseTrustRoots } from './environment';
import { evaluateCertificateRevocation } from './ocsp';
import {
	certFingerprintSha256,
	certPemFromBase64,
	evaluateTimestampAuthority,
} from './pki-validation';
import { checkSignedReferences, openSignaturePackage } from './reference-checks';
import type { SignaturePackage } from './reference-checks';
import { evaluateSignatureCoverage } from './signature-coverage';
import { parseXmlSignature } from './signature-document';
import { verifySignedInfo } from './signature-value';

type DetailBase = Omit<SignatureDetail, 'status'>;

const NOT_CHECKED_REVOCATION: Awaited<ReturnType<typeof evaluateCertificateRevocation>> = {
	status: 'not-checked',
	checkedOcspUrls: [],
	checkedCrlUrls: [],
};

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

async function inspectSignature(
	path: string,
	signatureXml: string,
	pkg: SignaturePackage,
	trustRootsPem: string[],
	checkRevocation: boolean,
): Promise<DetailBase> {
	const parsed = parseXmlSignature(signatureXml);
	if (!parsed.ok) {
		return { ...unverifiedDetail(path), structureError: parsed.error };
	}
	const { signature } = parsed;
	const timestamp = await evaluateTimestampAuthority(signatureXml);
	const base: DetailBase = {
		...unverifiedDetail(path),
		signatureMethod: signature.signatureMethod,
		canonicalizationMethod: signature.canonicalizationMethod,
		timestampAuthorityStatus: timestamp.status,
		timestampAuthorityError: timestamp.error,
	};

	const verification = verifySignedInfo(signature);
	if (verification.status !== 'verified') {
		const [firstCertificate] = signature.certificates;
		return {
			...base,
			signatureValueStatus: verification.status,
			certificate:
				firstCertificate !== undefined ? certificateInfoFromBase64(firstCertificate) : undefined,
		};
	}

	const { checks, coveredParts, signingTime } = await checkSignedReferences(
		signature,
		verification.canonicalSignedInfo,
		pkg,
	);
	const coverage = await evaluateSignatureCoverage(pkg, coveredParts);

	const chain = orderCertificateChain(verification.signerCertificate, signature.certificates);
	const trust = validateCertificateChain(chain, trustRootsPem);
	const [leafCert, issuerCert] = chain;
	const leafCertPem = leafCert !== undefined ? certPemFromBase64(leafCert) : undefined;
	const issuerCertPem = issuerCert !== undefined ? certPemFromBase64(issuerCert) : leafCertPem;
	const revocation =
		leafCertPem && checkRevocation
			? await evaluateCertificateRevocation(leafCertPem, issuerCertPem)
			: NOT_CHECKED_REVOCATION;

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
		certificate: leafCert !== undefined ? certificateInfoFromBase64(leafCert) : undefined,
		signatureValueStatus: 'verified',
		certificateTrustStatus: trust.status,
		certificateTrustError: trust.error,
		certificateRevocationStatus: revocation.status,
		certificateRevocationError: revocation.error,
		certificateFingerprintSha256: leafCertPem ? certFingerprintSha256(leafCertPem) : undefined,
	};
}

async function inspectSignatures(
	zip: JSZip,
	signaturePaths: string[],
	policy: SignatureValidationPolicy,
	checkRevocation: boolean,
): Promise<SignatureDetail[]> {
	const pkg = await openSignaturePackage(zip);
	const trustRootsPem = await loadEnterpriseTrustRoots();
	const details: SignatureDetail[] = [];
	for (const path of signaturePaths) {
		const signatureXml = (await zip.file(path)?.async('string')) ?? '';
		const detail = await inspectSignature(path, signatureXml, pkg, trustRootsPem, checkRevocation);
		details.push({ ...detail, status: computeDetailStatus(detail, policy) });
	}
	return details;
}

/** Options of {@link inspectPptxDigitalSignatures}. */
export interface InspectPptxSignatureOptions {
	/**
	 * Ask the OCSP responder named in each signing certificate for its revocation status. Default
	 * true; with false no request is made and the status is `not-checked`.
	 */
	checkRevocation?: boolean;
}

/**
 * Inspect all digital signatures in a PPTX package.
 *
 * `verified-trusted` means every signature value verified, every signed reference verified, the
 * signed references cover the document content, and the signer chains to a configured trust root
 * with a certificate that permits document signing. Trust roots come only from
 * `PPTX_VIEWER_TRUST_ROOTS_PEM` and `PPTX_VIEWER_TRUST_ROOTS_FILE`; with none configured, an intact
 * signature reports `verified-untrusted`. Parts outside the signature are listed in each detail's
 * `uncoveredParts`.
 *
 * @param data - The raw PPTX file bytes.
 * @param options - Set `checkRevocation: false` to make no network request.
 * @returns A comprehensive digital signature report.
 */
export async function inspectPptxDigitalSignatures(
	data: Uint8Array,
	options: InspectPptxSignatureOptions = {},
): Promise<DigitalSignatureReport> {
	try {
		const zip = await openPptxZip(Buffer.from(data));
		const { signaturePaths } = detectDigitalSignatures(
			Object.keys(zip.files).filter((name) => !zip.files[name]?.dir),
		);
		const rootRelsXml = await zip.file('_rels/.rels')?.async('string');
		const hasOriginRelationship = Boolean(rootRelsXml?.includes(DIGITAL_SIGNATURE_ORIGIN_REL_TYPE));
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

		const details = await inspectSignatures(
			zip,
			signaturePaths,
			getSignatureValidationPolicy(),
			options.checkRevocation !== false,
		);
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
