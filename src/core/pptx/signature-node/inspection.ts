/**
 * Full digital signature inspection (Node-only): the platform-neutral verifier of
 * `ooxml-core/opc` decides integrity, and this module adds what only Node can: the configured trust
 * roots, the certificate's revocation status and the time-stamp authority.
 */

import { openPptxZip } from '../core/utils/zip-open';

import type { DigitalSignatureReport } from '../../opc/signature/types';
import { verifyPackageSignatures, type SignerTrust } from '../../opc/signature/verify/inspect';
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

async function evaluateSigner(
	signerCertificate: string,
	certificates: string[],
	trustRootsPem: string[],
	checkRevocation: boolean,
): Promise<SignerTrust> {
	const chain = orderCertificateChain(signerCertificate, certificates);
	const trust = validateCertificateChain(chain, trustRootsPem);
	const [leafCert, issuerCert] = chain;
	const leafCertPem = leafCert !== undefined ? certPemFromBase64(leafCert) : undefined;
	const issuerCertPem = issuerCert !== undefined ? certPemFromBase64(issuerCert) : leafCertPem;
	const revocation =
		leafCertPem && checkRevocation
			? await evaluateCertificateRevocation(leafCertPem, issuerCertPem)
			: { status: 'not-checked' as const, error: undefined };
	return {
		certificate: leafCert !== undefined ? certificateInfoFromBase64(leafCert) : undefined,
		certificateTrustStatus: trust.status,
		certificateTrustError: trust.error,
		certificateRevocationStatus: revocation.status,
		certificateRevocationError: revocation.error,
		certificateFingerprintSha256: leafCertPem ? certFingerprintSha256(leafCertPem) : undefined,
	};
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
 * Inspect all digital signatures in an OPC package (any Office format, despite the name).
 *
 * `verified-trusted` means every signature value verified, every signed reference verified, the
 * signed references cover the document content, and the signer chains to a configured trust root
 * with a certificate that permits document signing. Trust roots come only from
 * `PPTX_VIEWER_TRUST_ROOTS_PEM` and `PPTX_VIEWER_TRUST_ROOTS_FILE`; with none configured, an intact
 * signature reports `verified-untrusted`. Parts outside the signature are listed in each detail's
 * `uncoveredParts`.
 *
 * @param data - The raw package bytes.
 * @param options - Set `checkRevocation: false` to make no network request.
 * @returns A comprehensive digital signature report.
 */
export async function inspectPptxDigitalSignatures(
	data: Uint8Array,
	options: InspectPptxSignatureOptions = {},
): Promise<DigitalSignatureReport> {
	let zip;
	try {
		zip = await openPptxZip(data);
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
	const trustRootsPem = await loadEnterpriseTrustRoots();
	const checkRevocation = options.checkRevocation !== false;
	return verifyPackageSignatures(zip, {
		policy: getSignatureValidationPolicy(),
		evaluateTimestamp: evaluateTimestampAuthority,
		evaluateSigner: (signer, certificates) =>
			evaluateSigner(signer, certificates, trustRootsPem, checkRevocation),
	});
}
