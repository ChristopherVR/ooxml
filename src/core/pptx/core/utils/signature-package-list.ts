import { verifyPackageSignatureSummary } from '../../../opc/signature/verify/inspect';
import type { SignatureSummary } from '../../../opc/signature/state';
import { openPptxZip } from './zip-open';

/**
 * Verify the digital signatures of a presentation's bytes with the platform-neutral verifier of
 * `ooxml-core/opc`: each signature value, the parts it covers and whether they cover the content.
 * No signer is trusted here (a browser has no trust store), so an intact signature is
 * `verified-untrusted`. Never throws: an unreadable package is `unsupported`, an unsigned one
 * `unsigned`.
 */
export async function inspectSignaturesFromBuffer(
	buffer: ArrayBuffer | Uint8Array,
): Promise<SignatureSummary> {
	let zip;
	try {
		zip = await openPptxZip(buffer);
	} catch {
		return { state: 'unsupported', signatures: [] };
	}
	return verifyPackageSignatureSummary(zip);
}
