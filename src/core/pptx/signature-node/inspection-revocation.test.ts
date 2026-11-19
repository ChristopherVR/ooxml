import { beforeEach, describe, expect, it, vi } from 'vitest';

import {
	createManifestReferences,
	createOfficeSignatureXml,
	createTestPackage,
	createTestParty,
	withSignature,
} from '../test-support/signature-package';
import { inspectPptxDigitalSignatures } from './inspection';
import { evaluateCertificateRevocation } from './ocsp';

vi.mock('./ocsp', () => ({
	evaluateCertificateRevocation: vi.fn(async () => ({
		status: 'good',
		checkedOcspUrls: ['http://ocsp.example/'],
		checkedCrlUrls: [],
	})),
}));

/** A deck whose signature value verifies, so inspection goes on to the revocation check. */
async function signedDeck(): Promise<Uint8Array> {
	const zip = createTestPackage();
	const signer = createTestParty({ name: 'Signer' });
	await withSignature(
		zip,
		createOfficeSignatureXml({ signer, manifestReferences: await createManifestReferences(zip) }),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('revocation checking during signature inspection', () => {
	beforeEach(() => vi.mocked(evaluateCertificateRevocation).mockClear());

	it('asks the responder by default', async () => {
		const report = await inspectPptxDigitalSignatures(await signedDeck());
		expect(evaluateCertificateRevocation).toHaveBeenCalledTimes(1);
		expect(report.details?.[0]?.certificateRevocationStatus).toBe('good');
	});

	it('makes no request with checkRevocation false', async () => {
		const report = await inspectPptxDigitalSignatures(await signedDeck(), {
			checkRevocation: false,
		});
		expect(evaluateCertificateRevocation).not.toHaveBeenCalled();
		expect(report.details?.[0]?.certificateRevocationStatus).toBe('not-checked');
	});
});
