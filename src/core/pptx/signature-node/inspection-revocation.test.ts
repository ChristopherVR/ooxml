import JSZip from 'jszip';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import { inspectPptxDigitalSignatures } from './inspection';
import { evaluateCertificateRevocation } from './ocsp';

vi.mock('./ocsp', () => ({
	evaluateCertificateRevocation: vi.fn(async () => ({
		status: 'good',
		checkedOcspUrls: ['http://ocsp.example/'],
		checkedCrlUrls: [],
	})),
}));

async function signedDeck(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'_xmlsignatures/sig1.xml',
		'<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><KeyInfo><X509Data>' +
			'<X509Certificate>QUJD</X509Certificate></X509Data></KeyInfo></Signature>',
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
