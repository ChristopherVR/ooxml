import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	createManifestReferences,
	createOfficeSignatureXml,
	createTestPackage,
	withSignature,
} from '../../../pptx/test-support/signature-package';
import {
	rsaKeyCertificate,
	signerCertificateOf,
	signForTest,
	testSigner,
	withKeyInfoCertificates,
} from '../test-support/sign';
import { verifyPackageSignatures } from './inspect';

const ones = (bits: number): bigint => (1n << BigInt(bits)) - 1n;

const signedPackage = async () =>
	signForTest(await createTestPackage().generateAsync({ type: 'uint8array' }));

async function timedVerify(bytes: Uint8Array) {
	const started = performance.now();
	const report = await verifyPackageSignatures(bytes);
	return { report, ms: performance.now() - started, detail: report.details?.[0] };
}

describe('verifier work limits', () => {
	it('does not compute with a 32768-bit key whose exponent is as large', async () => {
		const huge = rsaKeyCertificate(ones(32_768), ones(32_768));
		const bytes = await withKeyInfoCertificates(await signedPackage(), [huge], 32_768 / 8);
		const { report, ms, detail } = await timedVerify(bytes);
		expect(report.verificationStatus).toBe('present-not-verified');
		expect(detail?.signatureValueStatus).toBe('not-checked');
		expect(ms).toBeLessThan(100);
	});

	it.each([
		['a modulus under 1024 bits', ones(1016), 65_537n],
		['a modulus over 16384 bits', ones(16_392), 65_537n],
		['an exponent over 2^32', ones(2048), 2n ** 32n + 1n],
	])('reports a signature as not checked for a key with %s', async (_label, modulus, exponent) => {
		const bytes = await withKeyInfoCertificates(await signedPackage(), [
			rsaKeyCertificate(modulus, exponent),
		]);
		const { detail } = await timedVerify(bytes);
		expect(detail?.signatureValueStatus).toBe('not-checked');
		expect(detail?.status).toBe('structural-only');
	});

	it('tries only the first four certificates, so many large keys cost little', async () => {
		const signed = await signedPackage();
		const large = rsaKeyCertificate(ones(16_384), 2n ** 32n - 1n);
		const certificates = [
			...Array.from({ length: 256 }, () => large),
			await signerCertificateOf(signed),
		];
		const { report, ms, detail } = await timedVerify(
			await withKeyInfoCertificates(signed, certificates, 16_384 / 8),
		);
		expect(report.verificationStatus).toBe('present-not-verified');
		expect(detail?.signatureValueStatus).toBe('not-checked');
		expect(ms).toBeLessThan(500);
	});

	it('still finds the signer among the first four certificates', async () => {
		const signed = await signedPackage();
		const other = rsaKeyCertificate(ones(2048), 65_537n);
		const bytes = await withKeyInfoCertificates(signed, [
			other,
			other,
			other,
			await signerCertificateOf(signed),
		]);
		const { report } = await timedVerify(bytes);
		expect(report.verificationStatus).toBe('verified-untrusted');
	});

	it('reports a signature with more than 10000 references as not checked', async () => {
		const zip = createTestPackage();
		const references = await createManifestReferences(zip);
		const repeated = Array.from({ length: 10_001 }, (_, index) => references[index % 4]!);
		const xml = createOfficeSignatureXml({ signer: testSigner(), manifestReferences: repeated });
		const { report, detail } = await timedVerify(await withSignature(zip, xml));
		expect(report.verificationStatus).toBe('present-not-verified');
		expect(detail?.referenceChecks.at(-1)?.digestStatus).toBe('limit-exceeded');
	});

	it('does not read a signature part over 16 MB', async () => {
		const zip = await JSZip.loadAsync(await signedPackage());
		const path = Object.keys(zip.files).find((name) => name.startsWith('_xmlsignatures/sig'))!;
		const xml = await zip.file(path)!.async('string');
		zip.file(
			path,
			xml.replace('</Signature>', `<!--${' '.repeat(16 * 1024 * 1024)}--></Signature>`),
		);
		const { report, detail } = await timedVerify(await zip.generateAsync({ type: 'uint8array' }));
		expect(report.verificationStatus).toBe('present-not-verified');
		expect(detail?.structureError).toMatch(/too large/u);
	});
});
