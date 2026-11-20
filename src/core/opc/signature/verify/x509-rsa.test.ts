import crypto from 'node:crypto';

import { describe, expect, it } from 'vitest';

import { testRoot, testSigner } from '../test-support/sign';
import { verifyRsaPkcs1 } from './rsa';
import { readCertificateBase64 } from './x509';

describe('readCertificateBase64', () => {
	it('reads the names, validity and RSA key of a certificate', () => {
		const signer = testSigner();
		const fields = readCertificateBase64(signer.certBase64)!;
		expect(fields.subject).toBe('CN=OOXML Test Signer');
		expect(fields.issuer).toBe('CN=OOXML Test Root');
		expect(new Date(fields.validTo!).getTime()).toBeGreaterThan(Date.now());
		const jwk = crypto.createPublicKey(signer.certPem).export({ format: 'jwk' });
		expect(fields.rsaPublicKey?.modulus).toBe(
			BigInt(`0x${Buffer.from(jwk.n!, 'base64url').toString('hex')}`),
		);
		expect(fields.rsaPublicKey?.exponent).toBe(65537n);
	});

	it('returns nothing for bytes that are not a certificate', () => {
		expect(readCertificateBase64('AAAA')).toBeUndefined();
		expect(readCertificateBase64('not base64!')).toBeUndefined();
	});
});

describe('verifyRsaPkcs1', () => {
	const message = new TextEncoder().encode('<SignedInfo/>');

	it.each(['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512'])(
		'accepts a %s signature Node made',
		(hash) => {
			const signer = testSigner();
			const key = readCertificateBase64(signer.certBase64)!.rsaPublicKey!;
			const signature = crypto.sign(hash.replace('-', '').toLowerCase(), message, signer.keyPem);
			expect(verifyRsaPkcs1(key, hash, message, new Uint8Array(signature))).toBe(true);
		},
	);

	it('rejects a changed message, a changed signature and another key', () => {
		const signer = testSigner();
		const key = readCertificateBase64(signer.certBase64)!.rsaPublicKey!;
		const signature = new Uint8Array(crypto.sign('sha256', message, signer.keyPem));
		expect(verifyRsaPkcs1(key, 'SHA-256', new TextEncoder().encode('<Other/>'), signature)).toBe(
			false,
		);
		const flipped = signature.slice();
		flipped[flipped.length - 1]! ^= 1;
		expect(verifyRsaPkcs1(key, 'SHA-256', message, flipped)).toBe(false);
		const other = readCertificateBase64(testRoot().certBase64)!.rsaPublicKey!;
		expect(verifyRsaPkcs1(other, 'SHA-256', message, signature)).toBe(false);
		expect(verifyRsaPkcs1(key, 'MD5', message, signature)).toBe(false);
	});
});
