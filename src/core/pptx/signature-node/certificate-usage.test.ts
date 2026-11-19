import crypto from 'node:crypto';

import { beforeAll, describe, expect, it } from 'vitest';

import { createTestParty, type TestParty } from '../test-support/signature-package';
import { validateCertificateChain } from './certificate-utils';
import { isDocumentSigningCertificate } from './certificate-usage';

const x509 = (party: TestParty) => new crypto.X509Certificate(party.certPem);
const keyUsage = (flags: object) => ({ name: 'keyUsage', ...flags });
const extendedKeyUsage = (flags: object) => ({ name: 'extKeyUsage', ...flags });

describe('isDocumentSigningCertificate', () => {
	let root: TestParty;
	const leaf = (extensions: object[]) =>
		createTestParty({ name: 'Leaf', issuer: root, extensions });

	beforeAll(() => {
		root = createTestParty({ name: 'Usage Root', ca: true });
	});

	it('accepts a certificate that states no usage', () => {
		expect(isDocumentSigningCertificate(x509(leaf([])))).toBe(true);
	});

	it.each([
		['digitalSignature', { digitalSignature: true }],
		['nonRepudiation', { nonRepudiation: true }],
	])('accepts the %s key usage', (_name, flags) => {
		expect(isDocumentSigningCertificate(x509(leaf([keyUsage(flags)])))).toBe(true);
	});

	it('rejects a key usage without a signing bit', () => {
		const party = leaf([keyUsage({ keyEncipherment: true })]);
		expect(isDocumentSigningCertificate(x509(party))).toBe(false);
	});

	it('accepts the email protection extended key usage', () => {
		const party = leaf([extendedKeyUsage({ emailProtection: true })]);
		expect(isDocumentSigningCertificate(x509(party))).toBe(true);
	});

	it('accepts the document signing extended key usages by OID', () => {
		const party = leaf([
			extendedKeyUsage({ '1.3.6.1.5.5.7.3.36': true, '1.3.6.1.4.1.311.10.3.12': true }),
		]);
		expect(isDocumentSigningCertificate(x509(party))).toBe(true);
	});

	it.each([
		['server authentication', { serverAuth: true }],
		['code signing', { codeSigning: true }],
	])('rejects a %s certificate', (_name, flags) => {
		const party = leaf([extendedKeyUsage(flags)]);
		expect(isDocumentSigningCertificate(x509(party))).toBe(false);
	});

	it('leaves the certificate untrusted even when it chains to a configured root', () => {
		const party = leaf([extendedKeyUsage({ serverAuth: true })]);
		const result = validateCertificateChain([party.certBase64], [root.certPem]);

		expect(result.status).toBe('untrusted');
		expect(result.error).toMatch(/document signing/);
	});
});
