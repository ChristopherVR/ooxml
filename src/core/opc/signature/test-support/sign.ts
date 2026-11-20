/**
 * Signed packages for tests of every product, in each signature state. Signing uses the package
 * signer of `ooxml-core/pptx/signature-node` (format-neutral despite its name) with a certificate
 * generated for the test run, so no key or certificate is ever stored.
 */

import JSZip from 'jszip';

import { createTestParty, type TestParty } from '../../../pptx/test-support/signature-package';
import { signPptxWithCertificate } from '../../../pptx/signature-node/signing';

let root: TestParty | undefined;
let party: TestParty | undefined;

/** The root that issued {@link testSigner}: configure it as a trust root to make a signer trusted. */
export const testRoot = (): TestParty =>
	(root ??= createTestParty({ name: 'OOXML Test Root', ca: true }));

/** The certificate every signed test package uses, created once per test file. */
export const testSigner = (): TestParty =>
	(party ??= createTestParty({
		name: 'OOXML Test Signer',
		issuer: testRoot(),
		extensions: [{ name: 'extKeyUsage', emailProtection: true }],
	}));

/** `bytes` signed with {@link testSigner}: the signature verifies, the signer is not trusted. */
export async function signForTest(bytes: Uint8Array): Promise<Uint8Array> {
	const signer = testSigner();
	const pem = new TextEncoder().encode(`${signer.keyPem}\n${signer.certPem}`);
	const result = await signPptxWithCertificate(bytes, pem, { certificatePath: 'test.pem' });
	if (!result.success || !result.signedData) throw new Error(result.error ?? 'Signing failed.');
	return result.signedData;
}

async function rewrite(
	bytes: Uint8Array,
	change: (zip: JSZip) => Promise<void> | void,
): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	await change(zip);
	return zip.generateAsync({ type: 'uint8array' });
}

/** A signed package with one signed part changed after signing: its signature is invalid. */
export const tamperSignedPart = (signed: Uint8Array, part: string): Promise<Uint8Array> =>
	rewrite(signed, async (zip) => {
		const text = await zip.file(part)!.async('string');
		// A comment just inside the root element changes the bytes and nothing a reader models.
		const declaration = text.indexOf('?>');
		const rootEnd = text.indexOf('>', text.indexOf('<', declaration < 0 ? 0 : declaration + 2));
		zip.file(part, `${text.slice(0, rootEnd + 1)}<!-- changed -->${text.slice(rootEnd + 1)}`);
	});

/**
 * A signed package whose signature names a signature method the verifier does not implement, so
 * it can be verified neither way.
 */
export const withUnsupportedSignatureMethod = (signed: Uint8Array): Promise<Uint8Array> =>
	rewrite(signed, async (zip) => {
		const path = Object.keys(zip.files).find((name) => /^_xmlsignatures\/.+\.xml$/u.test(name))!;
		const xml = await zip.file(path)!.async('string');
		zip.file(
			path,
			xml.replace(
				/(<SignatureMethod Algorithm=")[^"]+"/u,
				'$1http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"',
			),
		);
	});
