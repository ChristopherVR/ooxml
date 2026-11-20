/**
 * Signed packages for tests of every product, in each signature state. Signing uses the package
 * signer of `ooxml-core/pptx/signature-node` (format-neutral despite its name) with a certificate
 * generated for the test run, so no key or certificate is ever stored.
 */

import JSZip from 'jszip';
import forge from 'node-forge';

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

const signaturePartOf = (zip: JSZip): string =>
	Object.keys(zip.files).find((name) => /^_xmlsignatures\/.+\.xml$/u.test(name))!;

/** A signed package with its signature part's text rewritten by `change`. */
const editSignature = (signed: Uint8Array, change: (xml: string) => string): Promise<Uint8Array> =>
	rewrite(signed, async (zip) => {
		const path = signaturePartOf(zip);
		zip.file(path, change(await zip.file(path)!.async('string')));
	});

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
	editSignature(signed, (xml) =>
		xml.replace(
			/(<SignatureMethod Algorithm=")[^"]+"/u,
			'$1http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"',
		),
	);

/**
 * A Base64 DER certificate, issued by {@link testRoot}, for an RSA public key of any size and
 * exponent. No private key exists for it, so nothing it carries can verify.
 */
export function rsaKeyCertificate(modulus: bigint, exponent: bigint): string {
	const integer = (value: bigint) => new forge.jsbn.BigInteger(value.toString(16), 16);
	const cert = forge.pki.createCertificate();
	cert.publicKey = forge.pki.setRsaPublicKey(integer(modulus), integer(exponent));
	cert.serialNumber = '01';
	cert.validity.notBefore = new Date(Date.now() - 24 * 3600 * 1000);
	cert.validity.notAfter = new Date(Date.now() + 24 * 3600 * 1000);
	cert.setSubject([{ name: 'commonName', value: 'OOXML Test Key' }]);
	cert.setIssuer(testRoot().cert.subject.attributes);
	cert.sign(testRoot().key, forge.md.sha256.create());
	return forge.util.encode64(forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes());
}

/**
 * A signed package whose `KeyInfo` carries `certificates` (Base64 DER, in this order) in place of
 * its own. `KeyInfo` is not signed, so only the keys the verifier tries change. With
 * `signatureValueBytes`, `SignatureValue` becomes that many bytes, below any modulus of that size,
 * so a verifier computes with each key instead of rejecting the length.
 */
export const withKeyInfoCertificates = (
	signed: Uint8Array,
	certificates: readonly string[],
	signatureValueBytes?: number,
): Promise<Uint8Array> =>
	editSignature(signed, (xml) => {
		const keyInfo = xml.replace(
			/<((?:\w+:)?)X509Data>[\s\S]*?<\/(?:\w+:)?X509Data>/u,
			(_data, prefix: string) => {
				const items = certificates.map(
					(certificate) => `<${prefix}X509Certificate>${certificate}</${prefix}X509Certificate>`,
				);
				return `<${prefix}X509Data>${items.join('')}</${prefix}X509Data>`;
			},
		);
		if (signatureValueBytes === undefined) return keyInfo;
		const value = new Uint8Array(signatureValueBytes).fill(0x5a);
		value[0] = 0;
		return keyInfo.replace(
			/(<(?:\w+:)?SignatureValue[^>]*>)[^<]*/u,
			`$1${Buffer.from(value).toString('base64')}`,
		);
	});

/** The Base64 DER certificate that signed `signed`, read from its `KeyInfo`. */
export async function signerCertificateOf(signed: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(signed);
	const xml = await zip.file(signaturePartOf(zip))!.async('string');
	return /<(?:\w+:)?X509Certificate>([^<]+)</u.exec(xml)![1]!.replace(/\s+/g, '');
}
