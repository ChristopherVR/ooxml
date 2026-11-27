import JSZip from 'jszip';
import forge from 'node-forge';
import { describe, expect, it } from 'vitest';
import { createDocument } from '../../../docx/model';
import { loadDocx } from '../../../docx/parse';
import { saveDocx } from '../../../docx/save';
import { verifySignedPackage } from '../verify/inspect';
import { testSigner } from '../test-support/sign';
import { withSignatureLineSignings } from './office-info';
import { signPackage } from './sign-package';
import { webCryptoPackageSigner, type PackageSigner } from './signer';
import { signatureLinesOf } from '../../../docx/signature-line';

const LINE = '{8F3B2A10-0C4D-4E5F-9A6B-7C8D9E0F1A2B}';

/** A Web Crypto signer over the generated test certificate. */
async function signer(): Promise<PackageSigner> {
	const party = testSigner();
	const pkcs8 = forge.asn1.toDer(
		forge.pki.wrapRsaPrivateKey(forge.pki.privateKeyToAsn1(party.key)),
	);
	const key = await crypto.subtle.importKey(
		'pkcs8',
		Uint8Array.from(pkcs8.getBytes(), (c) => c.charCodeAt(0)),
		{ name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
		false,
		['sign'],
	);
	return webCryptoPackageSigner(
		key,
		Uint8Array.from(atob(party.certBase64), (c) => c.charCodeAt(0)),
	);
}

describe('signPackage', () => {
	it('signs every part of a document so the verifier accepts it', async () => {
		const signed = await signPackage(await saveDocx(createDocument()), await signer(), {
			comments: 'Approved',
			now: new Date('2026-11-27T10:00:00Z'),
		});
		const summary = await verifySignedPackage(signed);
		expect(summary.state).toBe('verified-untrusted');
		expect(summary.signatures[0]?.signingTime).toContain('2026-11-27');
		const zip = await JSZip.loadAsync(signed);
		const xml = await zip.file('_xmlsignatures/sig1.xml')!.async('string');
		expect(xml).toContain('xd:SigningCertificate');
		expect(xml).toContain('RelationshipTransform');
		expect(xml).toContain('<SignatureType>1</SignatureType>');
	});

	it('records the signature line it signs and reads it back from a verified signature', async () => {
		const signed = await signPackage(await saveDocx(createDocument()), await signer(), {
			signatureLineId: LINE,
			signatureText: 'Ada Lovelace',
		});
		const zip = await JSZip.loadAsync(signed);
		const summary = await withSignatureLineSignings(zip, await verifySignedPackage(signed));
		expect(summary.signatures[0]).toMatchObject({
			signatureLineId: LINE,
			signatureText: 'Ada Lovelace',
		});
		const loaded = await loadDocx(signed);
		expect(loaded.model.signatures?.signatures[0]?.signatureLineId).toBe(LINE);
	});

	it('keeps an earlier signature valid when a second signer signs', async () => {
		const once = await signPackage(await saveDocx(createDocument()), await signer());
		const twice = await signPackage(once, await signer(), { signatureLineId: LINE });
		const summary = await verifySignedPackage(twice);
		expect(summary.signatures.map((entry) => entry.state)).toEqual([
			'verified-untrusted',
			'verified-untrusted',
		]);
	});

	it('does not report a signature line from a signature that no longer verifies', async () => {
		const signed = await signPackage(await saveDocx(createDocument()), await signer(), {
			signatureLineId: LINE,
		});
		const zip = await JSZip.loadAsync(signed);
		const document = await zip.file('word/document.xml')!.async('string');
		zip.file('word/document.xml', document.replace('<w:body>', '<w:body><!-- changed -->'));
		const tampered = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(tampered);
		expect(loaded.model.signatures?.state).toBe('invalid');
		expect(loaded.model.signatures?.signatures[0]?.signatureLineId).toBeUndefined();
	});

	it('refuses a signer whose signature does not verify', async () => {
		const good = await signer();
		const broken: PackageSigner = {
			certificate: good.certificate,
			sign: async () => new Uint8Array(256).fill(7),
		};
		await expect(signPackage(await saveDocx(createDocument()), broken)).rejects.toThrow(
			/does not verify/u,
		);
	});
});

describe('signatureLinesOf', () => {
	it('finds signature lines in paragraphs and table cells', () => {
		const line = { id: LINE, allowComments: false, showSignDate: true };
		const image = {
			relId: '',
			partName: 'word/media/s.png',
			contentType: 'image/png',
			widthPx: 1,
			heightPx: 1,
			signatureLine: line,
		};
		const paragraph = { type: 'paragraph', id: 'p', runs: [{ text: '', image }] };
		const blocks = [paragraph, { type: 'table', id: 't', rows: [[{ paragraphs: [paragraph] }]] }];
		expect(signatureLinesOf(blocks as never)).toEqual([line, line]);
	});
});
