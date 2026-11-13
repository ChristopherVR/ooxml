import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { inspectSignaturesFromBuffer, parseSignaturesFromBuffer } from './signature-package-list';

async function digestOf(text: string): Promise<string> {
	const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
	return btoa(String.fromCharCode(...new Uint8Array(hash)));
}

function signatureXml(uri: string, digest: string): string {
	return (
		'<?xml version="1.0" encoding="UTF-8"?>' +
		'<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">' +
		'<SignedInfo><SignatureMethod Algorithm="rsa-sha256"/>' +
		`<Reference URI="${uri}"><DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>` +
		`<DigestValue>${digest}</DigestValue></Reference></SignedInfo>` +
		'<SignatureValue>c2ln</SignatureValue>' +
		'<KeyInfo><X509Data><X509Certificate>QUJD</X509Certificate>' +
		'<X509SubjectName>CN=Caf&#xE9; Ltd</X509SubjectName></X509Data></KeyInfo></Signature>'
	);
}

async function signedPackage(digest: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('ppt/presentation.xml', '<p:presentation/>');
	zip.file('_xmlsignatures/sig1.xml', signatureXml('/ppt/presentation.xml', digest));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('package signature listing', () => {
	it('parses signatures with decoded certificate names', async () => {
		const [signature] = await parseSignaturesFromBuffer(await signedPackage('ZGln'));
		expect(signature?.signaturePath).toBe('_xmlsignatures/sig1.xml');
		expect(signature?.certificate?.subject).toBe('CN=Café Ltd');
	});

	it('reports intact and tampered parts as an integrity check only', async () => {
		const good = await inspectSignaturesFromBuffer(
			await signedPackage(await digestOf('<p:presentation/>')),
		);
		expect(good[0]?.integrity).toBe('digestsMatch');
		const bad = await inspectSignaturesFromBuffer(await signedPackage(await digestOf('other')));
		expect(bad[0]?.integrity).toBe('invalid');
	});

	it('returns an empty list for an unsigned or unreadable package', async () => {
		const zip = new JSZip();
		zip.file('a.xml', '<a/>');
		expect(
			await inspectSignaturesFromBuffer(await zip.generateAsync({ type: 'uint8array' })),
		).toEqual([]);
		expect(await inspectSignaturesFromBuffer(new Uint8Array([1, 2, 3]))).toEqual([]);
	});
});
