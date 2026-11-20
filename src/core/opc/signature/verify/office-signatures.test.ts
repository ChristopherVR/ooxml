import { readFileSync } from 'node:fs';

import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { inspectPptxDigitalSignatures } from '../../../pptx/signature-node/inspection';
import { ENTERPRISE_TRUST_ROOTS_PEM_ENV } from '../constants';
import { summarizeSignatures } from '../state';
import {
	signForTest,
	tamperSignedPart,
	testRoot,
	withUnsupportedSignatureMethod,
} from '../test-support/sign';
import { verifyPackageSignatureSummary, verifySignedPackage } from './inspect';

const fixture = (name: string): Uint8Array =>
	new Uint8Array(readFileSync(new URL(`../__fixtures__/office/${name}`, import.meta.url)));

const OFFICE_SIGNED = [
	['PowerPoint 2016', 'PPT2016withComment.pptx', 'ppt/slides/slide1.xml', 'CN=Test'],
	['Word 2010', 'hello-world-office-2010-technical-preview.docx', 'word/document.xml', 'CN=Test'],
	[
		'Word 2007, pretty-printed relationships',
		'office2007prettyPrintedRels.docx',
		'word/document.xml',
		'CN=Hamed',
	],
	['Excel 2007', 'hello-world-signed.xlsx', 'xl/workbook.xml', 'CN=Bart Hanssens (Signature)'],
] as const;

describe('signatures Office wrote', () => {
	it.each(OFFICE_SIGNED)(
		'verify in the platform-neutral verifier (%s)',
		async (_label, file, _part, signer) => {
			const summary = await verifyPackageSignatureSummary(fixture(file));
			expect(summary.state).toBe('verified-untrusted');
			expect(summary.signatures).toHaveLength(1);
			expect(summary.signatures[0]).toMatchObject({ state: 'verified-untrusted' });
			expect(summary.signatures[0]?.signer).toContain(signer);
		},
	);

	it.each(OFFICE_SIGNED)(
		'are invalid once a signed part changes (%s)',
		async (_label, file, part) => {
			const tampered = await tamperSignedPart(fixture(file), part);
			expect((await verifyPackageSignatureSummary(tampered)).state).toBe('invalid');
		},
	);

	it.each(OFFICE_SIGNED)('verify the same way in Node (%s)', async (_label, file) => {
		const report = await inspectPptxDigitalSignatures(fixture(file), { checkRevocation: false });
		expect(summarizeSignatures(report).state).toBe('verified-untrusted');
	});
});

describe('each signature state', () => {
	const document = async (): Promise<Uint8Array> => {
		const zip = new JSZip();
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		);
		zip.file(
			'_rels/.rels',
			'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
		);
		zip.file(
			'word/document.xml',
			'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Signed text</w:t></w:r></w:p></w:body></w:document>',
		);
		return zip.generateAsync({ type: 'uint8array' });
	};

	afterEach(() => {
		vi.unstubAllEnvs();
	});

	it('reports an unsigned package as unsigned', async () => {
		expect((await verifyPackageSignatureSummary(await document())).state).toBe('unsigned');
	});

	it('reports an intact signature from an untrusted signer as verified-untrusted', async () => {
		const signed = await signForTest(await document());
		expect((await verifyPackageSignatureSummary(signed)).state).toBe('verified-untrusted');
	});

	it('reports a signer that chains to a configured root as verified, in Node', async () => {
		vi.stubEnv(ENTERPRISE_TRUST_ROOTS_PEM_ENV, testRoot().certPem);
		const signed = await signForTest(await document());
		const report = await inspectPptxDigitalSignatures(signed, { checkRevocation: false });
		expect(summarizeSignatures(report).state).toBe('verified');
	});

	it('reports a changed document as invalid', async () => {
		const signed = await signForTest(await document());
		const tampered = await tamperSignedPart(signed, 'word/document.xml');
		expect((await verifyPackageSignatureSummary(tampered)).state).toBe('invalid');
	});

	it('reports a signature method it does not implement as unsupported, not invalid', async () => {
		const signed = await withUnsupportedSignatureMethod(await signForTest(await document()));
		expect((await verifyPackageSignatureSummary(signed)).state).toBe('unsupported');
	});

	it('reports signature metadata without a readable signature as unsupported', async () => {
		expect((await verifySignedPackage(await document())).state).toBe('unsupported');
	});

	it('verifies a signed relationships part stored as UTF-16', async () => {
		const zip = await JSZip.loadAsync(await document());
		const rels =
			'<?xml version="1.0" encoding="UTF-16"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
		const utf16 = new Uint8Array(2 + rels.length * 2);
		utf16.set([0xff, 0xfe]);
		for (let index = 0; index < rels.length; index += 1)
			utf16[2 + index * 2] = rels.charCodeAt(index);
		zip.file('word/_rels/document.xml.rels', utf16);
		const signed = await signForTest(await zip.generateAsync({ type: 'uint8array' }));
		expect((await verifyPackageSignatureSummary(signed)).state).toBe('verified-untrusted');
	});
});
