import JSZip from 'jszip';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import { ENTERPRISE_TRUST_ROOTS_PEM_ENV } from '../core/utils/signature-constants';
import type { DigitalSignatureReport, SignatureDetail } from '../core/utils/signature-types';
import {
	createManifestReferences,
	createOfficeSignatureXml,
	createTestPackage,
	createTestParty,
	OFFICE_DOCUMENT_RELATIONSHIP,
	relationshipsXml,
	relationshipXml,
	SLIDE_PART,
	withSignature,
	type TestParty,
} from '../test-support/signature-package';
import { inspectPptxDigitalSignatures } from './inspection';
import { signPptxWithCertificate } from './signing';

const DOCUMENT_SIGNING = { name: 'extKeyUsage', emailProtection: true };

let root: TestParty;
let signer: TestParty;
let serverCertificate: TestParty;

beforeAll(() => {
	root = createTestParty({ name: 'Test Root', ca: true });
	signer = createTestParty({ name: 'Signer', issuer: root, extensions: [DOCUMENT_SIGNING] });
	serverCertificate = createTestParty({
		name: 'server.test',
		issuer: root,
		extensions: [{ name: 'extKeyUsage', serverAuth: true }],
	});
});

beforeEach(() => {
	vi.stubEnv(ENTERPRISE_TRUST_ROOTS_PEM_ENV, root.certPem);
});

afterEach(() => {
	vi.unstubAllEnvs();
});

const detailOf = (report: DigitalSignatureReport): SignatureDetail => report.details![0]!;

/** A package signed in the Office layout, after `mutate` changes it. */
async function signedPackage(
	mutate: (zip: JSZip) => void | Promise<void> = () => undefined,
	signatureOptions: { extraCertificates?: string[]; unsignedSuffix?: string } = {},
	party: TestParty = signer,
): Promise<JSZip> {
	const zip = createTestPackage();
	const signatureXml = createOfficeSignatureXml({
		signer: party,
		manifestReferences: await createManifestReferences(zip),
		...signatureOptions,
	});
	await withSignature(zip, signatureXml);
	await mutate(zip);
	return zip;
}

const inspect = async (zip: JSZip) =>
	inspectPptxDigitalSignatures(await zip.generateAsync({ type: 'uint8array' }));

/** Inspect after rewriting the stored signature part. */
async function inspectWithSignature(zip: JSZip, rewrite: (xml: string) => string) {
	const xml = await zip.file('_xmlsignatures/sig1.xml')!.async('string');
	zip.file('_xmlsignatures/sig1.xml', rewrite(xml));
	return inspect(zip);
}

describe('inspectPptxDigitalSignatures with the Office signature layout', () => {
	it('verifies an intact signature and reports what it leaves uncovered', async () => {
		const report = await inspect(await signedPackage());
		const detail = detailOf(report);

		expect(report.verificationStatus).toBe('verified-trusted');
		expect(detail.referenceChecks[0]?.uri).toBe('#idPackageObject');
		expect(detail.referenceChecks.map((check) => check.digestStatus)).toEqual(
			Array(5).fill('verified'),
		);
		expect(detail.uncoveredParts).toEqual(['docProps/core.xml']);
		expect(detail.coverageGaps).toEqual([]);
		expect(detail.signingTime).toBe('2026-01-02T03:04:05Z');
	});

	it('finds the signer when KeyInfo lists another certificate first', async () => {
		const zip = await signedPackage(undefined, { extraCertificates: [root.certBase64] });
		const report = await inspect(zip);

		expect(report.verificationStatus).toBe('verified-trusted');
		expect(detailOf(report).certificate?.subject).toContain('CN=Signer');
	});

	it('fails when a signed part changes', async () => {
		const zip = await signedPackage((pkg) => void pkg.file(SLIDE_PART, '<sld>edited</sld>'));
		const report = await inspect(zip);
		const check = detailOf(report).referenceChecks.find((c) => c.resolvedPartPath === SLIDE_PART);

		expect(report.verificationStatus).toBe('digest-mismatch');
		expect(check?.digestStatus).toBe('mismatch');
	});

	it('fails when a Manifest digest is rewritten to match edited content', async () => {
		const zip = await signedPackage();
		const original = await zip.file('_xmlsignatures/sig1.xml')!.async('string');
		zip.file(SLIDE_PART, '<sld>edited</sld>');
		const edited = createOfficeSignatureXml({
			signer,
			manifestReferences: await createManifestReferences(zip),
		});
		const manifestOf = (xml: string) => /<Manifest>.*<\/Manifest>/s.exec(xml)![0];
		const report = await inspectWithSignature(zip, () =>
			original.replace(manifestOf(original), manifestOf(edited)),
		);

		expect(report.verificationStatus).toBe('digest-mismatch');
		expect(detailOf(report).referenceChecks).toHaveLength(1);
	});

	it('ignores references outside the signed content', async () => {
		const unsigned =
			`<Object><Manifest><Reference URI="/docProps/core.xml">` +
			`<DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>` +
			`<DigestValue>AAAA</DigestValue></Reference></Manifest></Object>`;
		const zip = await signedPackage((pkg) => void pkg.file(SLIDE_PART, '<sld>edited</sld>'), {
			unsignedSuffix: unsigned,
		});
		const report = await inspect(zip);

		expect(report.verificationStatus).toBe('digest-mismatch');
		expect(detailOf(report).referenceChecks.map((c) => c.uri)).not.toContain('/docProps/core.xml');
	});

	it('binds each reference to the content type the package declares', async () => {
		const zip = await signedPackage(async (pkg) => {
			const types = await pkg.file('[Content_Types].xml')!.async('string');
			pkg.file('[Content_Types].xml', types.replace('.slide+xml', '.notesSlide+xml'));
		});
		const report = await inspect(zip);
		const check = detailOf(report).referenceChecks.find((c) => c.resolvedPartPath === SLIDE_PART);

		expect(report.verificationStatus).toBe('digest-mismatch');
		expect(check?.digestStatus).toBe('content-type-mismatch');
	});

	it('does not verify when a signed reference cannot be checked', async () => {
		const zip = createTestPackage();
		const references = await createManifestReferences(zip);
		const signatureXml = createOfficeSignatureXml({
			signer,
			manifestReferences: [...references, references[3]!.replace(/URI="[^"]+"/, 'URI="#other"')],
		});
		const report = await inspectPptxDigitalSignatures(await withSignature(zip, signatureXml));

		expect(report.verificationStatus).toBe('present-not-verified');
	});
});

describe('signature coverage', () => {
	it('reports a relationship added beside the signed ones', async () => {
		const zip = await signedPackage((pkg) => {
			pkg.file('evil.xml', '<presentation/>');
			pkg.file(
				'_rels/.rels',
				relationshipsXml(
					relationshipXml('rId0', OFFICE_DOCUMENT_RELATIONSHIP, 'evil.xml'),
					relationshipXml('rId1', OFFICE_DOCUMENT_RELATIONSHIP, 'ppt/presentation.xml'),
				),
			);
		});
		const report = await inspect(zip);
		const detail = detailOf(report);

		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detail.coverageGaps).toContainEqual({
			kind: 'unsigned-relationship',
			part: '_rels/.rels',
			relationshipId: 'rId0',
		});
		expect(detail.uncoveredParts).toContain('evil.xml');
	});

	it('reports relationships added to a signed part that had none', async () => {
		const zip = await signedPackage(
			(pkg) =>
				void pkg.file(
					'ppt/slides/_rels/slide1.xml.rels',
					relationshipsXml(relationshipXml('rId1', 'urn:test:comments', 'comments.xml')),
				),
		);
		const report = await inspect(zip);

		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps).toContainEqual({
			kind: 'unsigned-relationships-part',
			part: 'ppt/slides/_rels/slide1.xml.rels',
		});
	});

	it('reports a signature that signs only part of the content', async () => {
		const zip = createTestPackage();
		const [packageRelationships, partRelationships, presentation] =
			await createManifestReferences(zip);
		const signatureXml = createOfficeSignatureXml({
			signer,
			manifestReferences: [packageRelationships!, partRelationships!, presentation!],
		});
		const report = await inspectPptxDigitalSignatures(await withSignature(zip, signatureXml));

		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps).toContainEqual({
			kind: 'unsigned-part',
			part: SLIDE_PART,
			relationshipId: 'rId1',
		});
	});
});

describe('signature structure', () => {
	const signedInfoOf = (xml: string) => /<SignedInfo>.*<\/SignedInfo>/s.exec(xml)![0];

	it.each([
		[
			'a second SignedInfo',
			(xml: string) => xml.replace('<SignatureValue>', `${signedInfoOf(xml)}<SignatureValue>`),
		],
		[
			'a repeated Id',
			(xml: string) => xml.replace('</Signature>', '<Object Id="idPackageObject"/></Signature>'),
		],
		[
			'a document type declaration',
			(xml: string) => xml.replace('<Signature ', '<!DOCTYPE Signature []><Signature '),
		],
		[
			'a wrapped Signature',
			(xml: string) => xml.replace(/<Signature /, '<Wrapper><Signature ') + '</Wrapper>',
		],
	])('rejects %s', async (_name, rewrite) => {
		const report = await inspectWithSignature(await signedPackage(), rewrite);

		expect(report.verificationStatus).toBe('signature-invalid');
		expect(detailOf(report).structureError).toBeDefined();
		expect(detailOf(report).referenceChecks).toEqual([]);
	});

	it('reports an altered SignedInfo as an invalid signature', async () => {
		const report = await inspectWithSignature(await signedPackage(), (xml) =>
			xml.replace('URI="#idPackageObject"', 'URI="#idPackageObject" Id="added"'),
		);

		expect(report.verificationStatus).toBe('signature-invalid');
		expect(detailOf(report).referenceChecks).toEqual([]);
	});
});

describe('signer trust', () => {
	it('does not trust a certificate that does not permit document signing', async () => {
		const report = await inspect(await signedPackage(undefined, {}, serverCertificate));
		const detail = detailOf(report);

		expect(report.verificationStatus).toBe('verified-untrusted');
		expect(detail.certificateTrustError).toMatch(/document signing/);
	});

	it('trusts no root unless one is configured', async () => {
		vi.stubEnv(ENTERPRISE_TRUST_ROOTS_PEM_ENV, '');
		const report = await inspect(await signedPackage());

		expect(report.verificationStatus).toBe('verified-untrusted');
	});
});

describe('signPptxWithCertificate output', () => {
	const material = () => new TextEncoder().encode(signer.keyPem + signer.certPem);
	const pptx = () => createTestPackage().generateAsync({ type: 'uint8array' });

	it('verifies, with a signing time the signature protects', async () => {
		const result = await signPptxWithCertificate(await pptx(), material(), {
			certificatePath: 'signer.pem',
		});

		expect(result.report.verificationStatus).toBe('verified-trusted');
		expect(detailOf(result.report).uncoveredParts).toEqual([]);
		expect(detailOf(result.report).signingTime).toBeDefined();
	});

	it('fails when the package changes after signing', async () => {
		const { signedData } = await signPptxWithCertificate(await pptx(), material(), {
			certificatePath: 'signer.pem',
		});
		const zip = await JSZip.loadAsync(signedData!);
		zip.file(SLIDE_PART, '<sld>edited</sld>');

		expect((await inspect(zip)).verificationStatus).toBe('digest-mismatch');
	});
});
