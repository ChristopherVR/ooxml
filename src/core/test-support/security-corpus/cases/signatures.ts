import crypto from 'node:crypto';

import JSZip from 'jszip';

import { withSignatureLineSignings } from '../../../opc/signature/sign/office-info';
import { verifyPackageSignatureSummary } from '../../../opc/signature/verify/inspect';
import {
	rsaKeyCertificate,
	signForTest,
	withKeyInfoCertificates,
} from '../../../opc/signature/test-support/sign';
import { ENTERPRISE_TRUST_ROOTS_PEM_ENV } from '../../../pptx/core/utils/signature-constants';
import { inspectPptxDigitalSignatures } from '../../../pptx/signature-node/inspection';
import { signPptxWithCertificate } from '../../../pptx/signature-node/signing';
import {
	createManifestReferences,
	createOfficeSignatureXml,
	createTestPackage,
	createTestParty,
	SLIDE_PART,
	type TestParty,
	withSignature,
} from '../../../pptx/test-support/signature-package';
import type { SecurityCase } from '../case';

const SHA256 = 'http://www.w3.org/2001/04/xmlenc#sha256';
const XMLDSIG = 'http://www.w3.org/2000/09/xmldsig#';
const SIGNATURE_PART = '_xmlsignatures/sig1.xml';
const EDITED_SLIDE = '<sld>edited</sld>';
const DOCUMENT_SIGNING = { name: 'extKeyUsage', emailProtection: true };
const SERVER_ONLY = { name: 'extKeyUsage', serverAuth: true };
const LIMITS = { ms: 10_000, heapMb: 160 };

const digest = (data: string | Uint8Array) =>
	crypto.createHash('sha256').update(data).digest('base64');

/** A leaf certificate of `usage`, issued by a local root that this worker alone trusts. */
function trustedSigner(usage: object): TestParty {
	const root = createTestParty({ name: 'Corpus Root', ca: true });
	process.env[ENTERPRISE_TRUST_ROOTS_PEM_ENV] = root.certPem;
	return createTestParty({ name: 'Corpus Signer', issuer: root, extensions: [usage] });
}

/**
 * Sign the test package with the package's own signer and a leaf certificate of `usage`, then let
 * `tamper` change the package and its signature part.
 */
async function signThenInspect(
	usage: object,
	tamper: (zip: JSZip, signatureXml: string) => string | Promise<string> = (_zip, xml) => xml,
): Promise<{ status: string }> {
	const signer = trustedSigner(usage);
	const { signedData } = await signPptxWithCertificate(
		await createTestPackage().generateAsync({ type: 'uint8array' }),
		new TextEncoder().encode(signer.keyPem + signer.certPem),
		{ certificatePath: 'signer.pem' },
	);
	const zip = await JSZip.loadAsync(signedData!);
	zip.file(SIGNATURE_PART, await tamper(zip, await zip.file(SIGNATURE_PART)!.async('string')));
	const report = await inspectPptxDigitalSignatures(
		await zip.generateAsync({ type: 'uint8array' }),
	);
	return { status: report.verificationStatus };
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'part lists outside the signed content do not vouch for an edited part',
		limits: LIMITS,
		run: () =>
			signThenInspect(DOCUMENT_SIGNING, (zip, xml) => {
				zip.file(SLIDE_PART, EDITED_SLIDE);
				// Both list the edited part with its new digest; SignedInfo covers neither.
				const unsigned =
					`<Object><Part xmlns="urn:pptx-viewer:ooxml-signature:v1" Name="/${SLIDE_PART}"` +
					` DigestMethod="${SHA256}" DigestValue="${digest(EDITED_SLIDE)}"/></Object>` +
					`<Object><Manifest><Reference URI="/${SLIDE_PART}"><DigestMethod Algorithm="${SHA256}"/>` +
					`<DigestValue>${digest(EDITED_SLIDE)}</DigestValue></Reference></Manifest></Object>`;
				return xml.replace(/<\/(\w+:)?Signature>\s*$/u, (end) => `${unsigned}${end}`);
			}),
		expected: { status: 'digest-mismatch' },
	},
	{
		name: 'a signed part list rewritten to match an edited part fails the digest of its object',
		limits: LIMITS,
		async run() {
			// The layout Office writes: SignedInfo covers one Object, whose Manifest lists the parts.
			const signer = trustedSigner(DOCUMENT_SIGNING);
			const zip = createTestPackage();
			const manifest = async () =>
				createOfficeSignatureXml({
					signer,
					manifestReferences: await createManifestReferences(zip),
				});
			const signed = await manifest();
			zip.file(SLIDE_PART, EDITED_SLIDE);
			const list = (xml: string) => /<Manifest>.*<\/Manifest>/su.exec(xml)![0];
			const tampered = signed.replace(list(signed), list(await manifest()));
			const report = await inspectPptxDigitalSignatures(await withSignature(zip, tampered));
			return { status: report.verificationStatus };
		},
		expected: { status: 'digest-mismatch' },
	},
	{
		name: 'a certificate that does not permit document signing is never a trusted signer',
		limits: LIMITS,
		run: () => signThenInspect(SERVER_ONLY),
		expected: { status: 'verified-untrusted' },
	},
	{
		name: 'an oversized signer key is not checked, so verifying a signature stays fast',
		limits: { ms: 4_000, heapMb: 64 },
		async run() {
			const bits = 32_768;
			const ones = (1n << BigInt(bits)) - 1n;
			const signed = await signForTest(
				await createTestPackage().generateAsync({ type: 'uint8array' }),
			);
			const bytes = await withKeyInfoCertificates(
				signed,
				[rsaKeyCertificate(ones, ones)],
				bits / 8,
			);
			return { state: (await verifyPackageSignatureSummary(bytes)).state };
		},
		expected: { state: 'unsupported' },
	},
	{
		name: 'a signature line is read only from a well-formed, signed Office object',
		limits: { ms: 2_000, heapMb: 64 },
		async run() {
			const line = '{8F3B2A10-0C4D-4E5F-9A6B-7C8D9E0F1A2B}';
			const info = `<SignatureInfoV1 xmlns="http://schemas.microsoft.com/office/2006/digsig"><SetupID>${line}</SetupID><SignatureImage>${'A'.repeat(3 * 1024 * 1024)}</SignatureImage></SignatureInfoV1>`;
			const parts = {
				// A DTD could make the tree differ from the bytes a signer saw.
				'_xmlsignatures/sig1.xml': `<!DOCTYPE Signature [<!ENTITY e "x">]><Signature xmlns="${XMLDSIG}"><Object Id="idOfficeObject">${info}</Object></Signature>`,
				// Not named by SignedInfo, so its digest protects nothing.
				'_xmlsignatures/sig2.xml': `<Signature xmlns="${XMLDSIG}"><SignedInfo/><Object Id="idOfficeObject">${info}</Object></Signature>`,
				// Named by SignedInfo: the line is read, the oversized picture is not.
				'_xmlsignatures/sig3.xml': `<Signature xmlns="${XMLDSIG}"><SignedInfo><Reference URI="#idOfficeObject"/></SignedInfo><Object Id="idOfficeObject">${info}</Object></Signature>`,
			};
			const zip = new JSZip();
			for (const [path, xml] of Object.entries(parts)) zip.file(path, xml);
			// The verifier's verdict is given here: only the Office object is under test.
			const summary = await withSignatureLineSignings(zip, {
				state: 'verified-untrusted',
				signatures: Object.keys(parts).map((path) => ({
					path,
					state: 'verified-untrusted' as const,
				})),
			});
			return summary.signatures.map((entry) => ({
				line: entry.signatureLineId ?? null,
				image: Boolean(entry.signatureImage),
			}));
		},
		expected: [
			{ line: null, image: false },
			{ line: null, image: false },
			{ line: '{8F3B2A10-0C4D-4E5F-9A6B-7C8D9E0F1A2B}', image: false },
		],
	},
];
