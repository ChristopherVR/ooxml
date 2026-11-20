import crypto from 'node:crypto';

import JSZip from 'jszip';

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
];
