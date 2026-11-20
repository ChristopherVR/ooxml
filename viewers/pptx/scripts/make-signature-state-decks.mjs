/**
 * Generates the four decks in `e2e/pptx/fixtures/signature-states/`, one per signature state a
 * browser can reach, behind `e2e/pptx/signature-state.spec.ts`:
 *
 * - `unsigned.pptx`: no signature parts.
 * - `verified-untrusted.pptx`: signed by a self-signed certificate. The signature verifies, and a
 *   browser trusts no signer, so it is `verified-untrusted`.
 * - `invalid.pptx`: the signed deck with its slide changed after signing (a digest mismatch).
 * - `unsupported.pptx`: the signed deck with its `SignatureMethod` swapped for one the verifier
 *   does not implement (ECDSA-SHA256), so it is verified neither way.
 *
 * The certificate and its private key exist only in memory for the length of one run: nothing
 * here writes a key, and each run signs with a fresh one, so the signed decks change on every
 * run. Rerun it only when the decks need to change, and commit all four together.
 *
 * Each deck is read back with the browser verifier (`inspectSignaturesFromBuffer`) and the run
 * fails unless it reaches its state. Needs the core built (`bun run build` at the repository root)
 * and the `pptx-viewer-core` entries (`bun run build` in `packages/core`).
 *
 *   bun run scripts/make-signature-state-decks.mjs
 */
import crypto from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { inspectSignaturesFromBuffer, PptxHandler } from 'pptx-viewer-core';
import { signPptxWithCertificate } from 'pptx-viewer-core/signature-node';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(here, '../../../e2e/pptx/fixtures/signature-states');

// node-forge is an optional peer of the core's signer and not a dependency of this folder; resolve
// it from the core package's own scope, as the e2e fixture writer does for JSZip.
const forge = createRequire(createRequire(import.meta.url).resolve('pptx-viewer-core'))(
	'node-forge',
);

const SLIDE_PART = 'ppt/slides/slide1.xml';
const FIXED_W3CDTF = '2000-01-01T00:00:00Z';

/** A one-slide deck with a title, its core-property dates pinned so the unsigned deck is stable. */
async function baseDeck() {
	const { handler, data, createSlide } = await PptxHandler.createBlank({
		title: 'Signature State Fixture',
		initialSlideCount: 0,
	});
	data.slides.push(
		createSlide('Blank')
			.addShape('rect', {
				x: 120,
				y: 200,
				width: 720,
				height: 160,
				fill: { type: 'solid', color: '#DCE6F2' },
				text: 'Signature state fixture',
				textStyle: { fontSize: 32, color: '#1F3864' },
			})
			.build(),
	);
	return rewrite(await handler.save(data.slides), async (zip) => {
		const xml = await zip.file('docProps/core.xml').async('string');
		zip.file(
			'docProps/core.xml',
			xml.replace(
				/(<dcterms:(created|modified)\b[^>]*>)[^<]*(<\/dcterms:\2>)/gu,
				`$1${FIXED_W3CDTF}$3`,
			),
		);
	});
}

async function rewrite(bytes, change) {
	const zip = await JSZip.loadAsync(bytes);
	await change(zip);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** A self-signed document-signing certificate and its key, as one PEM buffer held in memory. */
function selfSignedSigner() {
	const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
	const keyPem = privateKey.export({ type: 'pkcs1', format: 'pem' });
	const cert = forge.pki.createCertificate();
	cert.publicKey = forge.pki.publicKeyFromPem(publicKey.export({ type: 'spki', format: 'pem' }));
	cert.serialNumber = crypto
		.randomBytes(8)
		.toString('hex')
		.replace(/^[89a-f]/u, '1');
	// Valid well past any test run, so the committed decks never expire.
	cert.validity.notBefore = new Date(Date.UTC(2020, 0, 1));
	cert.validity.notAfter = new Date(Date.UTC(2100, 0, 1));
	const subject = [{ name: 'commonName', value: 'OOXML Signature Fixture' }];
	cert.setSubject(subject);
	cert.setIssuer(subject);
	cert.setExtensions([
		{ name: 'basicConstraints', cA: false },
		{ name: 'keyUsage', digitalSignature: true, nonRepudiation: true },
		{ name: 'extKeyUsage', emailProtection: true },
	]);
	cert.sign(forge.pki.privateKeyFromPem(keyPem), forge.md.sha256.create());
	return new TextEncoder().encode(`${keyPem}\n${forge.pki.certificateToPem(cert)}`);
}

async function sign(bytes) {
	const result = await signPptxWithCertificate(bytes, selfSignedSigner(), {
		certificatePath: 'signer.pem',
	});
	if (!result.success || !result.signedData) throw new Error(result.error ?? 'Signing failed.');
	return result.signedData;
}

/** A comment just inside the slide's root element changes its bytes and nothing a reader models. */
const tamperSlide = (signed) =>
	rewrite(signed, async (zip) => {
		const text = await zip.file(SLIDE_PART).async('string');
		const declaration = text.indexOf('?>');
		const rootEnd = text.indexOf('>', text.indexOf('<', declaration < 0 ? 0 : declaration + 2));
		zip.file(SLIDE_PART, `${text.slice(0, rootEnd + 1)}<!-- changed -->${text.slice(rootEnd + 1)}`);
	});

const withUnsupportedSignatureMethod = (signed) =>
	rewrite(signed, async (zip) => {
		const path = Object.keys(zip.files).find((name) => /^_xmlsignatures\/.+\.xml$/u.test(name));
		const xml = await zip.file(path).async('string');
		zip.file(
			path,
			xml.replace(
				/(<SignatureMethod Algorithm=")[^"]+"/u,
				'$1http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"',
			),
		);
	});

const unsigned = await baseDeck();
const signed = await sign(unsigned);
const decks = {
	unsigned,
	'verified-untrusted': signed,
	invalid: await tamperSlide(signed),
	unsupported: await withUnsupportedSignatureMethod(signed),
};

let failed = false;
for (const [expected, bytes] of Object.entries(decks)) {
	const { state, signatures } = await inspectSignaturesFromBuffer(bytes);
	failed ||= state !== expected;
	console.log(`${expected}.pptx: ${state} (${signatures.length} signature(s))`);
}
// Write nothing unless every deck reaches its state, so a failed run leaves the committed set whole.
if (failed) {
	console.error('A deck did not reach its signature state; nothing was written.');
	process.exit(1);
}
await mkdir(outDir, { recursive: true });
for (const [state, bytes] of Object.entries(decks)) {
	await writeFile(resolve(outDir, `${state}.pptx`), bytes);
}
console.log(`Wrote ${Object.keys(decks).length} decks to ${outDir}`);
