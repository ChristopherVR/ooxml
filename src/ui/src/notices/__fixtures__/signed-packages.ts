/**
 * Signed packages in each signature state, for the product tests: the Office-signed samples of
 * `src/core/opc/signature/__fixtures__/office` (they verify, but no signer is trusted), with one
 * signed part changed (invalid) or an unknown signature method (unsupported). Built with JSZip
 * only, so no signing code or key is involved.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

// Tests run from src/ui; a DOM environment's import.meta.url is not a file URL.
const FIXTURES = resolve(process.cwd(), '../core/opc/signature/__fixtures__');

export const officeSigned = (name: string): Uint8Array =>
	new Uint8Array(readFileSync(resolve(FIXTURES, 'office', name)));

/** A package the test signer signed (`generated/README.md`): it verifies, the signer is not trusted. */
export const testSigned = (name: string): Uint8Array =>
	new Uint8Array(readFileSync(resolve(FIXTURES, 'generated', name)));

async function rewrite(bytes: Uint8Array, change: (zip: JSZip) => Promise<void>) {
	const zip = await JSZip.loadAsync(bytes);
	await change(zip);
	return zip.generateAsync({ type: 'uint8array' });
}

/** `part` changed after signing: the signature no longer verifies. */
export const tampered = (signed: Uint8Array, part: string): Promise<Uint8Array> =>
	rewrite(signed, async (zip) => {
		const text = await zip.file(part)!.async('string');
		const end = text.indexOf('>', text.indexOf('<', text.indexOf('?>') + 2));
		zip.file(part, `${text.slice(0, end + 1)}<!-- changed -->${text.slice(end + 1)}`);
	});

/** The signature names a method the verifier does not implement: it is verified neither way. */
export const unsupportedMethod = (signed: Uint8Array): Promise<Uint8Array> =>
	rewrite(signed, async (zip) => {
		for (const path of Object.keys(zip.files).filter((name) =>
			/^_xmlsignatures\/.+\.xml$/u.test(name),
		)) {
			const xml = await zip.file(path)!.async('string');
			zip.file(
				path,
				xml.replace(
					/(<(?:\w+:)?SignatureMethod Algorithm=")[^"]+"/u,
					'$1http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"',
				),
			);
		}
	});
