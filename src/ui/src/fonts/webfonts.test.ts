import JSZip from 'jszip';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import {
	prefetchWebfonts,
	prepareWebfonts,
	resetWebfontSessionCache,
	syncGoogleWebfontStylesheet,
} from './webfonts';

const notInstalled = () => false;

let docxBytes: Uint8Array;
async function docxWith(font: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/styles.xml', `<w:styles><w:rFonts w:ascii="${font}"/></w:styles>`);
	zip.file(
		'word/fontTable.xml',
		'<w:fonts><w:font w:name="Lato"><w:embedRegular r:id="rId1"/></w:font></w:fonts>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('shared webfont loader', () => {
	beforeAll(async () => {
		docxBytes = await docxWith('Calibri');
	});
	afterEach(() => {
		document.getElementById('test-webfonts')?.remove();
		document.getElementById('test-webfonts-prefetch')?.remove();
		resetWebfontSessionCache();
	});

	it('adds nothing when the fallback is off', async () => {
		const result = await prepareWebfonts({
			document,
			linkId: 'test-webfonts',
			runs: [{ family: 'Calibri', text: 'a' }],
			base: null,
		});
		expect(result.href).toBeNull();
		expect(document.getElementById('test-webfonts')).toBeNull();
	});

	it('links the clone of a missing family from the chosen server, never an embedded one', async () => {
		const result = await prepareWebfonts({
			document,
			linkId: 'test-webfonts',
			runs: [
				{ family: 'Calibri', text: 'a' },
				{ family: 'Lobster', text: 'b' },
			],
			embedded: ['Lobster'],
			base: 'https://fonts.example.com/css2',
			isLocallyInstalled: notInstalled,
			timeoutMs: 10,
		});
		const link = document.getElementById('test-webfonts') as HTMLLinkElement;
		expect(link.getAttribute('href')).toBe(result.href);
		expect(result.href).toMatch(/^https:\/\/fonts\.example\.com\/css2\?family=Carlito/);
		expect(result.href).not.toContain('Lobster');
	});

	it('leaves an installed family alone and keeps an unchanged link', async () => {
		const result = await prepareWebfonts({
			document,
			linkId: 'test-webfonts',
			runs: [{ family: 'Roboto', text: 'a' }],
			base: 'https://fonts.example.com/css2',
			isLocallyInstalled: () => true,
		});
		expect(result.href).toBeNull();
		const first = syncGoogleWebfontStylesheet(document, 'test-webfonts', 'https://a/x');
		expect(syncGoogleWebfontStylesheet(document, 'test-webfonts', 'https://a/x')).toBe(first);
	});

	it('starts the declared fonts before parsing, never an embedded one, and keeps them while loading', async () => {
		const href = prefetchWebfonts({
			document,
			linkId: 'test-webfonts',
			bytes: docxBytes,
			base: 'https://fonts.example.com/css2',
			isLocallyInstalled: notInstalled,
		});
		const early = document.getElementById('test-webfonts-prefetch') as HTMLLinkElement;
		expect(early.getAttribute('href')).toBe(href);
		expect(href).toContain('family=Carlito');
		expect(href).not.toContain('Lato');
		// A binding clearing its own link mid-load leaves the early one alone.
		syncGoogleWebfontStylesheet(document, 'test-webfonts', null);
		expect(document.getElementById('test-webfonts-prefetch')).not.toBeNull();
		// Once parsed it settles; turning downloads off then removes both.
		await prepareWebfonts({
			document,
			linkId: 'test-webfonts',
			runs: [],
			base: 'https://fonts.example.com/css2',
			isLocallyInstalled: notInstalled,
		});
		syncGoogleWebfontStylesheet(document, 'test-webfonts', null);
		expect(document.getElementById('test-webfonts-prefetch')).toBeNull();
	});

	it('adds and keeps nothing early when downloads are off or the bytes are not a package', () => {
		prefetchWebfonts({
			document,
			linkId: 'test-webfonts',
			bytes: docxBytes,
			base: 'https://f.example/css2',
			isLocallyInstalled: notInstalled,
		});
		expect(
			prefetchWebfonts({
				document,
				linkId: 'test-webfonts',
				bytes: docxBytes,
				base: null,
			}),
		).toBeNull();
		expect(document.getElementById('test-webfonts-prefetch')).toBeNull();
		expect(
			prefetchWebfonts({
				document,
				linkId: 'test-webfonts',
				bytes: new Uint8Array([1, 2, 3]),
				base: 'https://f.example/css2',
			}),
		).toBeNull();
	});
});
