import { afterEach, describe, expect, it } from 'vitest';

import { prepareWebfonts, resetWebfontSessionCache, syncGoogleWebfontStylesheet } from './webfonts';

const notInstalled = () => false;

describe('shared webfont loader', () => {
	afterEach(() => {
		document.getElementById('test-webfonts')?.remove();
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
});
