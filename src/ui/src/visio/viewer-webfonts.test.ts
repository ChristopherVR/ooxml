import { afterEach, describe, expect, it, vi } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { prepareWebfonts, resetWebfontSessionCache } from '../fonts/webfonts';
import { ViewerWebfonts } from './viewer-webfonts';

const drawing = (fontFamily: string): VisioDocument =>
	({
		format: 'vsdx',
		diagnostics: [],
		pages: [
			{
				id: '0',
				name: 'Page-1',
				shapes: [
					{
						id: '1',
						name: 'Box',
						hidden: false,
						children: [],
						text: { plainText: 'Hello', runs: [], fontFamily, bold: false, italic: false },
					},
				],
			},
		],
	}) as unknown as VisioDocument;

// Nothing counts as installed and the wait is short, whatever the test DOM offers.
const quick: typeof prepareWebfonts = (options) =>
	prepareWebfonts({ ...options, isLocallyInstalled: () => false, timeoutMs: 20 });

const links = () =>
	document.head.querySelectorAll<HTMLLinkElement>('link[id^="visio-viewer-webfonts"]');

afterEach(() => {
	for (const link of links()) link.remove();
	resetWebfontSessionCache();
	vi.restoreAllMocks();
});

describe('ViewerWebfonts', () => {
	it('is off by default and requests nothing', async () => {
		const prepare = vi.fn(quick);
		const fonts = new ViewerWebfonts(() => document, prepare);
		expect(fonts.option).toBeUndefined();
		expect(fonts.base()).toBeNull();
		expect(await fonts.load(drawing('Calibri'))).toBeUndefined();
		expect(prepare).not.toHaveBeenCalled();
		expect(links()).toHaveLength(0);
	});

	it('makes no link when turned off with false', async () => {
		const fonts = new ViewerWebfonts(() => document, quick);
		fonts.option = false;
		expect(await fonts.load(drawing('Calibri'))).toBeUndefined();
		expect(links()).toHaveLength(0);
	});

	it('fails closed on an invalid baseUrl', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		const fonts = new ViewerWebfonts(() => document, quick);
		fonts.option = { baseUrl: 'ftp://fonts.example.test/css2?x=1' };
		expect(await fonts.load(drawing('Calibri'))).toBeUndefined();
		expect(links()).toHaveLength(0);
		expect(warn.mock.calls[0]?.[0]).toContain('[visio-viewer]');
	});

	it('links the drawing fonts to the given base, as their metric clones', async () => {
		const fonts = new ViewerWebfonts(() => document, quick);
		fonts.option = { baseUrl: 'https://fonts.example.test/css2' };
		await fonts.load(drawing('Calibri'));
		const link = document.getElementById(fonts.linkId) as HTMLLinkElement | null;
		expect(link?.getAttribute('href')).toMatch(
			/^https:\/\/fonts\.example\.test\/css2\?family=Carlito/u,
		);
		// Turning it off removes the stylesheet again.
		fonts.option = false;
		expect(document.getElementById(fonts.linkId)).toBeNull();
	});

	it('requests nothing for a family the catalogue does not serve', async () => {
		const fonts = new ViewerWebfonts(() => document, quick);
		fonts.option = { baseUrl: 'https://fonts.example.test/css2' };
		expect(await fonts.load(drawing('Private Corporate Face'))).toBeUndefined();
		expect(links()).toHaveLength(0);
	});

	it('gives every viewer its own link and ignores a superseded wait', async () => {
		let release: () => void = () => {};
		const slow: typeof prepareWebfonts = async (options) => {
			const result = await quick(options);
			await new Promise<void>((resolve) => (release = resolve));
			return { ...result, families: ['Carlito'] };
		};
		const first = new ViewerWebfonts(() => document, slow);
		const second = new ViewerWebfonts(() => document, quick);
		expect(first.linkId).not.toBe(second.linkId);
		first.option = { baseUrl: 'https://fonts.example.test/css2' };
		const pending = first.load(drawing('Calibri'));
		await vi.waitFor(() => expect(document.getElementById(first.linkId)).not.toBeNull());
		first.option = { baseUrl: 'https://mirror.example.test/css2' };
		release();
		expect(await pending).toBeUndefined();
	});
});
