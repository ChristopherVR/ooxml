// @vitest-environment jsdom
// The Trust Center font download choice: off by default, stored, overridden (and locked) by the
// host's `googleFonts`, and applied to the open workbook at once.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import {
	FONT_DOWNLOADS_KEY,
	clearReaderFontDownloads,
	readerAllowsFontDownloads,
} from './font-downloads';
import { XlsxWebfonts } from './webfonts';

const hosts: HTMLElement[] = [];

function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	hosts.push(host);
	const workbook = createWorkbook();
	const remeasure = vi.fn();
	const fonts = new XlsxWebfonts({ host, workbook: () => workbook, remeasure });
	return { fonts, remeasure, link: () => document.getElementById(fonts.linkId) };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 30));

afterEach(() => {
	for (const host of hosts.splice(0)) host.remove();
	for (const link of document.head.querySelectorAll('link')) link.remove();
	clearReaderFontDownloads();
	resetWebfontSessionCache();
	vi.restoreAllMocks();
});

describe('Excel Trust Center font downloads', () => {
	it('is off by default and left to the reader', () => {
		const { fonts, link } = setup();
		expect(readerAllowsFontDownloads()).toBe(false);
		expect(fonts.lock()).toBeUndefined();
		expect(fonts.allowed()).toBe(false);
		expect(fonts.base()).toBeNull();
		fonts.apply();
		expect(link()).toBeNull();
	});

	it('stores the reader choice and applies it to the open workbook at once', async () => {
		const { fonts, remeasure, link } = setup();
		fonts.setReaderAllows(true);
		expect(localStorage.getItem(FONT_DOWNLOADS_KEY)).toBe('true');
		expect(link()?.getAttribute('href')).toMatch(/^https:\/\/fonts\.googleapis\.com\/css2\?/);
		await settle();
		expect(remeasure).toHaveBeenCalledTimes(1);
		// Another editor (or a reload) reads the stored choice.
		expect(setup().fonts.allowed()).toBe(true);
		fonts.setReaderAllows(false);
		expect(localStorage.getItem(FONT_DOWNLOADS_KEY)).toBe('false');
		expect(link()).toBeNull();
		await settle();
		expect(remeasure).toHaveBeenCalledTimes(2);
		// On again: the stylesheet comes back (the earlier wait is not reused).
		fonts.setReaderAllows(true);
		expect(link()).not.toBeNull();
	});

	it('lets the host value win and lock the toggle to it', () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const { fonts, link } = setup();
		fonts.setReaderAllows(true);
		fonts.option = false;
		fonts.apply();
		expect(fonts.lock()).toBe(false);
		expect(fonts.allowed()).toBe(false);
		expect(link()).toBeNull();
		fonts.option = { baseUrl: 'not a url' };
		expect(fonts.lock()).toBe(false);
		fonts.setReaderAllows(false);
		fonts.option = { baseUrl: 'https://fonts.example/css2' };
		fonts.apply();
		expect(fonts.lock()).toBe(true);
		expect(fonts.allowed()).toBe(true);
		expect(link()?.getAttribute('href')).toMatch(/^https:\/\/fonts\.example\/css2\?/);
	});
});
