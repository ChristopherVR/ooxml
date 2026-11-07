// @vitest-environment jsdom
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import { createBackstage, type BackstageHandlers } from './backstage';
import type { BackstageOptions } from './backstage-options';
import { EMPTY_WORD_CUSTOMIZATION } from './customization';
import {
	DocxWebfonts,
	fontDownloadOptions,
	setFontDownloads,
	type DocxWebfontHost,
} from './docx-webfonts';
import {
	FONT_DOWNLOADS_KEY,
	clearReaderFontDownloads,
	readerAllowsFontDownloads,
	setReaderAllowsFontDownloads,
} from './font-downloads';
import { documentStats } from 'ooxml-core/docx';
import { deTrust, enTrust, esTrust, frTrust, zhTrust } from './locales/trust';

const BASE = 'https://fonts.example.com/css2';
const webfontLinks = () => document.head.querySelectorAll('link[id^="docx-editor-webfonts"]');

function calibriDocument(): DocumentModel {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hi', fontFamily: 'Calibri' }] }];
	return model;
}

function host(option?: DocxWebfontHost['webfonts']['option']) {
	const webfonts = new DocxWebfonts();
	webfonts.option = option;
	const remeasure = vi.fn();
	const value: DocxWebfontHost = {
		element: document.createElement('div'),
		webfonts,
		model: calibriDocument(),
		shell: { printLayout: { remeasure } },
	};
	return { value, remeasure };
}

function dialogFor(options: BackstageOptions) {
	const setOption = vi.fn();
	const model = createDocument();
	const handlers: BackstageHandlers = {
		fileCommand: () => {},
		close: vi.fn(),
		summary: () => ({
			fileName: 'Report.docx',
			model,
			words: 0,
			stats: documentStats(model),
			saveState: 'saved',
		}),
		options: () => options,
		setProperty: () => {},
		setOption,
		ribbon: () => undefined,
		hiddenActions: () => [],
		setHiddenActions: () => {},
		customization: () => EMPTY_WORD_CUSTOMIZATION,
		fileProviders: () => [],
		openCloudFile: async () => undefined,
		locale: () => 'en',
		regionContext: (region) => ({
			region,
			editor: document.body,
			fileName: 'Report.docx',
			locale: 'en',
			canEdit: true,
			close: () => undefined,
		}),
	};
	const backstage = createBackstage(handlers);
	document.body.append(backstage.element, backstage.optionsDialog);
	backstage.open('options');
	const dialog = backstage.optionsDialog as HTMLElement & { category: string };
	dialog.category = 'trust';
	const row = dialog.shadowRoot!.querySelector<HTMLElement>('[data-key="allowFontDownloads"]');
	return { dialog, row, input: row?.querySelector('input') ?? null, setOption };
}

describe('Word Trust Center: font downloads', () => {
	afterEach(() => {
		for (const link of webfontLinks()) link.remove();
		document.body.replaceChildren();
		clearReaderFontDownloads();
		resetWebfontSessionCache();
	});

	it('is off by default and remembers the reader choice in localStorage', () => {
		expect(readerAllowsFontDownloads()).toBe(false);
		expect(fontDownloadOptions(new DocxWebfonts())).toEqual({
			allowFontDownloads: false,
			fontDownloadsLocked: false,
		});
		setReaderAllowsFontDownloads(true);
		expect(localStorage.getItem(FONT_DOWNLOADS_KEY)).toBe('true');
		expect(FONT_DOWNLOADS_KEY).toBe('ooxml.docx.allowFontDownloads');
		expect(readerAllowsFontDownloads()).toBe(true);
	});

	it('shows the toggle in File > Options > Trust Center and applies it on OK', () => {
		const { row, input, dialog, setOption } = dialogFor({
			locale: 'en',
			theme: 'auto',
			author: 'Ann',
			allowFontDownloads: false,
			fontDownloadsLocked: false,
		});
		expect(row?.textContent).toContain(
			'Download fonts that are not installed or embedded (Google Fonts)',
		);
		expect(row?.querySelector('.info')?.getAttribute('title')).toContain('differ from Word.');
		expect(input?.checked).toBe(false);
		expect(input?.disabled).toBe(false);
		input!.checked = true;
		input!.dispatchEvent(new Event('change'));
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(setOption).toHaveBeenCalledWith('allowFontDownloads', 'true');
	});

	it('shows a host value read-only, with the managed-by reason', () => {
		expect(fontDownloadOptions(host({ baseUrl: BASE }).value.webfonts)).toEqual({
			allowFontDownloads: true,
			fontDownloadsLocked: true,
		});
		setReaderAllowsFontDownloads(true);
		expect(fontDownloadOptions(host(false).value.webfonts).allowFontDownloads).toBe(false);
		const { row, input } = dialogFor({
			locale: 'en',
			theme: 'auto',
			author: 'Ann',
			allowFontDownloads: true,
			fontDownloadsLocked: true,
		});
		expect(input?.checked).toBe(true);
		expect(input?.disabled).toBe(true);
		expect(row?.getAttribute('title')).toBe('This setting is managed by your organization');
	});

	it('applies to the open document at once: on adds the link, off removes it and relays out', async () => {
		const { value, remeasure } = host();
		setFontDownloads(value, true);
		await vi.waitFor(() => expect(webfontLinks()).toHaveLength(1));
		expect(webfontLinks()[0]?.getAttribute('href')).toContain(
			'https://fonts.googleapis.com/css2?family=Carlito',
		);
		setFontDownloads(value, false);
		await vi.waitFor(() => expect(webfontLinks()).toHaveLength(0));
		expect(remeasure).toHaveBeenCalled();
		expect(readerAllowsFontDownloads()).toBe(false);
	});

	it('leaves a host decision alone', async () => {
		const { value } = host(false);
		setFontDownloads(value, true);
		expect(readerAllowsFontDownloads()).toBe(false);
		await value.webfonts.prepare(document, value.model);
		expect(webfontLinks()).toHaveLength(0);
	});

	it('has every string in every Word locale', () => {
		const keys = Object.keys(enTrust).sort();
		for (const locale of [frTrust, deTrust, esTrust, zhTrust]) {
			expect(Object.keys(locale).sort()).toEqual(keys);
			expect(
				locale[
					'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.'
				],
			).toContain('Word');
		}
	});
});
