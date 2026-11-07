// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, saveDocx, type DocumentModel } from 'ooxml-core/docx';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import { DocxWebfonts, docxGoogleFontsBase } from './docx-webfonts';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

const BASE = 'https://fonts.example.com/css2';
const notInstalled = () => false;

function calibriDocument(): DocumentModel {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello', fontFamily: 'Calibri' }] }];
	return model;
}

const webfontLinks = () => document.head.querySelectorAll('link[id^="docx-editor-webfonts"]');

describe('Word webfont option', () => {
	afterEach(() => {
		for (const link of webfontLinks()) link.remove();
		document.body.replaceChildren();
		resetWebfontSessionCache();
		vi.restoreAllMocks();
	});

	it('is off unless the host sets a base URL', () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		expect(docxGoogleFontsBase(undefined)).toBeNull();
		expect(docxGoogleFontsBase(false)).toBeNull();
		expect(docxGoogleFontsBase({ baseUrl: 'ftp://fonts.example.com/css2' })).toBeNull();
		expect(warn).toHaveBeenCalledTimes(1);
		expect(docxGoogleFontsBase({ baseUrl: BASE })).toBe(BASE);
	});

	it.each([
		['unset', undefined],
		['false', false as const],
		['an invalid base URL', { baseUrl: 'https://user:pw@fonts.example.com/css2' }],
	])('adds no <link> when %s', async (_label, option) => {
		vi.spyOn(console, 'warn').mockImplementation(() => {});
		const webfonts = new DocxWebfonts();
		webfonts.option = option;
		const result = await webfonts.prepare(document, calibriDocument(), {
			isLocallyInstalled: notInstalled,
		});
		expect(result.href).toBeNull();
		expect(webfontLinks()).toHaveLength(0);
	});

	it('links the metric clone of a missing family from the host base URL, one link per editor', async () => {
		const first = new DocxWebfonts();
		const second = new DocxWebfonts();
		first.option = { baseUrl: BASE };
		const result = await first.prepare(document, calibriDocument(), {
			isLocallyInstalled: notInstalled,
			timeoutMs: 10,
		});
		expect(result.href).toMatch(/^https:\/\/fonts\.example\.com\/css2\?family=Carlito/);
		const link = document.getElementById(first.linkId);
		expect(link?.getAttribute('href')).toBe(result.href);
		expect(second.linkId).not.toBe(first.linkId);
		// Turning it off removes this editor's stylesheet.
		first.option = false;
		await first.prepare(document, calibriDocument());
		expect(document.getElementById(first.linkId)).toBeNull();
	});

	it('wires the element: off by default, a base URL requests the clone', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		expect(editor.googleFonts).toBeUndefined();
		await editor.load(await saveDocx(calibriDocument()));
		expect(webfontLinks()).toHaveLength(0);
		editor.googleFonts = { baseUrl: BASE };
		await vi.waitFor(() => expect(webfontLinks()).toHaveLength(1));
		expect(webfontLinks()[0]?.getAttribute('href')).toContain(`${BASE}?family=Carlito`);
		editor.googleFonts = false;
		await vi.waitFor(() => expect(webfontLinks()).toHaveLength(0));
	});
});
