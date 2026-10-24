import { resetGoogleWebfontSessionCache } from 'ooxml-ui/pptx';
// @vitest-environment happy-dom
/**
 * Regression coverage for the Google Fonts webfont fallback in
 * `useFontInjection`: a deck that references a family which is neither
 * installed nor embedded (PowerPoint would fetch it as a Microsoft 365 cloud
 * font) must get a Google Fonts stylesheet link when the bundled catalogue
 * lists the family, and must not when the deck embeds it itself or the
 * catalogue does not know it. Nothing may reach the network.
 */
import type { GoogleFontsOption } from 'ooxml-ui/pptx';
import type { PptxElement, PptxEmbeddedFont, PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useFontInjection } from './useFontInjection';

const LINK_ID = 'pptx-google-fonts';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	resetGoogleWebfontSessionCache();
	// happy-dom eagerly fetches injected `<link rel="stylesheet">` elements;
	// disable that so tests stay offline-deterministic, and silence its
	// "loading is disabled" report for the links this hook injects.
	vi.spyOn(console, 'error').mockImplementation(() => {});
	const happy = (
		window as unknown as { happyDOM?: { settings: { disableCSSFileLoading?: boolean } } }
	).happyDOM;
	if (happy) {
		happy.settings.disableCSSFileLoading = true;
	}
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	document.getElementById(LINK_ID)?.remove();
	document.getElementById('pptx-embedded-fonts')?.remove();
	vi.unstubAllGlobals();
	vi.restoreAllMocks();
});

function textEl(fontFamily: string): PptxElement {
	return {
		type: 'text',
		textSegments: [{ style: { fontFamily } }],
	} as unknown as PptxElement;
}

function slide(...elements: PptxElement[]): PptxSlide {
	return {
		id: 's1',
		rId: 'rId1',
		elements,
	} as unknown as PptxSlide;
}

function Probe({
	slides,
	fonts,
	googleFonts,
}: {
	slides: PptxSlide[];
	fonts: PptxEmbeddedFont[];
	googleFonts?: GoogleFontsOption;
}): null {
	useFontInjection({ embeddedFonts: fonts, slides, googleFonts });
	return null;
}

function render(
	slides: PptxSlide[],
	fonts: PptxEmbeddedFont[],
	googleFonts?: GoogleFontsOption,
): void {
	act(() => {
		root.render(<Probe slides={slides} fonts={fonts} googleFonts={googleFonts} />);
	});
}

async function settle(): Promise<void> {
	await act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, 0);
		});
	});
}

/** The bundled catalogue answers every lookup; nothing may reach the network. */
function stubNetwork(): void {
	vi.stubGlobal('fetch', vi.fn());
}

/** A data URL for a minimal TrueType file whose OS/2 table declares `weight`. */
function ttfDataUrl(weight: number, italic = false): string {
	const bytes = new Uint8Array(12 + 16 + 96);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, 0x00010000);
	view.setUint16(4, 1);
	[...'OS/2'].forEach((char, i) => view.setUint8(12 + i, char.charCodeAt(0)));
	view.setUint32(20, 28);
	view.setUint32(24, 96);
	view.setUint16(28 + 4, weight);
	view.setUint16(28 + 62, italic ? 0x1 : 0x40);
	return `data:font/ttf;base64,${btoa(String.fromCharCode(...bytes))}`;
}

describe('useFontInjection google webfonts fallback', () => {
	it('retains the loaded stylesheet without mutations during slide edits', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		const mutations: MutationRecord[] = [];
		const observer = new MutationObserver((records) => mutations.push(...records));
		observer.observe(document.head, { childList: true, attributes: true, subtree: true });
		render([slide({ ...textEl('ADLaM Display'), x: 200, text: 'Edited' })], []);
		expect(document.getElementById(LINK_ID)).toBe(link);
		await act(async () => {
			await new Promise((resolve) => {
				setTimeout(resolve, 0);
			});
		});
		observer.disconnect();
		expect(mutations).toHaveLength(0);
		expect(document.getElementById(LINK_ID)).toBe(link);
	});

	it('updates the URL when the referenced family changes', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		render([slide(textEl('Roboto'))], []);
		await vi.waitFor(() => expect(link?.getAttribute('href')).toContain('family=Roboto'));
		expect(document.getElementById(LINK_ID)).toBe(link);
	});

	it('injects a Google Fonts link for a catalogue family', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		expect(link?.getAttribute('rel')).toBe('stylesheet');
		expect(link?.getAttribute('href')).toContain('family=ADLaM%20Display');
	});

	it('skips the link when the referenced family is embedded', async () => {
		stubNetwork();
		const fetch = vi.fn();
		vi.stubGlobal('fetch', fetch);
		render(
			[slide(textEl('ADLaM Display'))],
			[{ name: 'ADLaM Display', dataUrl: 'data:font/ttf;base64,AA==' }],
		);
		await act(async () => {
			await Promise.resolve();
		});
		expect(document.getElementById(LINK_ID)).toBeNull();
		expect(fetch).not.toHaveBeenCalled();
	});

	it('injects no link when no referenced family is in the catalogue', async () => {
		stubNetwork();
		// 'Segoe UI' has no verified metric-compatible clone on the catalogue
		// (its closest match, Selawik, is not on Google Fonts; see
		// `google-webfonts-metric-clones.ts`), unlike 'Calibri' which now
		// resolves to Carlito.
		render([slide(textEl('Some Local Font'), textEl('Segoe UI'))], []);
		await act(async () => {
			await Promise.resolve();
		});
		expect(document.getElementById(LINK_ID)).toBeNull();
	});

	it('removes the link when the deck changes to unneeded families', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());

		stubNetwork();
		render([slide(textEl('Some Local Font'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).toBeNull());
	});

	it('requests a weight variant from Google instead of embedded files of another weight', async () => {
		stubNetwork();
		// Every slot of "DM Sans Light" holds DM Sans files; the regular one is SemiBold.
		const fonts: PptxEmbeddedFont[] = [
			{ name: 'DM Sans Light', dataUrl: ttfDataUrl(600) },
			{ name: 'DM Sans Light', dataUrl: ttfDataUrl(700), bold: true },
			{ name: 'Roboto Mono', dataUrl: ttfDataUrl(400) },
		];
		render([slide(textEl('DM Sans Light'), textEl('Roboto Mono'))], fonts);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const href = document.getElementById(LINK_ID)?.getAttribute('href') ?? '';
		expect(href).toContain(`family=${encodeURIComponent('DM Sans:ital,wght@0,300;0,400;0,700')}`);
		expect(href).not.toContain('Roboto');
		const css = document.getElementById('pptx-embedded-fonts')?.textContent ?? '';
		expect(css).not.toContain('DM Sans Light');
		expect(css).toContain('font-family: "Roboto Mono"');
	});

	it('registers matching embedded weight variants at their weight', async () => {
		stubNetwork();
		render(
			[slide(textEl('DM Sans Light'))],
			[
				{ name: 'DM Sans Light', dataUrl: ttfDataUrl(300) },
				{ name: 'DM Sans Light', dataUrl: ttfDataUrl(700), bold: true },
			],
		);
		await act(async () => {
			await Promise.resolve();
		});
		const css = document.getElementById('pptx-embedded-fonts')?.textContent ?? '';
		expect([...css.matchAll(/font-weight: (\d+)/g)].map((m) => m[1])).toStrictEqual(['300', '700']);
		expect(document.getElementById(LINK_ID)).toBeNull();
	});
});

describe('useFontInjection googleFonts host policy', () => {
	it('injects no link and fetches nothing when googleFonts is false', async () => {
		const fetch = vi.fn();
		vi.stubGlobal('fetch', fetch);
		render([slide(textEl('ADLaM Display'))], [], false);
		await settle();
		expect(document.getElementById(LINK_ID)).toBeNull();
		expect(fetch).not.toHaveBeenCalled();
	});

	it('points the link at a self-hosted mirror', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], [], { baseUrl: 'https://fonts.example.test/css2' });
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const href = document.getElementById(LINK_ID)?.getAttribute('href') ?? '';
		expect(href.startsWith('https://fonts.example.test/css2?')).toBe(true);
		expect(href).toContain('family=ADLaM%20Display');
		expect(href).not.toContain('googleapis');
	});

	it('removes an injected link when the host switches googleFonts to false', async () => {
		stubNetwork();
		render([slide(textEl('ADLaM Display'))], []);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		render([slide(textEl('ADLaM Display'))], [], false);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).toBeNull());
	});

	it('keeps the link untouched when an equal inline mirror object is passed again', async () => {
		stubNetwork();
		const slides = [slide(textEl('ADLaM Display'))];
		const fonts: PptxEmbeddedFont[] = [];
		render(slides, fonts, { baseUrl: 'https://fonts.example.test/css2' });
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		const mutations: MutationRecord[] = [];
		const observer = new MutationObserver((records) => mutations.push(...records));
		observer.observe(document.head, { childList: true, attributes: true, subtree: true });
		render(slides, fonts, { baseUrl: 'https://fonts.example.test/css2' });
		await settle();
		observer.disconnect();
		expect(mutations).toHaveLength(0);
		expect(document.getElementById(LINK_ID)).toBe(link);
	});
});
