import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { beforeEach, describe, expect, it } from 'vitest';

import { collectDeckWebfontFaces, prepareDeckWebfonts } from './deck-webfont-readiness';
import type { DeckWebfontProgress } from './deck-webfont-readiness';
import { resetGoogleWebfontSessionCache } from './google-webfonts';

function textEl(
	runs: Array<{ text: string; fontFamily?: string; bold?: boolean; italic?: boolean }>,
	base: { fontFamily?: string } = {},
): PptxElement {
	return {
		type: 'text',
		textStyle: base,
		textSegments: runs.map(({ text, ...style }) => ({ text, style })),
	} as unknown as PptxElement;
}

function slide(...elements: PptxElement[]): PptxSlide {
	return { elements } as unknown as PptxSlide;
}

const DECK = [
	slide(
		textEl([{ text: 'The Year Is', fontFamily: 'Playfair Display' }]),
		textEl([{ text: '2026' }], { fontFamily: 'Raleway' }),
		textEl([{ text: 'Unknown', fontFamily: 'Not A Catalogue Face' }]),
	),
];

interface FakeLink {
	id: string;
	rel: string;
	href: string;
	sheet: object | null;
	listeners: Map<string, () => void>;
	getAttribute(name: string): string | null;
	addEventListener(type: string, listener: () => void): void;
	fire(type: 'load' | 'error'): void;
	remove(): void;
}

function fakeLink(): FakeLink {
	return {
		id: '',
		rel: '',
		href: '',
		sheet: null,
		listeners: new Map(),
		getAttribute(name) {
			return name === 'href' ? this.href : null;
		},
		addEventListener(type, listener) {
			this.listeners.set(type, listener);
		},
		fire(type) {
			if (type === 'load') {
				this.sheet = {};
			}
			this.listeners.get(type)?.();
		},
		remove() {},
	};
}

/** A document whose `fonts.load` calls are recorded and settled by the test. */
function fakeDocument() {
	const elements = new Map<string, FakeLink>();
	const loads: Array<{ font: string; text: string; resolve: () => void }> = [];
	const doc = {
		getElementById: (id: string) => elements.get(id) ?? null,
		createElement: () => fakeLink(),
		head: { appendChild: (link: FakeLink) => elements.set(link.id, link) },
		fonts: {
			load: (font: string, text: string) =>
				new Promise<unknown[]>((resolve) => loads.push({ font, text, resolve: () => resolve([]) })),
		},
	};
	return { doc: doc as unknown as Document, elements, loads };
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

beforeEach(() => {
	resetGoogleWebfontSessionCache();
});

describe('collectDeckWebfontFaces', () => {
	it('lists each face the slides draw with the characters drawn in it', () => {
		const { faces, unavailable } = collectDeckWebfontFaces(DECK, [
			'Playfair Display',
			'Raleway',
			'Not A Catalogue Face',
		]);
		expect([...faces.keys()]).toEqual(['Playfair Display', 'Raleway']);
		expect(faces.get('Raleway')?.[0]?.font).toBe('400 16px "Raleway"');
		expect([...(faces.get('Raleway')?.[0]?.chars ?? [])].join('')).toBe('206');
		expect(unavailable).toEqual(['Not A Catalogue Face']);
	});

	it('keeps bold and italic runs as faces of their own', () => {
		const deck = [
			slide(
				textEl([
					{ text: 'a', fontFamily: 'Raleway' },
					{ text: 'b', fontFamily: 'Raleway', bold: true, italic: true },
				]),
			),
		];
		const fonts = collectDeckWebfontFaces(deck, ['Raleway'])
			.faces.get('Raleway')
			?.map((f) => f.font);
		expect(fonts).toEqual(['400 16px "Raleway"', 'italic 700 16px "Raleway"']);
	});
});

describe('prepareDeckWebfonts', () => {
	const base = 'https://fonts.googleapis.com/css2';
	const notInstalled = () => false;

	it('does nothing when the fallback is disabled', async () => {
		const { doc, elements } = fakeDocument();
		const result = await prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: null,
			isLocallyInstalled: notInstalled,
		});
		expect(result).toEqual({ href: null, families: [], unavailable: [], timedOut: false });
		expect(elements.size).toBe(0);
	});

	it('injects the link and waits for the stylesheet, then for every face', async () => {
		const { doc, elements, loads } = fakeDocument();
		const progress: (DeckWebfontProgress | null)[] = [];
		let settled = false;
		const pending = prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
			progressDelayMs: 5,
			onProgress: (p) => progress.push(p && { ...p }),
		}).then((result) => {
			settled = true;
			return result;
		});
		await flush();
		const link = elements.get('fonts');
		expect(link?.href).toContain('family=Playfair%20Display');
		// Nothing is loaded before the stylesheet has declared the faces.
		expect(loads).toHaveLength(0);
		await new Promise((resolve) => setTimeout(resolve, 10));
		expect(progress).toEqual([{ loaded: 0, total: 2, families: ['Playfair Display', 'Raleway'] }]);
		link?.fire('load');
		await flush();
		expect(loads.map((l) => l.font)).toEqual(['400 16px "Playfair Display"', '400 16px "Raleway"']);
		loads[0]?.resolve();
		await flush();
		expect(settled).toBe(false);
		expect(progress.at(-1)?.loaded).toBe(1);
		loads[1]?.resolve();
		const result = await pending;
		expect(result.families).toEqual(['Playfair Display', 'Raleway']);
		expect(result.unavailable).toEqual(['Not A Catalogue Face']);
		expect(result.timedOut).toBe(false);
	});

	it('reports no progress for a load faster than the delay', async () => {
		const { doc, elements, loads } = fakeDocument();
		const progress: (DeckWebfontProgress | null)[] = [];
		const pending = prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
			progressDelayMs: 1000,
			onProgress: (p) => progress.push(p),
		});
		await flush();
		elements.get('fonts')?.fire('load');
		await flush();
		for (const load of loads) load.resolve();
		await pending;
		expect(progress).toEqual([]);
	});

	it('stops waiting at the timeout and, with no follow, reports nothing afterwards', async () => {
		const { doc, elements, loads } = fakeDocument();
		const progress: (DeckWebfontProgress | null)[] = [];
		const pending = prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
			timeoutMs: 15,
			progressDelayMs: 5,
			followMs: 0,
			onProgress: (p) => progress.push(p),
		});
		await flush();
		elements.get('fonts')?.fire('load');
		const result = await pending;
		expect(result.timedOut).toBe(true);
		const reported = progress.length;
		for (const load of loads) load.resolve();
		await flush();
		expect(progress).toHaveLength(reported);
	});

	it('does not hang when the stylesheet fails to load', async () => {
		const { doc, elements } = fakeDocument();
		(doc as unknown as { fonts: { load: () => Promise<never> } }).fonts.load = () =>
			Promise.reject(new Error('no face'));
		const pending = prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
		});
		await flush();
		elements.get('fonts')?.fire('error');
		await expect(pending).resolves.toMatchObject({ timedOut: false });
	});

	it('keeps reporting a download that outlasts the wait, then reports null', async () => {
		const { doc, elements, loads } = fakeDocument();
		const progress: (DeckWebfontProgress | null)[] = [];
		const pending = prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
			timeoutMs: 5,
			progressDelayMs: 10,
			onProgress: (p) => progress.push(p && { ...p }),
		});
		await flush();
		elements.get('fonts')?.fire('load');
		// The deck is released while the faces are still downloading.
		await expect(pending).resolves.toMatchObject({ timedOut: true });
		expect(progress).toEqual([]);
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(progress).toEqual([{ loaded: 0, total: 2, families: ['Playfair Display', 'Raleway'] }]);
		loads[0]?.resolve();
		await flush();
		expect(progress.at(-1)).toMatchObject({ loaded: 1 });
		loads[1]?.resolve();
		await flush();
		await flush();
		expect(progress.at(-1)).toBeNull();
		const reported = progress.length;
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(progress).toHaveLength(reported);
	});

	it('gives up following at followMs and clears the progress', async () => {
		const { doc, elements } = fakeDocument();
		const progress: (DeckWebfontProgress | null)[] = [];
		await prepareDeckWebfonts({
			document: doc,
			linkId: 'fonts',
			slides: DECK,
			embeddedFonts: [],
			googleFontsBase: base,
			isLocallyInstalled: notInstalled,
			timeoutMs: 5,
			progressDelayMs: 10,
			followMs: 30,
			onProgress: (p) => progress.push(p && { ...p }),
		});
		elements.get('fonts')?.fire('load');
		await new Promise((resolve) => setTimeout(resolve, 50));
		expect(progress[0]).toMatchObject({ loaded: 0 });
		expect(progress.at(-1)).toBeNull();
	});
});
