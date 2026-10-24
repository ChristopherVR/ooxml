import { resetGoogleWebfontSessionCache } from 'ooxml-ui/pptx';
import type { GoogleFontsOption } from 'ooxml-ui/pptx';
// oxlint-disable react-hooks/rules-of-hooks
import type { PptxElement, PptxEmbeddedFont, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { effectScope, nextTick, ref } from 'vue';

import { useGoogleWebfonts } from './useGoogleWebfonts';

const LINK_ID = 'pptx-vue-google-fonts';

beforeEach(() => {
	resetGoogleWebfontSessionCache();
	// happy-dom eagerly fetches injected `<link rel="stylesheet">` elements;
	// disable that so tests stay offline-deterministic, and silence its
	// "loading is disabled" report for the links this composable injects.
	vi.spyOn(console, 'error').mockImplementation(() => {});
	const happy = (
		window as unknown as { happyDOM?: { settings: { disableCSSFileLoading?: boolean } } }
	).happyDOM;
	if (happy) {
		happy.settings.disableCSSFileLoading = true;
	}
});

afterEach(() => {
	document.getElementById(LINK_ID)?.remove();
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
	return { elements } as unknown as PptxSlide;
}

/** The bundled catalogue answers every lookup; nothing may reach the network. */
function stubNetwork(): void {
	vi.stubGlobal('fetch', vi.fn());
}

/** Run `fn` inside an effect scope so watchers + dispose work. */
function withScope<T>(fn: () => T): { result: T; stop: () => void } {
	const scope = effectScope();
	const result = scope.run(fn)!;
	return { result, stop: () => scope.stop() };
}

describe('useGoogleWebfonts', () => {
	it('retains the loaded stylesheet without mutations during slide edits', async () => {
		stubNetwork();
		const slides = ref([slide(textEl('ADLaM Display'))]);
		const { stop } = withScope(() => useGoogleWebfonts(slides, ref([])));
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		const mutations: MutationRecord[] = [];
		const observer = new MutationObserver((records) => mutations.push(...records));
		observer.observe(document.head, { childList: true, attributes: true, subtree: true });
		slides.value = [slide({ ...textEl('ADLaM Display'), x: 200, text: 'Edited' })];
		await nextTick();
		await new Promise((resolve) => {
			setTimeout(resolve, 0);
		});
		observer.disconnect();
		expect(mutations).toHaveLength(0);
		expect(document.getElementById(LINK_ID)).toBe(link);
		stop();
		expect(document.getElementById(LINK_ID)).toBeNull();
	});

	it('injects a Google Fonts link for a catalogue family', async () => {
		stubNetwork();
		const slides = ref([slide(textEl('ADLaM Display'))]);
		withScope(() => useGoogleWebfonts(slides, ref([])));
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const link = document.getElementById(LINK_ID);
		expect(link?.getAttribute('rel')).toBe('stylesheet');
		expect(link?.getAttribute('href')).toContain('family=ADLaM%20Display');
	});

	it('skips the link when the referenced family is embedded', async () => {
		stubNetwork();
		const fetch = vi.fn();
		vi.stubGlobal('fetch', fetch);
		const fonts = ref<PptxEmbeddedFont[]>([
			{ name: 'ADLaM Display', dataUrl: 'data:font/ttf;base64,AA==' },
		]);
		withScope(() => useGoogleWebfonts(ref([slide(textEl('ADLaM Display'))]), fonts));
		await nextTick();
		await Promise.resolve();
		expect(document.getElementById(LINK_ID)).toBeNull();
		expect(fetch).not.toHaveBeenCalled();
	});

	it('updates the managed link when slides change reactively', async () => {
		stubNetwork();
		const slides = ref([slide(textEl('ADLaM Display'))]);
		withScope(() => useGoogleWebfonts(slides, ref([])));
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());

		stubNetwork();
		slides.value = [slide(textEl('Some Local Font'))];
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).toBeNull());
	});

	it('removes the injected link on scope dispose', async () => {
		stubNetwork();
		const slides = ref([slide(textEl('ADLaM Display'))]);
		const { stop } = withScope(() => useGoogleWebfonts(slides, ref([])));
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());

		stop();
		expect(document.getElementById(LINK_ID)).toBeNull();
	});
});

describe('useGoogleWebfonts googleFonts host policy', () => {
	it('injects no link and fetches nothing when googleFonts is false', async () => {
		const fetch = vi.fn();
		vi.stubGlobal('fetch', fetch);
		const { stop } = withScope(() =>
			useGoogleWebfonts(ref([slide(textEl('ADLaM Display'))]), ref([]), () => false),
		);
		await nextTick();
		await new Promise((resolve) => {
			setTimeout(resolve, 0);
		});
		expect(document.getElementById(LINK_ID)).toBeNull();
		expect(fetch).not.toHaveBeenCalled();
		stop();
	});

	it('points the link at a self-hosted mirror', async () => {
		stubNetwork();
		const { stop } = withScope(() =>
			useGoogleWebfonts(ref([slide(textEl('ADLaM Display'))]), ref([]), () => ({
				baseUrl: 'https://fonts.example.test/css2',
			})),
		);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		const href = document.getElementById(LINK_ID)?.getAttribute('href') ?? '';
		expect(href.startsWith('https://fonts.example.test/css2?')).toBe(true);
		expect(href).not.toContain('googleapis');
		stop();
	});

	it('removes the link when the policy switches to false reactively', async () => {
		stubNetwork();
		const policy = ref<GoogleFontsOption | undefined>(undefined);
		const { stop } = withScope(() =>
			useGoogleWebfonts(ref([slide(textEl('ADLaM Display'))]), ref([]), policy),
		);
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).not.toBeNull());
		policy.value = false;
		await vi.waitFor(() => expect(document.getElementById(LINK_ID)).toBeNull());
		stop();
	});
});
