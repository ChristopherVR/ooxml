/**
 * Undo and Redo of a theme change take the theme back as well as the slides: in the editor's
 * theme, in the theme the viewer publishes, and in the handler's theme part a save writes.
 *
 * The real handler over `sample-deck.pptx` (an Office theme, "Office Theme", accent 1 4472C4,
 * headings in Calibri Light), wired to the real editor state and ribbon gallery host.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { GALLERY_THEME_PRESETS } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxTheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import {
	buildTwoThemesDeck,
	themePartHeadingFaces,
	TWO_THEMES,
} from '../../../../../../../src/core/pptx/__tests__/helpers/two-themes-deck';
import type { TwoThemesDeckOptions } from '../../../../../../../src/core/pptx/__tests__/helpers/two-themes-deck';
import { createRibbonGalleryHost } from '../components/ribbon/galleries/ribbon-gallery-host';
import { createInspectorDeckActions } from '../state/inspector-deck';
import type { PresentationLoader } from '../state/presentation-loader.svelte';
import { EditorState } from './editor-state.svelte';
import { applyThemePreset } from './editor-theme-scheme';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

const OFFICE = { name: 'Office Theme', accent1: '4472C4', major: 'Calibri Light' };
const FACET = { name: 'Facet', accent1: '90C226', major: 'Trebuchet MS' };

async function open() {
	const bytes = readFileSync(DECK);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const published: { theme: PptxTheme | undefined } = { theme: data.theme };
	const editor = new EditorState({
		getCurrent: () => 0,
		getHandler: () => handler,
		getThemeOptions: () => data.themeOptions ?? [],
		onThemeRestored: (theme) => {
			published.theme = theme;
		},
	});
	editor.editable = true;
	editor.setSlides(data.slides, data.slideMasters);
	editor.theme = data.theme;
	const host = createRibbonGalleryHost(editor, (theme) => {
		published.theme = theme;
	});
	/** Pick a tile the way the ribbon does. */
	const pick = async (id: RibbonGalleryId, itemId: string) => {
		await host.apply(id, itemId);
		expect(editor.canUndo).toBeTruthy();
	};
	/** The theme part as File > Save writes it now. */
	const saved = async () => {
		const zip = await JSZip.loadAsync(await editor.save());
		const xml = (await zip.file('ppt/theme/theme1.xml')?.async('string')) ?? '';
		const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
		return {
			name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
			accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u).toUpperCase(),
			major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
		};
	};
	return { editor, handler, host, published, pick, saved, loaded: data.theme };
}

describe('svelte theme undo', () => {
	it('a Themes pick writes the theme part, and Undo and Redo take the whole theme', async () => {
		const deck = await open();
		await deck.pick('themes', 'facet');
		expect(deck.editor.theme?.colorScheme?.accent1).toBe('#90C226');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);

		deck.editor.undo();
		expect(deck.editor.theme).toStrictEqual(deck.loaded);
		expect(deck.published.theme).toStrictEqual(deck.loaded);
		expect(deck.editor.canUndo).toBeFalsy();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		deck.editor.redo();
		expect(deck.editor.theme?.name).toBe('Facet');
		expect(deck.published.theme?.name).toBe('Facet');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
	});

	it('a Fonts pick re-resolves the slides as one undo step, and Undo takes the fonts back', async () => {
		const deck = await open();
		await deck.pick('themeFonts', 'arial');
		expect(deck.editor.theme?.fontScheme?.majorFont?.latin).toBe('Arial');
		expect(deck.editor.theme?.name).toBeUndefined();
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, major: 'Arial' });
		deck.editor.undo();
		expect(deck.editor.theme).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});

	it('a Colors pick that changes nothing on the slides is still one undo step', async () => {
		const deck = await open();
		// The sample deck's slides use no scheme colour, so only the theme records this pick.
		await deck.pick('themeColors', 'office');
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, accent1: '156082' });
		deck.editor.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});

	it('the theme editor and Browse Themes are undone in the theme part too', async () => {
		const deck = await open();
		const facet = GALLERY_THEME_PRESETS.find((preset) => preset.id === 'facet')!;
		await deck.host.applyThemePreset(facet);
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
		deck.editor.undo();
		expect(deck.editor.theme).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		// The theme editor (DeckThemeEditor.svelte) applies its edit as a custom preset.
		deck.host.publishTheme(
			await applyThemePreset(deck.editor, deck.handler, deck.editor.theme, {
				id: 'custom',
				name: 'Edited',
				colorScheme: { ...deck.loaded!.colorScheme!, accent1: '#123456' },
				fontScheme: deck.loaded!.fontScheme!,
			}),
		);
		await expect(deck.saved()).resolves.toStrictEqual({
			...OFFICE,
			name: 'Edited',
			accent1: '123456',
		});
		deck.editor.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});
});

/**
 * Presentation properties > Theme on the two-themes deck (slides 1 and 2 on master 1, "Tide", accent
 * 1 blue; slide 3 on master 2, "Ember", accent 1 red), through the inspector's deck actions.
 */
async function openTwoThemes(options?: TwoThemesDeckOptions) {
	const bytes = await buildTwoThemesDeck(options);
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer.slice(0) as ArrayBuffer);
	// Slide 2 is on screen.
	const current = 1;
	const editor = new EditorState({
		getCurrent: () => current,
		getHandler: () => handler,
		getThemeOptions: () => data.themeOptions ?? [],
		onThemeRestored: (theme) => {
			loader.presentationTheme = theme;
		},
	});
	editor.editable = true;
	editor.setSlides(data.slides, data.slideMasters);
	editor.theme = data.theme;
	// Only what the deck actions read of the loader for a theme switch.
	const loader = {
		handler,
		presentationTheme: data.theme,
		colorScheme: data.theme?.colorScheme,
		themeOptions: data.themeOptions ?? [],
	} as unknown as PresentationLoader;
	const actions = createInspectorDeckActions({ loader, editor });
	const accents = () =>
		editor.slides.map((slide) =>
			slide.elements
				.find((element) => element.name === TWO_THEMES.shapeName)
				?.shapeStyle?.fillColor?.replace('#', '')
				.toUpperCase(),
		);
	/** The masters' theme relationships and each part's accent 1, as File > Save writes them now. */
	const saved = async () => {
		const zip = await JSZip.loadAsync(await editor.save());
		const text = async (path: string) => (await zip.file(path)?.async('string')) ?? '';
		const master = async (n: number) =>
			/Target="\.\.\/theme\/(theme\d\.xml)"/u.exec(
				await text(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`),
			)?.[1];
		const accent = async (path: string) =>
			/<a:accent1>\s*<a:srgbClr val="([0-9A-F]{6})"/iu.exec(await text(path))?.[1]?.toUpperCase();
		return {
			masters: [await master(1), await master(2)],
			accents: [await accent(TWO_THEMES.theme1.path), await accent(TWO_THEMES.theme2.path)],
		};
	};
	/** Each theme part's heading face, as File > Save writes it now. */
	const faces = async () => themePartHeadingFaces(await editor.save());
	return { editor, loader, actions, accents, saved, faces };
}

describe('svelte theme by path', () => {
	const { theme1, theme2 } = TWO_THEMES;
	const LOADED = {
		masters: ['theme1.xml', 'theme2.xml'],
		accents: [theme1.colors.accent1, theme2.colors.accent1],
	};

	it('every master on theme 2: in place, one undo step, and Undo puts master 1 back', async () => {
		const deck = await openTwoThemes();
		const selected = deck.editor.slides[1]!.elements.find(
			(element) => element.name === TWO_THEMES.shapeName,
		)!;
		deck.editor.select(selected.id);
		const third = deck.editor.slides[2];
		await deck.actions.applyThemeByPath(theme2.path, true);
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.editor.slides[2]).toBe(third);
		expect(deck.editor.selectedElement?.id).toBe(selected.id);
		expect(deck.editor.slideMasters.map((master) => master.themePath)).toStrictEqual([
			theme2.path,
			theme2.path,
		]);
		expect(deck.editor.theme?.name).toBe('Ember');
		expect(deck.loader.presentationTheme?.name).toBe('Ember');
		expect(deck.editor.canUndo).toBeTruthy();
		await expect(deck.saved()).resolves.toStrictEqual({
			...LOADED,
			masters: ['theme2.xml', 'theme2.xml'],
		});

		deck.editor.undo();
		expect(deck.accents()).toStrictEqual([
			theme1.colors.accent1,
			theme1.colors.accent1,
			theme2.colors.accent1,
		]);
		expect(deck.editor.slideMasters.map((master) => master.themePath)).toStrictEqual([
			theme1.path,
			theme2.path,
		]);
		expect(deck.editor.theme?.colorScheme?.accent1?.replace('#', '')).toBe(theme1.colors.accent1);
		expect(deck.editor.canUndo).toBeFalsy();
		// Undo went back to the old part: neither part was rewritten.
		await expect(deck.saved()).resolves.toStrictEqual(LOADED);

		deck.editor.redo();
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		await expect(deck.saved()).resolves.toStrictEqual({
			...LOADED,
			masters: ['theme2.xml', 'theme2.xml'],
		});
	});

	it("every master on the first master's part: only slide 3 changes", async () => {
		const deck = await openTwoThemes();
		const [first, second] = deck.editor.slides;
		await deck.actions.applyThemeByPath(theme1.path, true);
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme1.colors.accent1));
		expect(deck.editor.slides[0]).toBe(first);
		expect(deck.editor.slides[1]).toBe(second);
		deck.editor.undo();
		await expect(deck.saved()).resolves.toStrictEqual(LOADED);
	});

	it('parts that differ in their script faces: Undo writes neither part', async () => {
		const deck = await openTwoThemes({ scriptFaces: true });
		await deck.actions.applyThemeByPath(theme2.path, false);
		deck.editor.undo();
		await expect(deck.faces()).resolves.toStrictEqual({
			[theme1.path]: theme1.major,
			[theme2.path]: theme2.major,
		});
		await expect(deck.saved()).resolves.toStrictEqual(LOADED);
	});
});
