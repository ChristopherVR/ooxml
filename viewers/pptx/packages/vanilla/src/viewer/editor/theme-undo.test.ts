/**
 * Undo and Redo of a theme change take the theme back as well as the slides: in the store's theme
 * fields and in the handler's theme part a save writes.
 *
 * The real handler over `sample-deck.pptx` (an Office theme, "Office Theme", accent 1 4472C4,
 * headings in Calibri Light), wired to the real editor operations, gallery actions, deck actions
 * and Browse Themes the way the viewer wires them.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { applyRibbonGalleryItem, partitionTemplateElements } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { buildThemeColorMap, PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { EditorController } from '../editor';
import type { LoadingController } from '../loading-controller';
import { applyPresentationThemePreset } from '../presentation-theme-controller';
import { createInitialViewerState, createStore } from '../state';
import { createDeckActions } from './editor-deck-actions';
import { createGalleryActions } from './editor-gallery-actions';
import { createEditorOps } from './editor-operations';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

const OFFICE = { name: 'Office Theme', accent1: '4472C4', major: 'Calibri Light' };
const FACET = { name: 'Facet', accent1: '90C226', major: 'Trebuchet MS' };

async function open() {
	const bytes = readFileSync(DECK);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const store = createStore({
		...createInitialViewerState(),
		editable: true,
		slides: data.slides,
		slideMasters: data.slideMasters ?? [],
		themeOptions: data.themeOptions ?? [],
		colorScheme: data.theme?.colorScheme,
		fontScheme: data.theme?.fontScheme,
		formatScheme: data.theme?.formatScheme,
		themeName: data.theme?.name,
	});
	const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: () => {} });
	const deck = createDeckActions({ store, ops, getHandler: () => handler });
	const gallery = createGalleryActions({ store, ops, getHandler: () => handler });
	const theme = () => {
		const state = store.get();
		return { colorScheme: state.colorScheme, fontScheme: state.fontScheme, name: state.themeName };
	};
	const loaded = theme();
	/** Pick a tile the way the ribbon does: the shared apply result, then the viewer's dispatch. */
	const pick = async (id: RibbonGalleryId, itemId: string) => {
		const current = theme();
		const result = applyRibbonGalleryItem(id, itemId, {
			element: null,
			theme: current,
			themeColorMap: current.colorScheme ? buildThemeColorMap(current.colorScheme) : undefined,
		});
		expect(result).not.toBeNull();
		gallery.applyRibbonGalleryResult(result!);
		await expect.poll(() => ops.canUndo()).toBeTruthy();
	};
	/** The theme part as File > Save writes it now. */
	const saved = async () => {
		const zip = await JSZip.loadAsync(await ops.save());
		const xml = (await zip.file('ppt/theme/theme1.xml')?.async('string')) ?? '';
		const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
		return {
			name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
			accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u).toUpperCase(),
			major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
		};
	};
	return { store, ops, deck, handler, theme, loaded, pick, saved };
}

describe('vanilla theme undo', () => {
	it('a Themes pick writes the theme part, and Undo and Redo take the whole theme', async () => {
		const deck = await open();
		await deck.pick('themes', 'facet');
		expect(deck.theme().colorScheme?.accent1).toBe('#90C226');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);

		deck.ops.undo();
		expect(deck.theme()).toStrictEqual(deck.loaded);
		expect(deck.ops.canUndo()).toBeFalsy();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		deck.ops.redo();
		expect(deck.theme().name).toBe('Facet');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
	});

	it('a Fonts pick writes only the fonts and keeps the name, and Undo takes them back', async () => {
		const deck = await open();
		await deck.pick('themeFonts', 'arial');
		expect(deck.theme().name).toBe(deck.loaded.name);
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, major: 'Arial' });
		deck.ops.undo();
		expect(deck.theme()).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});

	it('a Colors pick keeps the name, and Undo takes the colour scheme back', async () => {
		const deck = await open();
		// The sample deck's slides use no scheme colour, so only the theme records this pick.
		await deck.pick('themeColors', 'office');
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, accent1: '156082' });
		deck.ops.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});

	it('the theme editor and Browse Themes are undone in the theme part too', async () => {
		const deck = await open();
		await deck.deck.applyThemeEdit({
			name: 'Edited',
			colorScheme: { ...deck.loaded.colorScheme!, accent1: '#123456' },
			fontScheme: deck.loaded.fontScheme!,
		});
		await expect(deck.saved()).resolves.toStrictEqual({
			...OFFICE,
			name: 'Edited',
			accent1: '123456',
		});
		deck.ops.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		const editor = {
			commitSlides: (slides: never, current: number) => {
				deck.ops.pushHistory();
				deck.store.set({ slides, currentSlide: current });
				deck.ops.commitChange();
			},
		} as unknown as EditorController;
		await applyPresentationThemePreset({
			presetId: 'facet',
			loading: { getHandler: () => deck.handler } as unknown as LoadingController,
			store: deck.store,
			editor,
		});
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
		deck.ops.undo();
		expect(deck.theme()).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});
});

/** What the tests read of the two-themes deck (`TWO_THEMES` in its definition). */
interface TwoThemesDeck {
	theme1: { path: string; colors: { accent1: string } };
	theme2: { path: string; colors: { accent1: string } };
	shapeName: string;
}

/**
 * The two-themes deck from its one definition, next to the core tests. Loaded at run time rather
 * than imported, because this package's type check (`rootDir: src`) cannot follow an import out of
 * the package.
 */
const TWO_THEMES_HELPER = resolve(
	process.cwd(),
	'../../../../src/core/pptx/__tests__/helpers/two-themes-deck.ts',
);

/**
 * Presentation properties > Theme on the two-themes deck (slides 1 and 2 on master 1, "Tide", accent
 * 1 blue; slide 3 on master 2, "Ember", accent 1 red), through the inspector's deck action.
 */
async function openTwoThemes() {
	const helper = (await import(/* @vite-ignore */ TWO_THEMES_HELPER)) as {
		buildTwoThemesDeck(): Promise<Uint8Array>;
		TWO_THEMES: TwoThemesDeck;
	};
	const bytes = await helper.buildTwoThemesDeck();
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer.slice(0) as ArrayBuffer);
	const partition = partitionTemplateElements(data.slides);
	const store = createStore({
		...createInitialViewerState(),
		editable: true,
		slides: partition.slides,
		templateElementsBySlideId: partition.templateElementsBySlideId,
		currentSlide: 1,
		slideMasters: data.slideMasters ?? [],
		themeOptions: data.themeOptions ?? [],
		colorScheme: data.theme?.colorScheme,
		fontScheme: data.theme?.fontScheme,
		formatScheme: data.theme?.formatScheme,
		themeName: data.theme?.name,
	});
	const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: () => {} });
	const deck = createDeckActions({ store, ops, getHandler: () => handler });
	const shapeName = helper.TWO_THEMES.shapeName;
	const accents = () =>
		store.get().slides.map((slide) => {
			const shape = slide.elements.find((element) => element.name === shapeName) as
				| { shapeStyle?: { fillColor?: string } }
				| undefined;
			return shape?.shapeStyle?.fillColor?.replace('#', '').toUpperCase();
		});
	const masterThemes = () => store.get().slideMasters.map((master) => master.themePath);
	/** The masters' theme relationships and each part's accent 1, as File > Save writes them now. */
	const saved = async () => {
		const zip = await JSZip.loadAsync(await ops.save());
		const text = async (path: string) => (await zip.file(path)?.async('string')) ?? '';
		const master = async (n: number) =>
			/Target="\.\.\/theme\/(theme\d\.xml)"/u.exec(
				await text(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`),
			)?.[1];
		const accent = async (path: string) =>
			/<a:accent1>\s*<a:srgbClr val="([0-9A-F]{6})"/iu.exec(await text(path))?.[1]?.toUpperCase();
		const { theme1, theme2 } = helper.TWO_THEMES;
		return {
			masters: [await master(1), await master(2)],
			accents: [await accent(theme1.path), await accent(theme2.path)],
		};
	};
	return { store, ops, deck, accents, masterThemes, saved, ...helper.TWO_THEMES };
}

describe('vanilla theme by path', () => {
	it('every master on theme 2 re-themes the slides, as one undo step that Undo takes back', async () => {
		const deck = await openTwoThemes();
		const { theme1, theme2 } = deck;
		const loaded = {
			masters: ['theme1.xml', 'theme2.xml'],
			accents: [theme1.colors.accent1, theme2.colors.accent1],
		};
		const selected = deck.store.get().slides[1]!.elements[0]!.id;
		deck.ops.select(selected);
		const third = deck.store.get().slides[2];
		await deck.deck.applyThemeByPath(theme2.path, true);
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.store.get().slides[2]).toBe(third);
		expect(deck.store.get().currentSlide).toBe(1);
		expect(deck.store.get().selectedElementId).toBe(selected);
		expect(deck.masterThemes()).toStrictEqual([theme2.path, theme2.path]);
		expect(deck.store.get().themeName).toBe('Ember');
		expect(deck.store.get().colorScheme?.accent1?.replace('#', '')).toBe(theme2.colors.accent1);
		expect(deck.ops.canUndo()).toBeTruthy();
		await expect(deck.saved()).resolves.toStrictEqual({
			...loaded,
			masters: ['theme2.xml', 'theme2.xml'],
		});

		deck.ops.undo();
		expect(deck.accents()).toStrictEqual([
			theme1.colors.accent1,
			theme1.colors.accent1,
			theme2.colors.accent1,
		]);
		expect(deck.masterThemes()).toStrictEqual([theme1.path, theme2.path]);
		expect(deck.store.get().colorScheme?.accent1?.replace('#', '')).toBe(theme1.colors.accent1);
		expect(deck.ops.canUndo()).toBeFalsy();
		// Undo went back to the old part: neither part was rewritten.
		await expect(deck.saved()).resolves.toStrictEqual(loaded);

		deck.ops.redo();
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.masterThemes()).toStrictEqual([theme2.path, theme2.path]);
	});

	it("every master on the first master's part re-themes only slide 3", async () => {
		const deck = await openTwoThemes();
		const [first, second] = deck.store.get().slides;
		await deck.deck.applyThemeByPath(deck.theme1.path, true);
		expect(deck.accents()).toStrictEqual(Array(3).fill(deck.theme1.colors.accent1));
		expect(deck.store.get().slides[0]).toBe(first);
		expect(deck.store.get().slides[1]).toBe(second);
		expect(deck.masterThemes()).toStrictEqual([deck.theme1.path, deck.theme1.path]);
	});
});
