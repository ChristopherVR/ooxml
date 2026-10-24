/**
 * Undo and Redo of a theme change take the theme back as well as the slides: in the loader's theme
 * and colour map, and in the handler's theme part a save writes.
 *
 * The real handler over `sample-deck.pptx` (an Office theme, "Office Theme", accent 1 4472C4,
 * headings in Calibri Light), wired to the real editor state, loader, Design gallery commit and
 * theme gallery services the way the viewer component provides them.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import JSZip from 'jszip';
import { applyRibbonGalleryItem, GALLERY_THEME_PRESETS, saveDeckWithPassword } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import {
	buildTwoThemesDeck,
	TWO_THEMES,
} from '../../../../../../src/core/pptx/__tests__/helpers/two-themes-deck';
import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { dispatchGalleryResult } from './ribbon-gallery-helpers';
import { buildSaveSlides } from './template-mode';
import { ThemeSelectorCardComponent } from './theme-selector-card.component';
import { ViewerDeckStyleService } from './viewer-deck-style.service';
import { ViewerThemeGalleryService } from './viewer-theme-gallery.service';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

const OFFICE = { name: 'Office Theme', accent1: '4472C4', major: 'Calibri Light' };
const FACET = { name: 'Facet', accent1: '90C226', major: 'Trebuchet MS' };

async function open() {
	const bytes = readFileSync(DECK);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const destroyRefStub: Pick<DestroyRef, 'onDestroy'> = { onDestroy: () => () => {} };
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: destroyRefStub },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			{ provide: EditorStateService, useClass: EditorStateService },
			{ provide: LoadContentService, useClass: LoadContentService },
			{ provide: ViewerThemeGalleryService, useClass: ViewerThemeGalleryService },
			{ provide: ViewerDeckStyleService, useClass: ViewerDeckStyleService },
		],
	});
	const get = <T>(token: new (...args: never[]) => T): T =>
		runInInjectionContext(injector, () => injector.get(token));
	const editor = get(EditorStateService);
	const loader = get(LoadContentService);
	const themes = get(ViewerThemeGalleryService);
	const deckStyle = get(ViewerDeckStyleService);
	vi.spyOn(loader, 'getHandler').mockReturnValue(handler);
	loader.theme.set(data.theme);
	loader.themeColorMap.set(data.themeColorMap);
	loader.themeOptions.set(data.themeOptions ?? []);
	loader.slideMasters.set(data.slideMasters ?? []);
	editor.setSlides(data.slides);
	/** Pick a tile the way the ribbon does: the shared apply result, then the viewer's dispatch. */
	const pick = async (id: RibbonGalleryId, itemId: string) => {
		const result = applyRibbonGalleryItem(id, itemId, {
			element: null,
			theme: loader.theme(),
			themeColorMap: loader.themeColorMap(),
		});
		expect(
			dispatchGalleryResult(result, { editor, slideIndex: 0, themes, deckStyle: deckStyle.host }),
		).toBeTruthy();
		await expect.poll(() => editor.canUndo()).toBeTruthy();
	};
	/** The theme part as File > Save writes it now. */
	const saved = async () => {
		const slides = buildSaveSlides(editor.slides(), editor.templateElementsBySlideId());
		const zip = await JSZip.loadAsync(await saveDeckWithPassword(handler, slides, {}, null));
		const xml = (await zip.file('ppt/theme/theme1.xml')?.async('string')) ?? '';
		const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
		return {
			name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
			accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u).toUpperCase(),
			major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
		};
	};
	return { editor, loader, themes, pick, saved, loaded: data.theme, loadedMap: data.themeColorMap };
}

describe('angular theme undo', () => {
	it('a Themes pick writes the theme part, and Undo and Redo take the whole theme', async () => {
		const deck = await open();
		await deck.pick('themes', 'facet');
		expect(deck.loader.theme()?.colorScheme?.accent1).toBe('#90C226');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);

		deck.editor.undo();
		expect(deck.loader.theme()).toStrictEqual(deck.loaded);
		expect(deck.loader.themeColorMap()).toStrictEqual(deck.loadedMap);
		expect(deck.editor.canUndo()).toBeFalsy();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		deck.editor.redo();
		expect(deck.loader.theme()?.name).toBe('Facet');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
	});

	it('a Fonts pick writes only the fonts and keeps the name, and Undo takes them back', async () => {
		const deck = await open();
		await deck.pick('themeFonts', 'arial');
		expect(deck.loader.theme()?.name).toBeUndefined();
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, major: 'Arial' });
		deck.editor.undo();
		expect(deck.loader.theme()).toStrictEqual(deck.loaded);
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
		await deck.themes.applyThemeEditor({
			name: 'Edited',
			colorScheme: { ...deck.loaded!.colorScheme!, accent1: '#123456' },
			fontScheme: deck.loaded!.fontScheme!,
		});
		await expect(deck.saved()).resolves.toStrictEqual({
			...OFFICE,
			name: 'Edited',
			accent1: '123456',
		});
		deck.editor.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		deck.themes.applyThemePreset(GALLERY_THEME_PRESETS.find((preset) => preset.id === 'facet')!);
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
		deck.editor.undo();
		expect(deck.loader.theme()).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
	});
});

/**
 * Presentation properties > Theme on the two-themes deck (slides 1 and 2 on master 1, "Tide", accent
 * 1 blue; slide 3 on master 2, "Ember", accent 1 red), through the inspector's theme card.
 */
async function openTwoThemes() {
	const bytes = await buildTwoThemesDeck();
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer.slice(0) as ArrayBuffer);
	const destroyRefStub: Pick<DestroyRef, 'onDestroy'> = { onDestroy: () => () => {} };
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: destroyRefStub },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			{ provide: EditorStateService, useClass: EditorStateService },
			{ provide: LoadContentService, useClass: LoadContentService },
		],
	});
	const editor = runInInjectionContext(injector, () => injector.get(EditorStateService));
	const loader = runInInjectionContext(injector, () => injector.get(LoadContentService));
	vi.spyOn(loader, 'getHandler').mockReturnValue(handler);
	loader.theme.set(data.theme);
	loader.themeColorMap.set(data.themeColorMap);
	loader.themeOptions.set(data.themeOptions ?? []);
	loader.slideMasters.set(data.slideMasters ?? []);
	editor.setSlides(data.slides);
	const card = runInInjectionContext(
		injector,
		() => new ThemeSelectorCardComponent(),
	) as unknown as {
		onSelect(event: Event): void;
		apply(allMasters: boolean): Promise<void>;
	};
	/** Pick `themePath` in the card's theme list, then press Apply First / All Masters. */
	const apply = async (themePath: string, allMasters: boolean) => {
		card.onSelect({ target: { value: themePath } } as unknown as Event);
		await card.apply(allMasters);
	};
	const accents = () =>
		editor.slides().map((slide) =>
			slide.elements
				.find((element) => element.name === TWO_THEMES.shapeName)
				?.shapeStyle?.fillColor?.replace('#', '')
				.toUpperCase(),
		);
	const masterThemes = () => loader.slideMasters().map((master) => master.themePath);
	/** The masters' theme relationships and each part's accent 1, as File > Save writes them now. */
	const saved = async () => {
		const slides = buildSaveSlides(editor.slides(), editor.templateElementsBySlideId());
		const zip = await JSZip.loadAsync(await saveDeckWithPassword(handler, slides, {}, null));
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
	return { editor, loader, apply, accents, masterThemes, saved };
}

describe('angular theme by path', () => {
	const { theme1, theme2 } = TWO_THEMES;
	const LOADED = {
		masters: ['theme1.xml', 'theme2.xml'],
		accents: [theme1.colors.accent1, theme2.colors.accent1],
	};

	it('every master on theme 2 re-themes the slides, as one undo step that Undo takes back', async () => {
		const deck = await openTwoThemes();
		const selected = deck.editor.slides()[1]!.elements[0]!.id;
		deck.editor.selectedIds.set([selected]);
		const third = deck.editor.slides()[2];
		await deck.apply(theme2.path, true);
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.editor.slides()[2]).toStrictEqual(third);
		expect(deck.editor.selectedIds()).toStrictEqual([selected]);
		expect(deck.masterThemes()).toStrictEqual([theme2.path, theme2.path]);
		expect(deck.loader.theme()?.name).toBe('Ember');
		expect(deck.loader.themeColorMap()?.accent1?.replace('#', '')).toBe(theme2.colors.accent1);
		expect(deck.editor.dirty()).toBeTruthy();
		expect(deck.editor.canUndo()).toBeTruthy();
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
		expect(deck.masterThemes()).toStrictEqual([theme1.path, theme2.path]);
		expect(deck.loader.theme()?.colorScheme?.accent1?.replace('#', '')).toBe(theme1.colors.accent1);
		expect(deck.editor.canUndo()).toBeFalsy();
		// Undo went back to the old part: neither part was rewritten.
		await expect(deck.saved()).resolves.toStrictEqual(LOADED);

		deck.editor.redo();
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.masterThemes()).toStrictEqual([theme2.path, theme2.path]);
	});

	it('the first master alone on theme 2 re-themes only its slides', async () => {
		const deck = await openTwoThemes();
		await deck.apply(theme2.path, false);
		expect(deck.accents()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
		expect(deck.masterThemes()).toStrictEqual([theme2.path, theme2.path]);
		deck.editor.undo();
		await expect(deck.saved()).resolves.toStrictEqual(LOADED);
	});
});
