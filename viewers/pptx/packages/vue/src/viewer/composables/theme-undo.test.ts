/**
 * Undo and Redo of a theme change take the theme back as well as the slides: in the viewer's theme
 * and colour map, and in the handler's theme part a save writes.
 *
 * The real handler over `sample-deck.pptx` (an Office theme, "Office Theme", accent 1 4472C4,
 * headings in Calibri Light), wired to the real history, Design gallery commit and theme editing
 * the way `PowerPointViewer.vue` wires them.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { applyRibbonGalleryItem, loadedThemeName, saveDeckWithPassword } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { effectScope, ref, shallowRef } from 'vue';

import { useDeckStyleHost } from './useDeckStyleHost';
import { useEditorHistory } from './useEditorHistory';
import { dispatchRibbonGalleryResult } from './useRibbonGalleryHost';
import { useThemeEditing } from './useThemeEditing';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

const OFFICE = { name: 'Office Theme', accent1: '4472C4', major: 'Calibri Light' };
const FACET = { name: 'Facet', accent1: '90C226', major: 'Trebuchet MS' };

async function open() {
	const bytes = readFileSync(DECK);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const scope = effectScope();
	const slides = ref<PptxSlide[]>(data.slides);
	const templateElementsBySlideId = shallowRef<Record<string, PptxElement[]>>({});
	const pptxTheme = ref<PptxTheme | undefined>(data.theme);
	const themeColorMap = ref<Record<string, string> | undefined>(data.themeColorMap);
	const handlerRef = shallowRef<PptxHandler | null>(handler);
	const loadedName = () => loadedThemeName(data.themeOptions, data.slideMasters?.[0]?.themePath);
	const parts = scope.run(() => {
		const history = useEditorHistory(slides, templateElementsBySlideId, {
			theme: { theme: pptxTheme, themeColorMap, handler: handlerRef, loadedThemeName: loadedName },
		});
		const deckStyle = useDeckStyleHost({
			slides,
			templateElementsBySlideId,
			pptxTheme,
			themeColorMap,
			handler: handlerRef,
			selectedSlideIds: () => [],
			pushHistory: history.pushHistory,
			held: ref(false),
		});
		const editing = useThemeEditing({
			handler: handlerRef,
			slides,
			pptxTheme,
			themeColorMap,
			pushHistory: history.pushHistory,
			themeGalleryOpen: ref(false),
			themeEditorOpen: ref(false),
			templateElementsBySlideId,
			loadedThemeName: loadedName,
		});
		return { history, deckStyle, editing };
	})!;
	/** Pick a tile the way the ribbon does: the shared apply result, then the viewer's dispatch. */
	const pick = async (gallery: RibbonGalleryId, itemId: string) => {
		const result = applyRibbonGalleryItem(gallery, itemId, {
			element: null,
			theme: pptxTheme.value,
			themeColorMap: themeColorMap.value,
		});
		expect(result).not.toBeNull();
		dispatchRibbonGalleryResult(result!, {
			updateElement: () => {},
			// What PowerPointViewer.vue gives the galleries (its `applyAiTheme`).
			applyTheme: (updates) => void parts.editing.applyThemeUpdates(updates),
			deckStyle: parts.deckStyle,
		});
		await expect.poll(() => parts.history.canUndo.value).toBeTruthy();
	};
	/** The theme part as File > Save writes it now. */
	const saved = async () => {
		const zip = await JSZip.loadAsync(await saveDeckWithPassword(handler, slides.value, {}, null));
		const xml = (await zip.file('ppt/theme/theme1.xml')?.async('string')) ?? '';
		const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
		return {
			name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
			accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u).toUpperCase(),
			major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
		};
	};
	const viewer = () => ({
		accent1: pptxTheme.value?.colorScheme?.accent1,
		major: pptxTheme.value?.fontScheme?.majorFont?.latin,
		mapAccent1: themeColorMap.value?.accent1?.replace('#', ''),
	});
	return { ...parts, pptxTheme, themeColorMap, scope, saved, viewer, pick, loaded: data.theme };
}

describe('vue theme undo', () => {
	it('a Themes pick writes the theme part, and Undo and Redo take the whole theme', async () => {
		const deck = await open();
		const loadedColorMap = { ...deck.themeColorMap.value };
		await deck.pick('themes', 'facet');
		expect(deck.viewer()).toStrictEqual({
			accent1: '#90C226',
			major: 'Trebuchet MS',
			mapAccent1: '90C226',
		});
		await expect(deck.saved()).resolves.toStrictEqual(FACET);

		deck.history.undo();
		expect(deck.pptxTheme.value).toStrictEqual(deck.loaded);
		expect(deck.themeColorMap.value).toStrictEqual(loadedColorMap);
		expect(deck.history.canUndo.value).toBeFalsy();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		deck.history.redo();
		expect(deck.viewer().accent1).toBe('#90C226');
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
		deck.scope.stop();
	});

	it('a Fonts pick writes only the fonts and keeps the name, and Undo takes them back', async () => {
		const deck = await open();
		await deck.pick('themeFonts', 'arial');
		expect(deck.pptxTheme.value?.name).toBeUndefined();
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, major: 'Arial' });
		deck.history.undo();
		expect(deck.pptxTheme.value).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
		deck.scope.stop();
	});

	it('a Colors pick that changes nothing on the slides is still one undo step', async () => {
		const deck = await open();
		// The sample deck's slides use no scheme colour, so only the theme records this pick.
		await deck.pick('themeColors', 'office');
		await expect(deck.saved()).resolves.toStrictEqual({ ...OFFICE, accent1: '156082' });
		deck.history.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
		deck.scope.stop();
	});

	it('the theme editor and Browse Themes write the theme part, and Undo takes it back', async () => {
		const deck = await open();
		await deck.editing.applyThemeEdit({
			name: 'Edited',
			colorScheme: { ...deck.loaded!.colorScheme!, accent1: '#123456' },
			fontScheme: deck.loaded!.fontScheme!,
		});
		await expect(deck.saved()).resolves.toStrictEqual({
			...OFFICE,
			name: 'Edited',
			accent1: '123456',
		});
		deck.history.undo();
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);

		const facet = applyRibbonGalleryItem('themes', 'facet', { element: null });
		if (facet?.kind !== 'theme') {
			throw new Error('no Facet tile');
		}
		deck.editing.applyThemePreset({ id: 'facet', ...facet });
		await expect(deck.saved()).resolves.toStrictEqual(FACET);
		deck.history.undo();
		expect(deck.pptxTheme.value).toStrictEqual(deck.loaded);
		await expect(deck.saved()).resolves.toStrictEqual(OFFICE);
		deck.scope.stop();
	});
});
