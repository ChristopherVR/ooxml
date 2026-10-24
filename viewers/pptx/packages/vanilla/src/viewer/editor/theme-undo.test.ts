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
import { applyRibbonGalleryItem } from 'ooxml-ui/pptx';
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
