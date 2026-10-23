// @vitest-environment happy-dom
/**
 * A Design > Themes, Variants > Fonts or Variants > Colors pick applies in place, and Undo and Redo
 * take the theme back with the slides.
 *
 * Applied in place, a pick is one undo step. That step must take the THEME back as well as the
 * slides: the undo stack held only the slides, so after Undo the slides showed the old faces while
 * the handler's theme part (what a save writes) and the viewer's theme kept the pick. A pick that
 * changes nothing on the slides (a Colors pick on a deck with no scheme colours) left no undo step
 * at all.
 *
 * `field-substitution.pptx` has title placeholders that inherit the theme's heading font
 * ("Calibri Light") under an Office theme named "Office Theme".
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { applyRibbonGalleryItem } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { act } from 'react';
import { beforeAll, describe, expect, it } from 'vitest';

import {
	betaTitleFont,
	flush,
	flushUntil,
	live,
	session,
	withHarness,
} from './theme-apply-harness.test-support';

let deckBytes: Uint8Array;

beforeAll(() => {
	deckBytes = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/field-substitution.pptx')),
	);
});

/** A pick that applied in place: same content, same slide and selection, dirty, undoable. */
const KEPT_SESSION = {
	reloaded: false,
	selectionCleared: false,
	slideOnScreen: 2,
	dirty: true,
	canUndo: true,
};

interface SavedTheme {
	name: string;
	accent1: string;
	major: string;
	minor: string;
}

/** The theme as a save writes it: name, accent 1 and the heading and body faces. */
async function savedTheme(): Promise<SavedTheme> {
	return themePartOf(await live().save());
}

async function themePartOf(bytes: Uint8Array): Promise<SavedTheme> {
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/theme/theme1.xml')?.async('string')) ?? '';
	const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
	return {
		name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
		accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u),
		major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
		minor: read(/<a:minorFont>\s*<a:latin typeface="([^"]*)"/u),
	};
}

const LOADED_THEME = {
	name: 'Office Theme',
	accent1: '4472C4',
	major: 'Calibri Light',
	minor: 'Calibri',
};

const FACET_THEME = {
	name: 'Facet',
	accent1: '90C226',
	major: 'Trebuchet MS',
	minor: 'Trebuchet MS',
};

/** The same four facts read from the viewer's own theme (the loaded one has no name). */
function viewerTheme() {
	const theme = live().theme;
	return {
		name: theme?.name ?? 'none',
		accent1: theme?.colorScheme?.accent1?.replace('#', '') ?? 'missing',
		major: theme?.fontScheme?.majorFont?.latin ?? 'missing',
		minor: theme?.fontScheme?.minorFont?.latin ?? 'missing',
	};
}

/** The theme in the viewer and in the file, so a failure shows both at once. */
async function bothThemes() {
	return { viewer: viewerTheme(), file: await savedTheme() };
}

const LOADED_BOTH = { viewer: { ...LOADED_THEME, name: 'none' }, file: LOADED_THEME };

async function openOnSlideTwo(): Promise<{ parses: number; selectionClears: number }> {
	await flushUntil(() => live().parses >= 1 && !live().loading);
	await act(async () => {
		live().setActiveSlideIndex(1);
	});
	await flush();
	expect(betaTitleFont()).toBe('Calibri Light');
	expect(live().history.canUndo).toBeFalsy();
	await expect(savedTheme()).resolves.toStrictEqual(LOADED_THEME);
	return { parses: live().parses, selectionClears: live().selectionClears };
}

/** Pick a tile the way `RibbonGallery` does: the shared apply result, then the host's dispatch. */
async function pick(gallery: RibbonGalleryId, itemId: string, expectedFont: string): Promise<void> {
	const result = applyRibbonGalleryItem(gallery, itemId, live().gallery.context);
	expect(result).not.toBeNull();
	await act(async () => {
		live().gallery.dispatch(result!);
	});
	await flushUntil(() => betaTitleFont() === expectedFont && !live().loading);
	// Long enough for a reload the pick may have started to land.
	await flush(150);
}

/** Undo (or Redo) one step and let the theme part writes behind it settle. */
async function step(direction: 'undo' | 'redo'): Promise<void> {
	await act(async () => {
		if (direction === 'undo') {
			live().history.handleUndo();
		} else {
			live().history.handleRedo();
		}
	});
	await flush(50);
}

describe('design gallery theme picks apply in place', () => {
	it('a Themes pick keeps the session, and Undo and Redo take the whole theme', async () => {
		await withHarness(deckBytes, async () => {
			const baseline = await openOnSlideTwo();
			const loaded = live().theme;
			await pick('themes', 'facet', 'Trebuchet MS');

			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			const picked = live().theme;
			expect(picked?.name).toBe('Facet');
			expect(picked?.colorScheme?.accent1).toBe('#90C226');
			const facet = FACET_THEME;
			await expect(savedTheme()).resolves.toStrictEqual(facet);

			await step('undo');
			expect(betaTitleFont()).toBe('Calibri Light');
			expect(live().history.canUndo).toBeFalsy();
			expect(live().activeSlideIndex).toBe(1);
			await expect(bothThemes()).resolves.toStrictEqual(LOADED_BOTH);
			expect(live().theme).toStrictEqual(loaded);

			await step('redo');
			expect(betaTitleFont()).toBe('Trebuchet MS');
			await expect(bothThemes()).resolves.toStrictEqual({ viewer: facet, file: facet });
			expect(live().theme).toStrictEqual(picked);
		});
	});

	it('a Fonts pick keeps the session, the colours and the name, and Undo takes the fonts', async () => {
		await withHarness(deckBytes, async () => {
			const baseline = await openOnSlideTwo();
			const loaded = live().theme;
			await pick('themeFonts', 'arial', 'Arial');

			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			expect(live().theme?.fontScheme?.majorFont?.latin).toBe('Arial');
			expect(live().theme?.colorScheme).toStrictEqual(loaded?.colorScheme);
			expect(live().theme?.name).toBe(loaded?.name);
			const arial = { ...LOADED_THEME, major: 'Arial', minor: 'Arial' };
			await expect(savedTheme()).resolves.toStrictEqual(arial);

			await step('undo');
			expect(betaTitleFont()).toBe('Calibri Light');
			expect(live().history.canUndo).toBeFalsy();
			expect(live().activeSlideIndex).toBe(1);
			await expect(bothThemes()).resolves.toStrictEqual(LOADED_BOTH);
			expect(live().theme).toStrictEqual(loaded);

			await step('redo');
			expect(betaTitleFont()).toBe('Arial');
			await expect(bothThemes()).resolves.toStrictEqual({
				viewer: { ...arial, name: 'none' },
				file: arial,
			});
		});
	});

	it('a Colors pick, in place already, now also takes its theme part back on Undo', async () => {
		await withHarness(deckBytes, async () => {
			const baseline = await openOnSlideTwo();
			// Nothing on these slides uses a scheme colour, so only the theme records the pick.
			await pick('themeColors', 'office', 'Calibri Light');

			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			const office = { ...LOADED_THEME, accent1: '156082' };
			await expect(bothThemes()).resolves.toStrictEqual({
				viewer: { ...office, name: 'none' },
				file: office,
			});

			await step('undo');
			expect(live().history.canUndo).toBeFalsy();
			await expect(bothThemes()).resolves.toStrictEqual(LOADED_BOTH);
		});
	});

	it('the Browse Themes overlay commits through the theme handlers, and Undo takes it back too', async () => {
		await withHarness(deckBytes, async () => {
			const baseline = await openOnSlideTwo();
			const loaded = live().theme;
			const facet = applyRibbonGalleryItem('themes', 'facet', live().gallery.context);
			if (facet?.kind !== 'theme') {
				throw new Error('the Themes gallery has no Facet tile');
			}
			await act(async () => {
				await live().themeHandlers.handleApplyThemeData(
					facet.colorScheme,
					facet.fontScheme,
					facet.name,
				);
			});
			await flushUntil(() => betaTitleFont() === 'Trebuchet MS');
			await flush(150);
			expect(session(baseline)).toStrictEqual(KEPT_SESSION);

			await step('undo');
			expect(betaTitleFont()).toBe('Calibri Light');
			await expect(bothThemes()).resolves.toStrictEqual(LOADED_BOTH);
			expect(live().theme).toStrictEqual(loaded);
		});
	});
});
