// @vitest-environment happy-dom
/**
 * Presentation properties > Theme (pick a theme part already in the package, for the first master
 * or every master) applies in place and is one undo step.
 *
 * `handleApplyTheme`, `handleSwitchTheme` and `handleApplyThemeToPresentation` serialised the deck
 * and loaded it again: the view went back to slide 1, the selection was cleared and Undo was gone.
 *
 * The two-themes deck: slides 1 and 2 on master 1 ("Tide", accent 1 blue), slide 3 on master 2
 * ("Ember", accent 1 red); every slide has an `accent1` rectangle.
 */
import JSZip from 'jszip';
import { act } from 'react';
import { describe, expect, it } from 'vitest';

import {
	buildTwoThemesDeck,
	themePartHeadingFaces,
	TWO_THEMES,
} from '../../../../../../../src/core/pptx/__tests__/helpers/two-themes-deck';
import { flush, flushUntil, live, session, withHarness } from './theme-apply-harness.test-support';

const { theme1, theme2, shapeName } = TWO_THEMES;

/** The `accent1` rectangle's fill on each slide, as drawn now. */
function accentFills(): (string | undefined)[] {
	return live().slides.map((slide) => {
		const shape = slide.elements.find((element) => element.name === shapeName);
		return shape?.shapeStyle?.fillColor?.replace('#', '').toUpperCase();
	});
}

const masterThemes = () => live().slideMasters.map((master) => master.themePath);

/** The masters' theme parts and accent 1 of each part, as a save writes them now. */
async function saved() {
	const zip = await JSZip.loadAsync(await live().save());
	const text = (path: string) => zip.file(path)?.async('string') ?? Promise.resolve('');
	const themeOf = async (master: number) =>
		/Target="\.\.\/theme\/(theme\d\.xml)"/u.exec(
			await text(`ppt/slideMasters/_rels/slideMaster${master}.xml.rels`),
		)?.[1];
	const accentOf = async (path: string) =>
		/<a:accent1>\s*<a:srgbClr val="([0-9A-F]{6})"/iu.exec(await text(path))?.[1]?.toUpperCase();
	return {
		masters: [await themeOf(1), await themeOf(2)],
		accents: [await accentOf(theme1.path), await accentOf(theme2.path)],
	};
}

const LOADED = {
	masters: ['theme1.xml', 'theme2.xml'],
	accents: [theme1.colors.accent1, theme2.colors.accent1],
};

async function openOnSlideTwo() {
	await flushUntil(() => live().parses >= 1 && !live().loading);
	await act(async () => {
		live().setActiveSlideIndex(1);
	});
	await flush();
	expect(accentFills()).toStrictEqual([
		theme1.colors.accent1,
		theme1.colors.accent1,
		theme2.colors.accent1,
	]);
	expect(live().history.canUndo).toBeFalsy();
	return { parses: live().parses, selectionClears: live().selectionClears };
}

const KEPT_SESSION = {
	reloaded: false,
	selectionCleared: false,
	slideOnScreen: 2,
	dirty: true,
	canUndo: true,
};

describe('react theme by path', () => {
	it('every master on theme 2: in place, one undo step, and Undo puts master 1 back', async () => {
		await withHarness(await buildTwoThemesDeck(), async () => {
			const baseline = await openOnSlideTwo();
			await act(async () => {
				await live().themeHandlers.handleApplyTheme(theme2.path, true);
			});
			await flush();
			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			expect(accentFills()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
			expect(masterThemes()).toStrictEqual([theme2.path, theme2.path]);
			expect(live().theme?.name).toBe('Ember');
			await expect(saved()).resolves.toStrictEqual({
				...LOADED,
				masters: ['theme2.xml', 'theme2.xml'],
			});

			await act(async () => {
				live().history.handleUndo();
			});
			await flush();
			expect(accentFills()).toStrictEqual([
				theme1.colors.accent1,
				theme1.colors.accent1,
				theme2.colors.accent1,
			]);
			expect(masterThemes()).toStrictEqual([theme1.path, theme2.path]);
			expect(live().theme?.colorScheme?.accent1?.replace('#', '')).toBe(theme1.colors.accent1);
			expect(live().history.canUndo).toBeFalsy();
			// Undo went back to the old part: neither part was rewritten.
			await expect(saved()).resolves.toStrictEqual(LOADED);

			await act(async () => {
				live().history.handleRedo();
			});
			await flush();
			expect(masterThemes()).toStrictEqual([theme2.path, theme2.path]);
			await expect(saved()).resolves.toStrictEqual({
				...LOADED,
				masters: ['theme2.xml', 'theme2.xml'],
			});
		});
	});

	it('switching the theme and applying the first one to every master re-theme in place too', async () => {
		await withHarness(await buildTwoThemesDeck(), async () => {
			const baseline = await openOnSlideTwo();
			// Every master on the first master's part: only slide 3's master moves.
			await act(async () => {
				await live().themeHandlers.handleApplyThemeToPresentation();
			});
			await flush();
			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			expect(accentFills()).toStrictEqual(Array(3).fill(theme1.colors.accent1));
			expect(masterThemes()).toStrictEqual([theme1.path, theme1.path]);

			await act(async () => {
				await live().themeHandlers.handleSwitchTheme(theme2.path);
			});
			await flush();
			expect(session(baseline)).toStrictEqual(KEPT_SESSION);
			expect(accentFills()).toStrictEqual(Array(3).fill(theme2.colors.accent1));
			expect(masterThemes()).toStrictEqual([theme2.path, theme2.path]);

			await act(async () => {
				live().history.handleUndo();
			});
			await flush();
			expect(accentFills()).toStrictEqual(Array(3).fill(theme1.colors.accent1));
			await expect(saved()).resolves.toStrictEqual({
				...LOADED,
				masters: ['theme1.xml', 'theme1.xml'],
			});
		});
	});

	it('parts that differ in their script faces: Undo writes neither part', async () => {
		await withHarness(await buildTwoThemesDeck({ scriptFaces: true }), async () => {
			await openOnSlideTwo();
			await act(async () => {
				await live().themeHandlers.handleApplyTheme(theme2.path, false);
			});
			await flush();
			await act(async () => {
				live().history.handleUndo();
			});
			await flush();
			await expect(themePartHeadingFaces(await live().save())).resolves.toStrictEqual({
				[theme1.path]: theme1.major,
				[theme2.path]: theme2.major,
			});
			await expect(saved()).resolves.toStrictEqual(LOADED);
		});
	});

	it("the moved master's style references take theme 2's format scheme, and Undo takes it back", async () => {
		await withHarness(await buildTwoThemesDeck({ formatSchemes: true }), async () => {
			await openOnSlideTwo();
			const shadow = () =>
				live().slides[0]?.elements.find((element) => element.name === TWO_THEMES.styledShapeName)
					?.shapeStyle?.shadowColor;
			expect(shadow()).toBeUndefined();
			await act(async () => {
				await live().themeHandlers.handleApplyTheme(theme2.path, false);
			});
			await flush();
			expect(shadow()).toBe('#000000');
			expect(live().theme?.formatScheme?.name).toBe('Ember');
			await act(async () => {
				live().history.handleUndo();
			});
			await flush();
			expect(shadow()).toBeUndefined();
			expect(live().theme?.formatScheme?.name).toBe('Tide');
		});
	});
});
