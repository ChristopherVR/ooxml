import { PptxHandler } from 'ooxml-core/pptx';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import {
	buildTwoThemesDeck,
	TWO_THEMES,
} from '../../../../core/pptx/__tests__/helpers/two-themes-deck';
import { partitionTemplateElements } from './template-editing';
import { applyThemeByPath } from './theme-by-path';
import type { ThemeByPathHandler } from './theme-by-path';
import { historyTheme, settleThemeWrites, writeThemeParts } from './theme-history';
import type { ThemePartSchemes } from './theme-history';
import { queueThemeWrite } from './theme-write-queue';

const { theme1, theme2, shapeName, masterBarName } = TWO_THEMES;

/** The two-themes deck as a binding holds it: template elements apart from the slides. */
async function open(options?: Parameters<typeof buildTwoThemesDeck>[0]) {
	const bytes = await buildTwoThemesDeck(options);
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer.slice(0) as ArrayBuffer);
	const { slides, templateElementsBySlideId } = partitionTemplateElements(data.slides);
	return {
		handler,
		slides,
		templateElementsBySlideId,
		slideMasters: data.slideMasters ?? [],
		theme: data.theme,
	};
}

const hex = (value: string | undefined) => value?.replace('#', '').toUpperCase();
const named = (elements: readonly PptxElement[] | undefined, name: string) =>
	elements?.find((element) => element.name === name);
const fillOf = (elements: readonly PptxElement[] | undefined, name: string) =>
	hex(named(elements, name)?.shapeStyle?.fillColor);
const headingFace = (slide: PptxSlide | undefined) => {
	const heading = named(slide?.elements, 'Heading');
	return heading && 'textSegments' in heading
		? heading.textSegments?.[0]?.style?.fontFamily
		: undefined;
};

/** The heading face of every slide after a save and a reload. */
async function savedHeadingFaces(handler: PptxHandler, slides: PptxSlide[]) {
	const reopened = new PptxHandler();
	const bytes = await handler.save(slides);
	const data = await reopened.load(bytes.buffer.slice(0) as ArrayBuffer);
	return data.slides.map(headingFace);
}

async function savedMasterThemes(handler: PptxHandler, slides: PptxSlide[]) {
	const reopened = new PptxHandler();
	const bytes = await handler.save(slides);
	const data = await reopened.load(bytes.buffer.slice(0) as ArrayBuffer);
	return data.slideMasters?.map((master) => master.themePath);
}

describe('applyThemeByPath', () => {
	it("every master: re-themes the moved master's slides from its old part, and only them", async () => {
		const deck = await open();
		expect(fillOf(deck.slides[1]?.elements, shapeName)).toBe(theme1.colors.accent1);
		const change = await applyThemeByPath(deck.handler, {
			themePath: theme2.path,
			applyToAllMasters: true,
			...deck,
		});
		expect(change).toBeDefined();
		for (const index of [0, 1]) {
			const slide = change!.slides[index];
			expect(fillOf(slide?.elements, shapeName)).toBe(theme2.colors.accent1);
			expect(headingFace(slide)).toBe(theme2.major);
			expect(fillOf(change!.templateElementsBySlideId[slide!.id], masterBarName)).toBe(
				theme2.colors.accent1,
			);
		}
		// Slide 3's master was on the new part already.
		const third = deck.slides[2]!;
		expect(change!.slides[2]).toBe(third);
		expect(change!.templateElementsBySlideId[third.id]).toBe(
			deck.templateElementsBySlideId[third.id],
		);
		expect(change!.slideMasters.map((master) => master.themePath)).toStrictEqual([
			theme2.path,
			theme2.path,
		]);
		expect(change!.slideMasters[1]).toBe(deck.slideMasters[1]);
		expect(change!.deckTheme?.theme.name).toBe('Ember');
		expect(hex(change!.deckTheme?.theme.colorScheme?.accent1)).toBe(theme2.colors.accent1);
		expect(change!.deckTheme?.theme.fontScheme?.majorFont?.latin).toBe(theme2.major);
		expect(hex(change!.deckTheme?.themeColorMap.accent1)).toBe(theme2.colors.accent1);
		await expect(savedMasterThemes(deck.handler, change!.slides)).resolves.toStrictEqual([
			theme2.path,
			theme2.path,
		]);
	});

	it("draws the moved master's style references with the new part's format scheme", async () => {
		const deck = await open({ formatSchemes: true });
		const styled = (slide: PptxSlide | undefined) =>
			named(slide?.elements, TWO_THEMES.styledShapeName)?.shapeStyle;
		expect(styled(deck.slides[0])?.shadowColor).toBeUndefined();
		expect(styled(deck.slides[2])?.shadowColor).toBeDefined();
		const change = await applyThemeByPath(deck.handler, {
			themePath: theme2.path,
			applyToAllMasters: false,
			...deck,
		});
		// Theme 2's first line style is 3 pt (4 px) and its first effect style an outer shadow.
		for (const index of [0, 1]) {
			expect(styled(change!.slides[index])?.strokeWidth).toBe(styled(deck.slides[2])?.strokeWidth);
			expect(styled(change!.slides[index])?.shadowColor).toBe(styled(deck.slides[2])?.shadowColor);
			expect(hex(styled(change!.slides[index])?.fillColor)).toBe(theme2.colors.accent1);
		}
		expect(change!.slides[2]).toBe(deck.slides[2]);
		expect(change!.deckTheme?.theme.formatScheme?.name).toBe('Ember');
	});

	it("every master, to the first master's part: only the second master's slide changes", async () => {
		const deck = await open();
		const change = await applyThemeByPath(deck.handler, {
			themePath: theme1.path,
			applyToAllMasters: true,
			...deck,
		});
		expect(change!.slides[0]).toBe(deck.slides[0]);
		expect(change!.slides[1]).toBe(deck.slides[1]);
		expect(fillOf(change!.slides[2]?.elements, shapeName)).toBe(theme1.colors.accent1);
		expect(headingFace(change!.slides[2])).toBe(theme1.major);
		expect(change!.deckTheme).toBeUndefined();
		expect(change!.slideMasters.map((master) => master.themePath)).toStrictEqual([
			theme1.path,
			theme1.path,
		]);
	});

	it('does nothing for a part the masters use already, or one the package lacks', async () => {
		const deck = await open();
		await expect(
			applyThemeByPath(deck.handler, { themePath: theme1.path, applyToAllMasters: false, ...deck }),
		).resolves.toBeUndefined();
		await expect(
			applyThemeByPath(deck.handler, {
				themePath: 'ppt/theme/theme9.xml',
				applyToAllMasters: true,
				...deck,
			}),
		).resolves.toBeUndefined();
		await expect(
			applyThemeByPath(null, { themePath: theme2.path, applyToAllMasters: true, ...deck }),
		).resolves.toBeUndefined();
		await expect(savedMasterThemes(deck.handler, deck.slides)).resolves.toStrictEqual([
			theme1.path,
			theme2.path,
		]);
	});

	it('runs after the theme writes queued before it, and a save waits for it', async () => {
		const log: string[] = [];
		const part = (path: string): ThemePartSchemes => ({
			path,
			colorScheme: { accent1: '#111111' } as ThemePartSchemes['colorScheme'],
			fontScheme: {},
		});
		const handler: ThemeByPathHandler = {
			getThemePartSchemes: async (path) => {
				log.push(`read ${path}`);
				return part(path);
			},
			setSlideMasterTheme: async (master, path) => {
				log.push(`${master} -> ${path}`);
			},
			reResolveThemeStyleReferences: (slides, templateElementsBySlideId) => ({
				slides: [...slides],
				templateElementsBySlideId: { ...templateElementsBySlideId },
			}),
		};
		void queueThemeWrite(handler, async () => {
			await new Promise((done) => setTimeout(done, 20));
			log.push('pick written');
		});
		const switched = applyThemeByPath(handler, {
			themePath: 'ppt/theme/theme2.xml',
			applyToAllMasters: false,
			slides: [],
			slideMasters: [{ path: 'm1', themePath: 'ppt/theme/theme1.xml' }, { path: 'm2' }],
			theme: undefined,
		});
		await settleThemeWrites(handler);
		expect(log).toStrictEqual([
			'pick written',
			'read ppt/theme/theme2.xml',
			'read ppt/theme/theme1.xml',
			'm1 -> ppt/theme/theme2.xml',
			// The deck theme is read from the part once the deck is on it.
			'read ppt/theme/theme2.xml',
		]);
		const change = await switched;
		expect(change?.slideMasters).toStrictEqual([
			{ path: 'm1', themePath: 'ppt/theme/theme2.xml' },
			{ path: 'm2' },
		]);
		expect(change?.deckTheme?.theme).toStrictEqual({
			colorScheme: { accent1: '#111111' },
			fontScheme: {},
		});
	});

	describe('Undo of a switch', () => {
		/** Switch the first master to theme 2, then undo it the way every binding does. */
		async function switchAndUndo(deck: Awaited<ReturnType<typeof open>>) {
			const before = historyTheme(deck.theme, deck.slideMasters);
			const change = await applyThemeByPath(deck.handler, {
				themePath: theme2.path,
				applyToAllMasters: false,
				...deck,
			});
			const after = historyTheme(change!.deckTheme!.theme, change!.slideMasters);
			const color = vi.spyOn(deck.handler, 'updateThemeColorScheme');
			const font = vi.spyOn(deck.handler, 'updateThemeFontScheme');
			await writeThemeParts(deck.handler, after, before);
			await settleThemeWrites(deck.handler);
			return { color, font };
		}

		it('writes nothing when the parts differ in their per-script faces', async () => {
			// The deck theme merges both parts' script faces, so it matches neither part's own.
			const deck = await open({ scriptFaces: true });
			expect(deck.theme?.fontScheme?.majorFont?.byScript).toStrictEqual({
				Jpan: 'Yu Mincho',
				Hang: 'Batang',
			});
			const { color, font } = await switchAndUndo(deck);
			const second = await deck.handler.getThemePartSchemes(theme2.path);
			expect(second?.fontScheme.majorFont?.latin).toBe(theme2.major);
			await expect(savedHeadingFaces(deck.handler, deck.slides)).resolves.toStrictEqual([
				theme1.major,
				theme1.major,
				theme2.major,
			]);
			expect(font).not.toHaveBeenCalled();
			expect(color).not.toHaveBeenCalled();
		});

		it("writes what the deck's part no longer holds into that part only", async () => {
			const deck = await open();
			const before = historyTheme(deck.theme, deck.slideMasters);
			// An edit the record predates reaches every part the masters use.
			const edited = { ...deck.theme!.colorScheme!, accent1: '#123456' };
			await deck.handler.updateThemeColorScheme(edited);
			const change = await applyThemeByPath(deck.handler, {
				themePath: theme2.path,
				applyToAllMasters: false,
				...deck,
			});
			const after = historyTheme(change!.deckTheme!.theme, change!.slideMasters);
			// The part's own schemes: a format scheme read resolves its placeholder colour against
			// whichever theme is active.
			const schemesOf = async (path: string) => {
				const part = await deck.handler.getThemePartSchemes(path);
				return { name: part?.name, colorScheme: part?.colorScheme, fontScheme: part?.fontScheme };
			};
			const held = await schemesOf(theme2.path);
			expect(hex(held.colorScheme?.accent1)).toBe('123456');
			// A step from after the switch back to before the edit.
			await writeThemeParts(deck.handler, after, before);
			await settleThemeWrites(deck.handler);
			const first = await deck.handler.getThemePartSchemes(theme1.path);
			expect(hex(first?.colorScheme.accent1)).toBe(theme1.colors.accent1);
			// The record describes the deck's part, not the part master 2 kept.
			await expect(schemesOf(theme2.path)).resolves.toStrictEqual(held);
		});
	});
});
